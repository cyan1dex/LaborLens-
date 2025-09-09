// TimecardImporter.cs  (C# 7.3 compatible)
using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Microsoft.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using ExcelDataReader;

namespace LaborLens {
   public class TimecardImporter {
      private readonly string _connectionString;

      public TimecardImporter(string connectionString)
      {
         _connectionString = connectionString;

         // Needed on .NET Core/5+ so ExcelDataReader can parse legacy encodings.
         try {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
         } catch {
            // On .NET Framework this provider already exists; ignore.
         }
      }

      public void ImportExcel(string filePath, string projectKey)
      {
         var tvp = MakeTvpTable();
         var batchId = Guid.NewGuid();
         var fileName = Path.GetFileName(filePath);

         using (var fs = File.Open(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
         using (var reader = ExcelReaderFactory.CreateReader(fs)) {
            do {
               if (!reader.Read()) continue; // header row
               var headers = ReadHeaders(reader);

               int colEmp = Find(headers, new[] { "id", "eid", "filenumber", "employeeid", "emp","name" ,"positionid"}, true);
               int colDate = Find(headers, new[] { "shiftdate", "date", "workdate", "shift_date" }, false);

               int colHours = Find(headers, new[] { "total", "reg_hrs", "regularhours", "hours" ,"ListedHrs","actualhours"}, false);

               var inCols = FindIndexed(headers, new[] { "timein", "in", "clockin", "start" ,"time_in", "ActualPunchInTime" });
               var outCols = FindIndexed(headers, new[] { "timeout", "out", "clockout", "end","time_out", "ActualPunchOutTime" });

               int rowNum = 1; // after header
               while (reader.Read()) {
                  rowNum++;
                  if (RowIsEmpty(reader)) continue;

                  var rawEmp = GetString(reader, colEmp);
                  if (string.IsNullOrWhiteSpace(rawEmp)) continue;
                  var eid = rawEmp.Trim();

                  DateTime? shiftDateFromSheet = TryDate(GetObj(reader, colDate));

                  // ensure stable ordering of pairs
                  foreach (var k in inCols.Keys.OrderBy(i => i)) {
                     int outColIndex;
                     if (!outCols.TryGetValue(k, out outColIndex)) continue;

                     DateTime tin, tout;
                     if (!TryTime(GetObj(reader, inCols[k]), out tin)) continue;
                     if (!TryTime(GetObj(reader, outColIndex), out tout)) continue;

                     var shiftDate = shiftDateFromSheet.HasValue ? shiftDateFromSheet.Value : tin.Date;
                     var notes = (tout < tin) ? "Cross-midnight adjusted" : null;

                     decimal parsedHours;
                     decimal? reg = null;
                     if (colHours >= 0 && TryHours(reader.GetValue(colHours), out parsedHours)) {
                        reg = parsedHours; // keep the source’s rounding/precision
                     }


                     var r = tvp.NewRow();
                     r["ProjectKey"] = projectKey;
                     r["BatchId"] = batchId;
                     r["FileName"] = fileName;
                     r["RowNum"] = rowNum;
                     r["PairIndex"] = k;
                     r["EmployeeId"] = eid;
                     r["ShiftDate"] = shiftDate.Date;
                     r["InTime"] = tin.TimeOfDay;
                     r["OutTime"] = tout.TimeOfDay;
                     r["RegHours"] = (object)(reg.HasValue ? reg.Value : (object)DBNull.Value);
                     r["Notes"] = (object)(notes ?? (object)DBNull.Value);
                     tvp.Rows.Add(r);
                  }
               }

            } while (reader.NextResult()); // next worksheet
         }

         // push to SQL
         BulkInsert(tvp);
      }

      // ---------- SQL ----------
      private void BulkInsert(DataTable tvp)
      {
         using (var cn = new SqlConnection(_connectionString)) {
            cn.Open();
            using (var cmd = new SqlCommand("dbo.Timecards_Ingest", cn)) {
               cmd.CommandType = CommandType.StoredProcedure;
               var p = cmd.Parameters.AddWithValue("@Rows", tvp);
               p.SqlDbType = SqlDbType.Structured;
               p.TypeName = "dbo.TimecardSegment_TVP"; // must match SQL type
               cmd.ExecuteNonQuery();
            }
         }
      }

      // ---------- Helpers ----------
      private static DataTable MakeTvpTable()
      {
         var dt = new DataTable();
         dt.Columns.Add("ProjectKey", typeof(string));
         dt.Columns.Add("BatchId", typeof(Guid));
         dt.Columns.Add("FileName", typeof(string));
         dt.Columns.Add("RowNum", typeof(int));
         dt.Columns.Add("PairIndex", typeof(int));
         dt.Columns.Add("EmployeeId", typeof(string));
         dt.Columns.Add("ShiftDate", typeof(DateTime));
         dt.Columns.Add("InTime", typeof(TimeSpan));
         dt.Columns.Add("OutTime", typeof(TimeSpan));
         dt.Columns.Add("RegHours", typeof(decimal));
         dt.Columns.Add("Notes", typeof(string));
         return dt;
      }

      private static string[] ReadHeaders(IExcelDataReader reader)
      {
         var headers = new List<string>();
         for (int i = 0; i < reader.FieldCount; i++) {
            var raw = reader.GetValue(i);
            headers.Add(Norm(raw == null ? "" : raw.ToString()));
         }
         return headers.ToArray();
      }

      private static string Norm(string s)
      {
         if (s == null) return "";
         s = s.ToLowerInvariant().Trim();
         return Regex.Replace(s, @"\s+|[^\w]", "");
      }

      private static int Find(string[] header, IEnumerable<string> synonyms, bool required)
      {
         var set = new HashSet<string>(synonyms.Select(Norm));
         for (int i = 0; i < header.Length; i++) {
            if (set.Contains(header[i])) return i;
         }
         if (required) throw new Exception("Required column not found: " + string.Join("/", synonyms));
         return -1;
      }


      // Finds columns like: In, Out, In1, Out1, In2, Out2 (or timein/timeout/clockin/clockout/start/end)
      // Rule: unnumbered = pair 1; numbered N = pair (N+1) if an unnumbered exists, else pair N
      private static Dictionary<int, int> FindIndexed(string[] header, IEnumerable<string> bases)
      {
         var baseSet = new HashSet<string>(bases.Select(Norm));
         var rx = new Regex(@"^(?<base>[a-z]+?)(?<idx>\d+)?$");
         var hits = new List<(int col, string b, int? idx)>();

         for (int i = 0; i < header.Length; i++) {
            var m = rx.Match(header[i]);
            if (!m.Success) continue;
            var b = m.Groups["base"].Value;
            if (!baseSet.Contains(b)) continue;

            int? idx = null;
            if (m.Groups["idx"].Success) {
               int parsed;
               if (int.TryParse(m.Groups["idx"].Value, out parsed)) idx = parsed;
            }
            hits.Add((i, b, idx));
         }

         // if any unnumbered exists (e.g., "in"), shift numbered ones by +1
         bool hasUnnumbered = hits.Any(h => h.idx == null);
         int offset = hasUnnumbered ? 1 : 0;

         var map = new Dictionary<int, int>();
         foreach (var h in hits.OrderBy(h => h.col))   // preserve left-to-right order
         {
            int k = h.idx.HasValue ? h.idx.Value + offset : 1;
            while (map.ContainsKey(k)) k++;           // avoid collision
            map[k] = h.col;
         }
         return map;
      }



      private static bool RowIsEmpty(IExcelDataReader reader)
      {
         for (int i = 0; i < reader.FieldCount; i++) {
            var v = reader.GetValue(i);
            if (v != null && !string.IsNullOrWhiteSpace(v.ToString())) return false;
         }
         return true;
      }

      private static object GetObj(IExcelDataReader reader, int col)
      {
         if (col < 0) return null;
         return reader.GetValue(col);
      }

      private static string GetString(IExcelDataReader reader, int col)
      {
         var o = GetObj(reader, col);
         return o == null ? null : o.ToString();
      }

      private static DateTime? TryDate(object v)
      {
         if (v == null) return null;
         var dt = v as DateTime?;
         if (dt.HasValue) return dt.Value.Date;

         DateTime parsed;
         if (DateTime.TryParse(v.ToString(), out parsed)) return parsed.Date;
         return null;
      }

      private static bool TryTime(object v, out DateTime t)
      {
         // Excel commonly gives DateTime (date + time) or a string time
         if (v is DateTime) {
            t = (DateTime)v;
            return true;
         }
         TimeSpan ts;
         if (v != null && TimeSpan.TryParse(v.ToString(), out ts)) {
            t = DateTime.Today.Add(ts);
            return true;
         }
         DateTime dt;
         if (v != null && DateTime.TryParse(v.ToString(), out dt)) {
            t = dt;
            return true;
         }
         t = default(DateTime);
         return false;
      }

      // Returns true if we parsed hours; supports "h:mm", "hh:mm:ss", DateTime/TimeSpan,
      // Excel numeric time (fraction of a day), or plain decimal hours.
      private static bool TryHours(object v, out decimal hours)
      {
         hours = 0m;
         if (v == null) return false;

         // If Excel gave you a DateTime or TimeSpan (common for time cells)
         if (v is DateTime dt) {
            hours = (decimal)dt.TimeOfDay.TotalHours;
            return true;
         }
         if (v is TimeSpan ts) {
            hours = (decimal)ts.TotalHours;
            return true;
         }

         // If Excel gave you a number (double/decimal)
         // - For true decimal hours (e.g., 9.25) value > 1
         // - For time-of-day/duration cells formatted as time, Excel stores fraction of day (<= 1)
         double d;
         if (v is double && (d = (double)v) >= 0) {
            hours = (decimal)(d <= 1.5 ? d * 24.0 : d); // <=1.5 => treat as fraction-of-day
            return true;
         }
         decimal dec;
         if (v is decimal && (dec = (decimal)v) >= 0) {
            hours = dec; // assume decimal hours
            return true;
         }

         // String cases: "9:12", "09:12:30", "9.2"
         var s = v.ToString().Trim();
         if (string.IsNullOrEmpty(s)) return false;

         // colon => parse as TimeSpan
         TimeSpan ts2;
         if (s.IndexOf(':') >= 0 && TimeSpan.TryParse(s, out ts2)) {
            hours = (decimal)ts2.TotalHours;
            return true;
         }

         // plain number string (decimal hours)
         decimal dh;
         if (decimal.TryParse(s, out dh)) {
            hours = dh;
            return true;
         }

         return false;
      }


      private static decimal? TryDecimal(object v)
      {
         if (v == null) return null;
         decimal d;
         if (decimal.TryParse(v.ToString(), out d)) return d;
         return null;
      }
   }
}
