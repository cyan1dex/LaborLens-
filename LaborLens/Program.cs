using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.Collections.Specialized;
using System.Reflection;
using System.IO;
using System.Xml.Linq;


namespace LaborLens {
    class Program {

      public static string payrollFilename = "paydata2.txt";
      public static string timecardFilename = "timedata.txt";
      public static string project = "Air";

      static void Main(string[] args)
      {
         string currentDir = AppDomain.CurrentDomain.BaseDirectory;

         if (!currentDir.ToUpper().Contains(project.ToUpper())) {
            throw new Exception("DB in use is not correct");
         }

         #region timecard importer
         //var connString = @"Data Source=CODICI;User ID=codici;Password=agppci22;Initial Catalog=staging;Encrypt=False";
         //var importer = new TimecardImporter(connString);

         //string projectDir = Path.GetFullPath(Path.Combine(Directory.GetCurrentDirectory(), "..", "..", "..", "..", ".."));
         ////get all the paths of files with the name time in a directory
         //var paths = Directory.EnumerateFiles(projectDir, "*.*", SearchOption.TopDirectoryOnly)
         //.Where(p =>
         //{
         //   var name = Path.GetFileName(p);
         //   if (name.StartsWith("~$", StringComparison.OrdinalIgnoreCase)) return false; // temp
         //   var ext = Path.GetExtension(name).ToLowerInvariant();
         //   if (ext != ".xlsx" && ext != ".xls") return false;                           // no CSV, no code files
         //   return name.IndexOf("time", StringComparison.OrdinalIgnoreCase) >= 0;        // must contain "time"
         //})
         //.ToList();

         //foreach (var path in paths) {
         //   importer.ImportExcel(path, project);
         //}
         #endregion

         #region one time ingestion of Timedata from PDFs if needed
         //Ingest TNA1
         //var clocks = new Parser().ParseSingleLinePDF(Program.workComputer + Program.timecardFilename);
         //SQL.SQLRepository repo = new SQL.SQLRepository();
         //foreach (var card in clocks) {
         //   repo.AddTimecards(dbName, card.identifier, card.shifDate, card.clockIn, card.ClockOut, card.lp == true ? 1 : 0, card.hrsWorked, card.paycode == null ? String.Empty : card.paycode, card.breakTotal);
         //}
         #endregion

         #region SQL Data Parser - Timecards
        // var timecards = new SQL.SQLRepository().GetStagingTimecards(project);
        // var empCards = new SQL.SQLRepository().ConvertDataToDict(timecards);

          var empCards = new Dictionary<string, List<Timecard>>();
           var timecards2 = new SQL.SQLRepository().GetTimecards2(project);
            empCards = new SQL.SQLRepository().ConvertDataToDict(empCards, timecards2);

        // new ExcelWriter().WriteTimecardsFlat(empCards);
         #endregion
         #region write Data
         //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\CODICI\Desktop\output.txt", true)) {
         //   foreach (KeyValuePair<string, List<Timecard>> entry in empCards)
         //      foreach (Timecard s in entry.Value)
         //         if (s.timepunches.Count == 4 && s.totalHrsActual.TotalHours > 8) {

         //            double diff = s.totalHrsActual.TotalHours - s.regHrsListed;
         //            file.WriteLine(String.Format("{0}^{1}", diff > 0 ? diff : 0, diff < 0 ? Math.Abs(diff) : 0));

         //         }
         //}
         #endregion
         #region SQL Pay Parser
         var stubdata = new SQL.SQLRepository().GetPaydata(project);
          var stubs = new SQL.SQLRepository().ConvertPayDataToDict(stubdata);

          //Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         // var stubdata2 = new SQL.SQLRepository().GetPaydata2(dbName);
         // stubs = new SQL.SQLRepository().ConvertPayDataToDict2(stubs, stubdata2);

         #region PayStubs From Text
         //Dictionary<string, List<PayStub>> stubs = new Dictionary<string, List<PayStub>>();
         //var stubs = new PayStubParser().GetStubsFromText(@ConfigurationManager.AppSettings.Get("userDirectory") + "paydata1.txt");
         // stubs = new PayStubParser().GetStubsFromText(stubs, @ConfigurationManager.AppSettings.Get("userDirectory") + Program.payrollFilename);
         #endregion
         #endregion
         #region Rest Breaks Analysis
        var analyzer = new RestPeriodAnalyzer();
         var results = analyzer.AnalyzeEmployeeRestPeriods(empCards);
         var summary = analyzer.GenerateRestPeriodSummary(results);
         #endregion
         #region ADP PDF to CSV Parser
         //ADPParser adpParser = new ADPParser();
         //var adpCards = adpParser.ProcessTimecards(@"C:\Users\CYAN1\OneDrive\Desktop\Law Cases\JGill\Gonzalez v. Eminent\pay\paydata.csv");
         //var adpStubs = adpParser.ProcessPayStubs(@"C:\Users\CYAN1\OneDrive\Desktop\Law Cases\JGill\Gonzalez v. Eminent\pay\paydata.csv");
         #endregion
         #region Roster Analysis
         var dateProcessor = new DataProcessor();
         var rosterResults = dateProcessor.ProcessDateRanges(empCards, stubs);
         #endregion
         #region Create Shift analysis calculations
         var shifts = dateProcessor.GetAnalyzedShifts(empCards);

         Analysis analysis = new Analysis(shifts, stubs);
         analysis.totalEmployeesTimedata = empCards.Count;   
         
         int over35 = Shift.over35;
         int totalShifts = 0;
         double totalLength = 0;

         int c = 0;
         int meal60 = 0;

         foreach (KeyValuePair<string, List<Timecard>> entry in empCards) {
            c += entry.Value.Where(x => x.totalHrsActual.TotalHours == 8).Count();
            meal60 += entry.Value.Where(x => x.mealIs60).Count();

            foreach (Timecard t in entry.Value) {
               totalShifts++;
               totalLength += t.totalHrsActual.TotalHours;
            }
         }
         analysis.avgShiftlength = totalLength / totalShifts;
         #endregion

         #region count split shifts
         int split1hr = 0; //shifts separated by 1 hour or less
         int split2hr = 0; //shifts separated by 2 hour or less

         foreach (KeyValuePair<string, List<Timecard>> entry in empCards) {

            for (int pos = 0; pos < entry.Value.Count - 1; pos++) {
               double timeBtwnShfts = entry.Value[pos + 1].timepunches[0].datetime.Subtract(entry.Value[pos].timepunches[entry.Value[pos].timepunches.Count - 1].datetime).TotalHours;
               if (timeBtwnShfts <= 2 && timeBtwnShfts > 0)
                  split2hr++;
               if (timeBtwnShfts <= 1 && timeBtwnShfts > 0)
                  split1hr++;
            }
         }
         #endregion

         #region PAGA ANALYSIS (timesheets without stubs)
         ///////////Create timesheets from payperiod dates, only if there are timecards available (not stubs needed --> good in case they were not provided)
         Paga p = new Paga(Globals.pagaInitDt, Globals.pagaEndDt); //Create 1st to 15th, then remaining part of the month
        // Paga p = new Paga(Globals.pagaInitDt, Globals.pagaEndDt, 14); //Use this one to calculate at 14 or 7 days apart

         //////These timesheets do not have the paystubs, since some could be missing//////////////////////////
         Dictionary<string, List<Timesheet>> timesheetsWithoutStubs = p.PopulatePeriods(empCards);

      //    new ExcelWriter().WritePagaViolations(timesheetsWithoutStubs, p.periods);

         ////////Create PAGA penalty for the months//////////////////////
         var pagaData = new PayPeriods();
         pagaData.CalculatePagaPeriods(timesheetsWithoutStubs, Globals.pagaInitDt); //Use the PAGA start date, populate using Sheets w/o stubs

         analysis.totalWorkweeks = analysis.PeriodAnalysis(timesheetsWithoutStubs);
         #endregion

         #region Populate Timesheets
         var timeSheets = new PayPeriods().PopulateADPTimesheets(stubs, empCards); //MUST DO FOR PAY PERIODS    
         analysis.PeriodAnalysis(timeSheets); //Get Period analysis
          new PayPeriods().CalculatePagaPeriods(timeSheets, Globals.pagaInitDt); //Use the PAGA start date, populate using Sheets w/o stubs

         #endregion

         analysis.UnpaidOvertime(timeSheets);

         #region Rounding


         // new ExcelWriter().WriteRoundingActualVsListed(timesheetsWithoutStubs);
         #endregion

         #region AutoDeductions
         AutoDeductions deductions = new AutoDeductions();
         var autoDeductHhrs = deductions.AutoDeductHours(timeSheets);
         #endregion

         #region Create Word Analysis Output
         analysis.CompleteCalcuations();
         // new DocWriter().WriteDocument(analysis);
         #endregion

         #region Create Analysis Spreadsheet
         analysis.meal30 = Shift.mealIs30;
         analysis.mealsTaken = Shift.totalMeals;
         analysis.shift8 = Shift.shiftIs8;

         new ExcelWriter().PoulateGraphData(shifts, empCards.Count, pagaData, over35, analysis);
         #endregion

         #region Salary analysis
        //  new ExcelWriter().PoulateRoster(rosterResults);
       //   new ExcelWriter().PoulateSummaryPayData(stubs, analysis); //Write pay data by year
        //   new ExcelWriter().WritePayDetails(stubs); //Write pay data by employee

         //double totalHrs = 0;
         //double cnt = 0;

         //using (System.IO.StreamWriter file = new System.IO.StreamWriter(output, true)) {
         //   foreach (KeyValuePair<string, List<PayStub>> entry in stubs) {
         //      // c += entry.Value.Where(x => x.totalHrsActual.TotalHours < 2).Count();
         //      DateTime lastPaycheck = DateTime.MinValue;

         //      foreach (PayStub t in entry.Value) {
         //         if (!t.periodBegin.HasValue)
         //            continue;
         //         //  cnt++;
         //         if (t.periodBegin.Value > lastPaycheck)
         //            lastPaycheck = t.periodEnd.Value;
         //      }

         //      file.WriteLine(String.Format("{0} | {1}", entry.Key, lastPaycheck.ToShortDateString()));
         //   }
         //}
         #endregion

        // new ExcelWriter().ExportOvertimeAuditTsv(timeSheets, project);

         new ExcelWriter().WriteTimesheetViolations(timeSheets);
         #region Debug overtime of employer vs actual

         string targetEmp = "2074";
         var suspectPeriods = new HashSet<DateTime>
         {
          new DateTime(2021, 12, 26),
             new DateTime(2022, 02, 06),
             new DateTime(2022, 07, 10)
         };

         List<Timesheet> empSheets = timeSheets.TryGetValue(targetEmp, out var listForEmp)
             ? listForEmp
             : new List<Timesheet>();

         var debugSheets = empSheets
             .Where(ts => ts?.stub != null && ts.stub.periodEnd.HasValue &&
                          suspectPeriods.Contains(ts.stub.periodEnd.Value.Date))
             .OrderBy(ts => ts.stub.periodEnd.Value)
             .ToList();

         foreach (var ts in debugSheets) {
            Console.WriteLine($"\n=== DEBUG TRACE {ts.identifier}  PB={ts.periodBegin:MM/dd/yyyy}  PE={ts.periodEnd:MM/dd/yyyy} ===");

            // Recalculate fully
            ts.AnalyzeADPHours();

            // group by week (same as CalculateOvertime does)
            var groupedWeeks = ts.timeCards
               .OrderBy(tc => tc.shiftDate!.Value.Date)
               .GroupBy(tc =>
               {
                  var d = tc.shiftDate!.Value.Date;                 // strip time
                  var sunday = d.AddDays(-(int)d.DayOfWeek).Date;   // anchor to Sunday
                  return sunday;
               });

            foreach (var wk in groupedWeeks) {
               Console.WriteLine($"  WEEK {wk.Key:MM/dd}–{wk.Key.AddDays(6):MM/dd}");

               double weekTotal = 0, weekOT = 0, weekDT = 0, weekReg = 0;
               foreach (var card in wk) {
                  double hrs = card.totalHrsActual.TotalHours;
                  double reg = Math.Min(8, hrs);
                  double ot = 0;
                  double dt = 0;

                  if (hrs > 8 && hrs <= 12)
                     ot = hrs - 8;
                  else if (hrs > 12) {
                     ot = 4;
                     dt = hrs - 12;
                  }

                  weekTotal += hrs;
                  weekOT += ot;
                  weekDT += dt;
                  weekReg += reg;

                  Console.WriteLine($"    {card.shiftDate:MM/dd}  Hrs={hrs,5:F2}  Reg={reg,5:F2}  OT={ot,5:F2}  DT={dt,5:F2}");
               }

               Console.WriteLine($"    --- WEEKLY TOTALS ---  REG={weekReg:F2}  OT={weekOT:F2}  DT={weekDT:F2}  SUM={weekTotal:F2}");
               Console.WriteLine();
            }

            Console.WriteLine($"PAY PERIOD SUMMARY  Actual={ts.actualHours.TotalHours + ts.actualOT.TotalHours + ts.actualDblOT.TotalHours:F2}  " +
                              $"REG={ts.actualHours.TotalHours:F2}  OT={ts.actualOT.TotalHours:F2}  DT={ts.actualDblOT.TotalHours:F2}");
            Console.WriteLine($"CheckStub Summary   REG={ts.stub.regHrs:F2}  OT={ts.stub.otHrs:F2}  DT={ts.stub.doubleOtHrs:F2}  TOTAL={ts.stub.regHrs + ts.stub.otHrs + ts.stub.doubleOtHrs:F2}");
         }
         #endregion
      }



      public static int GetTotalWorkweeks(Dictionary<string, List<Timesheet>> empSheets)
      {
         int total = 0;
         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            foreach (Timesheet s in entry.Value) {
               total += s.GetWorkWeeks();
            }
         }
         return total;
      }

      public static Dictionary<DateTime, int> GetPayPeriodTotals(Dictionary<string, List<Timesheet>> empSheets)
      {
         Dictionary<DateTime, int> dict = new Dictionary<DateTime, int>();

         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            foreach (Timesheet s in entry.Value) {
               if (s.periodBegin != null) {

                  if (!dict.ContainsKey(s.periodBegin.Value))
                     dict[s.periodBegin.Value] = 0;

                  dict[s.periodBegin.Value]++;
               }
            }
         }

         return dict;
      }
   }
}

