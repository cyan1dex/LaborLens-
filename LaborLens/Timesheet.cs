using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LaborLens {
   public class OvertimeResult {
      public double RegularHours { get; set; }
      public double OvertimeHours { get; set; }
      public double DoubletimeHours { get; set; }
   }

   public class Timesheet {
      // ===== persisted / external =====
      public PayStub stub;
      public List<Timecard> timeCards = new List<Timecard>();
      public string identifier;

      // ===== period inputs (from stub) =====
      public DateTime? periodBegin;
      public DateTime? periodEnd;

      // ===== period-derived hour buckets =====
      public TimeSpan listHours;
      public TimeSpan listOT;

      public TimeSpan actualHours;
      public TimeSpan actualOT;
      public TimeSpan actualDblOT;

      public double listedTotalHours;
      public double actualTotalHours;
      public double checkTotalHours;

      public double regPay;
      public double otPay;
      public double regRate;
      public double otRate;

      public bool invalid;
      public bool missingPeriodDates;

      public static int totalWorkweeks; // kept as-is for broader aggregation

      // rounding diagnostics kept (unchanged signatures)
      public int roundedShiftsForCompany;
      public int roundedShiftsForEmployee;
      public Decimal roudingBalance;
      public int roundedWOrksWeeks;

      public int splitShiftOneHr;
      public int splitShiftTwohr;

      public HashSet<DateTime> holidays;
      public bool workedAllDaysOfaWeek;
      public bool sevenInArow;

      public static int sevenDaysCnt;

      // ===== Optional, silent CSV for select periods (OFF by default) =====
      public bool EnableCsvForSuspectPeriods { get; set; } = false;
      public HashSet<DateTime> SuspectPeriodEnds { get; set; } = new HashSet<DateTime>(); // add PE dates
      public string CsvOutputDirectory { get; set; } = AppDomain.CurrentDomain.BaseDirectory;

      // ===== Basic utilities =====
      public bool PaidPenalties()
      {
         foreach (var t in timeCards)
            if (t.penalties > 0) return true;
         return false;
      }
      public bool HasViolation()
      {
         foreach (var t in timeCards)
            if (t.HasViolation()) return true;
         return false;
      }
      public bool HasFirstMealViolation()
      {
         foreach (var t in timeCards)
            if (t.HasFirstMealViolation()) return true;
         return false;
      }
      public bool HasSecondMealViolation()
      {
         foreach (var t in timeCards)
            if (t.HasSecondMealViolation()) return true;
         return false;
      }

      public void PopulateTimeCardShifts()
      {
         if (!missingPeriodDates && periodBegin.HasValue)
            foreach (var t in timeCards) t.PopulateShiftDate(periodBegin);
      }

      public int GetWorkWeeks()
      {
         if (!periodBegin.HasValue) return 2;
         int daysInWeekOneWrkd = 0, daysInWeekTwoWrked = 0;
         var midweek = periodBegin.Value.AddDays(6);

         foreach (var card in timeCards) {
            if (card.shiftDate <= midweek) daysInWeekOneWrkd++;
            else daysInWeekTwoWrked++;
         }
         if (daysInWeekOneWrkd > 0 && daysInWeekTwoWrked > 0) return 2;
         if (daysInWeekOneWrkd > 0 || daysInWeekTwoWrked > 0) return 1;
         return 0;
      }

      public bool ProcessRounding()
      {
         foreach (var card in timeCards) {
            if (card.mealsTaken == 0 && card.totalHrsActual.TotalHours > 5) {
               double actHrs = card.totalHrsActual.TotalHours - 1;
               if (actHrs > 8) {
                  if (actHrs - card.regHrsListed > 0) roundedShiftsForCompany++;
                  else if (actHrs - card.regHrsListed < 0) roundedShiftsForEmployee++;
                  roudingBalance += (Decimal)(actHrs - card.regHrsListed);
               }
            }
         }
         if (timeCards.Count > 0 && timeCards[0].shiftDate < new DateTime(2016, 2, 6))
            roundedWOrksWeeks += 2;

         return roudingBalance > 0;
      }

      public void SetListedHoursFromTimecard()
      {
         listedTotalHours = 0;
         foreach (var card in timeCards)
            listedTotalHours += card.regHrsListed + card.otListed + card.dtListed;
      }

      // ===== Overtime engines =====

      // Public because other parts of your code call this with this signature.
      public OvertimeResult CalculateOvertime(List<Timecard> timecards, bool trace = false, string tag = "")
      {
         if (timecards == null || timecards.Count == 0) return new OvertimeResult();

         var total = new OvertimeResult();

         // group by workweek (Sun–Sat)
         var byWeek = timecards
            .OrderBy(t => t.shiftDate.Value.Date)
            .GroupBy(t => GetWorkweekStartDate(t.shiftDate.Value.Date));

         foreach (var wk in byWeek) {
            var r = CalculateWeeklyOvertime(wk.ToList(), trace, tag);
            total.RegularHours += r.RegularHours;
            total.OvertimeHours += r.OvertimeHours;
            total.DoubletimeHours += r.DoubletimeHours;
         }

         return total;
      }
      public void CalculateNonStub()
      {
         // Reset period buckets
         actualHours = TimeSpan.Zero;
         actualOT = TimeSpan.Zero;
         actualDblOT = TimeSpan.Zero;
         listedTotalHours = 0;

         if (timeCards == null || timeCards.Count == 0)
            return;

         // Keep your legacy consecutive-days flag behavior
         sevenInArow = false;
         var ordered = timeCards.OrderBy(tc => tc.shiftDate.Value.Date).ToList();
         int consec = 1;
         for (int i = 1; i < ordered.Count; i++) {
            var prev = ordered[i - 1].shiftDate.Value.Date;
            var cur = ordered[i].shiftDate.Value.Date;
            if ((cur - prev).Days == 1) consec++;
            else consec = 1;

            if (consec >= 7) { sevenInArow = true; break; }
         }

         // Total worked hours in the period
         double totalWorked = ordered.Sum(tc => tc.totalHrsActual.TotalHours);

         // Use the same CA rules engine as AnalyzeADPHours -> weekly (Sun–Sat) with anti-pyramiding + 7th day
         var byWeek = ordered
            .GroupBy(t => GetWorkweekStartDate(t.shiftDate.Value.Date))
            .OrderBy(g => g.Key);

         double reg = 0, ot = 0, dt = 0;

         foreach (var wk in byWeek) {
            var r = CalculateWeeklyOvertime(wk.ToList(), trace: false, tag: "");
            reg += r.RegularHours;
            ot += r.OvertimeHours;
            dt += r.DoubletimeHours;
         }

         actualOT = TimeSpan.FromHours(ot);
         actualDblOT = TimeSpan.FromHours(dt);
         actualHours = TimeSpan.FromHours(totalWorked) - actualOT - actualDblOT;

         // Maintain your listed totals behavior
         SetListedHoursFromTimecard();

         // (Optional) If you need reconciliation numbers updated here as well:
         // checkTotalHours   = stub.regHrs + stub.otHrs + stub.doubleOtHrs;
         // actualTotalHours  = actualHours.TotalHours + actualOT.TotalHours + actualDblOT.TotalHours;
      }

      // Weekly engine: daily OT/DT first, then weekly 40+ overlay (anti-pyramiding), then 7th-day (within week).
      private OvertimeResult CalculateWeeklyOvertime(List<Timecard> weekTimecards, bool trace = false, string tag = "")
      {
         var result = new OvertimeResult();

         // day totals
         var daily = weekTimecards
            .GroupBy(t => t.shiftDate.Value.Date)
            .OrderBy(g => g.Key)
            .Select(g => new { Day = g.Key, Hrs = g.Sum(t => t.totalHrsActual.TotalHours) })
            .ToList();

         double straight = 0;
         double dailyOT = 0;
         double dailyDT = 0;

         foreach (var d in daily) {
            if (d.Hrs <= 8.0) { straight += d.Hrs; } else if (d.Hrs <= 12.0) { straight += 8.0; dailyOT += (d.Hrs - 8.0); } else { straight += 8.0; dailyOT += 4.0; dailyDT += (d.Hrs - 12.0); }
         }

         // 7th consecutive day rule within THIS workweek
         var (seventhOT, seventhDT) = Apply7thDayOvertime(weekTimecards);
         // We add these premium hours; they’re NOT double-counted against daily buckets above (we didn’t isolate 7th-day separately there).
         dailyOT += seventhOT;
         dailyDT += seventhDT;

         // weekly OT applies to straight-time > 40
         double weeklyOT = Math.Max(0.0, straight - 40.0);

         // anti-pyramiding: final OT is the GREATER of dailyOT vs weeklyOT
         double finalOT = Math.Max(dailyOT, weeklyOT);

         // final regular hours come from straight minus any additional OT we needed to match weekly
         double finalReg = straight;
         if (weeklyOT > dailyOT) {
            double addl = weeklyOT - dailyOT;
            finalReg = Math.Max(0, straight - addl);
         }

         result.RegularHours = finalReg;
         result.OvertimeHours = finalOT;
         result.DoubletimeHours = dailyDT;
         return result;
      }

      // Semi-monthly engine: evaluate each Sunday–Saturday week inside the period; add one 7th-day sequence across the entire period if it exists.
      public (double regularHours, double overtimeHours, double doubleTimeHours) CalculateOTHoursBiMonthly(Timesheet timesheet)
      {
         double totalRegular = 0, totalOT = 0, totalDT = 0;

         // one 7th-consecutive overlay across the entire semi-monthly block
         var (global7thOT, global7thDT, seventhDayDate) = FindSeventhConsecutiveDay(timesheet.timeCards.ToList());

         // process by week
         var weeks = timesheet.timeCards
             .GroupBy(tc => GetWorkweekStartDate(tc.shiftDate.Value))
             .OrderBy(g => g.Key)
             .ToList();

         foreach (var week in weeks) {
            // day totals per week
            var daily = week
                .GroupBy(tc => tc.shiftDate.Value.Date)
                .Select(g => new { Day = g.Key, Hrs = g.Sum(tc => tc.totalHrsActual.TotalHours) })
                .OrderBy(x => x.Day)
                .ToList();

            double straight = 0, dailyOT = 0, dailyDT = 0;

            foreach (var d in daily) {
               // skip the 7th-day date if it belongs to the global overlay; we’ll add it once after loop
               if (seventhDayDate.HasValue && d.Day.Date == seventhDayDate.Value.Date) continue;

               if (d.Hrs <= 8.0) straight += d.Hrs;
               else if (d.Hrs <= 12.0) { straight += 8.0; dailyOT += (d.Hrs - 8.0); } else { straight += 8.0; dailyOT += 4.0; dailyDT += (d.Hrs - 12.0); }
            }

            double weeklyOT = Math.Max(0.0, straight - 40.0);
            double finalOT = Math.Max(dailyOT, weeklyOT);
            double finalReg = straight;
            if (weeklyOT > dailyOT) {
               double addl = weeklyOT - dailyOT;
               finalReg = Math.Max(0, straight - addl);
            }

            totalRegular += finalReg;
            totalOT += finalOT;
            totalDT += dailyDT;
         }

         // Add the single 7th-consecutive overlay for the period (if any)
         totalOT += global7thOT;
         totalDT += global7thDT;

         return (totalRegular, totalOT, totalDT);
      }

      // CA 7th consecutive day rule WITHIN a workweek (Sunday–Saturday)
      private (double ot, double dt) Apply7thDayOvertime(List<Timecard> weekCards)
      {
         var workDays = weekCards
            .Select(t => t.shiftDate.Value.Date)
            .Distinct()
            .OrderBy(d => d)
            .ToList();

         if (workDays.Count < 7) return (0.0, 0.0);

         for (int i = 0; i <= workDays.Count - 7; i++) {
            bool consecutive = true;
            for (int j = 0; j < 6; j++)
               if ((workDays[i + j + 1] - workDays[i + j]).Days != 1) { consecutive = false; break; }

            if (consecutive) {
               var seventhDay = workDays[i + 6];
               double hrs = weekCards.Where(t => t.shiftDate.Value.Date == seventhDay)
                                     .Sum(t => t.totalHrsActual.TotalHours);
               double ot = Math.Min(8.0, hrs);
               double dt = Math.Max(0.0, hrs - 8.0);
               return (ot, dt);
            }
         }
         return (0.0, 0.0);
      }

      // 7th-day overlay across an ENTIRE semi-monthly period
      private (double ot, double dt, DateTime? seventhDay) FindSeventhConsecutiveDay(List<Timecard> cards)
      {
         var workDays = cards
            .Select(t => t.shiftDate.Value.Date)
            .Distinct()
            .OrderBy(d => d)
            .ToList();

         if (workDays.Count < 7) return (0.0, 0.0, null);

         for (int i = 0; i <= workDays.Count - 7; i++) {
            bool consecutive = true;
            for (int j = 0; j < 6; j++)
               if ((workDays[i + j + 1] - workDays[i + j]).Days != 1) { consecutive = false; break; }

            if (consecutive) {
               var seventh = workDays[i + 6];
               double hrs = cards.Where(t => t.shiftDate.Value.Date == seventh).Sum(t => t.totalHrsActual.TotalHours);
               double ot = Math.Min(8.0, hrs);
               double dt = Math.Max(0.0, hrs - 8.0);
               return (ot, dt, seventh);
            }
         }
         return (0.0, 0.0, null);
      }

      private DateTime GetWorkweekStartDate(DateTime date)
      {
         int daysToSubtract = (int)date.DayOfWeek; // Sunday = 0
         return date.Date.AddDays(-daysToSubtract);
      }

      private bool IsBiWeekly(DateTime? begin, DateTime? end)
      {
         if (!begin.HasValue || !end.HasValue) return false;
         // inclusive span (guard for DST/export off-by-one)
         int days = (end.Value.Date - begin.Value.Date).Days + 1;
         return (days >= 14 && days <= 15);
      }

      // ===== Top-level analysis (quiet) =====
      public void AnalyzeADPHours()
      {
         // reset
         actualHours = TimeSpan.Zero;
         actualOT = TimeSpan.Zero;
         actualDblOT = TimeSpan.Zero;

         // map stub
         listHours = TimeSpan.FromHours(stub.regHrs);
         listOT = stub.otHrs != 0 ? TimeSpan.FromHours(stub.otHrs) : TimeSpan.Zero;

         regPay = stub.regPay;
         regRate = stub.regRate;
         otPay = stub.otPay;
         otRate = stub.otRate;

         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;

         // sanity + counters
         int daysInWeekOneWrkd = 0, daysInWeekTwoWrked = 0;
         DateTime midweek = periodBegin.HasValue ? periodBegin.Value.AddDays(6) : DateTime.MinValue;

         foreach (var card in timeCards) {
            if (card.totalHrsActual.TotalHours < 0) { this.invalid = true; card.invalid = true; }
            if (periodBegin.HasValue) {
               if (card.shiftDate <= midweek) daysInWeekOneWrkd++;
               else daysInWeekTwoWrked++;
            }
         }

         double totalHoursWorked = timeCards.Sum(tc => tc.totalHrsActual.TotalHours);

         // select engine
         OvertimeResult res;
         if (IsBiWeekly(periodBegin, periodEnd)) {
            res = CalculateOvertime(timeCards, trace: false, tag: "");
         } else {
            var v = CalculateOTHoursBiMonthly(this);
            res = new OvertimeResult { RegularHours = v.regularHours, OvertimeHours = v.overtimeHours, DoubletimeHours = v.doubleTimeHours };
         }

         actualOT = TimeSpan.FromHours(res.OvertimeHours);
         actualDblOT = TimeSpan.FromHours(res.DoubletimeHours);
         actualHours = TimeSpan.FromHours(totalHoursWorked) - actualOT - actualDblOT;

         // reconciliation
         checkTotalHours = stub.regHrs + stub.otHrs + stub.doubleOtHrs;
         actualTotalHours = actualHours.TotalHours + actualOT.TotalHours + actualDblOT.TotalHours;

         SetListedHoursFromTimecard();

         // workweek counters
         if (daysInWeekOneWrkd > 0) totalWorkweeks++;
         if (daysInWeekTwoWrked > 0) totalWorkweeks++;

         // OPTIONAL: silent CSV for suspect PEs (off by default)
         if (EnableCsvForSuspectPeriods && periodEnd.HasValue && SuspectPeriodEnds.Contains(periodEnd.Value.Date))
            WriteCompactOtCsvForPeriod();
      }

      // ===== Minimal CSV (silent) for spot-checking suspect periods =====
      private void WriteCompactOtCsvForPeriod()
      {
         var file = Path.Combine(
            string.IsNullOrWhiteSpace(CsvOutputDirectory) ? AppDomain.CurrentDomain.BaseDirectory : CsvOutputDirectory,
            $"OT_Debug_{identifier}_{periodEnd:yyyyMMdd}.csv"
         );

         var sb = new StringBuilder();
         if (!File.Exists(file))
            sb.AppendLine("EmpId,PeriodEnd,WeekStart,Date,ShiftHours,RegHrs,OTHrs,DTHrs,WeekStraight,WeekDailyOT,WeekDT,WeeklyOT,FinalReg,FinalOT,FinalDT");

         var byWeek = timeCards
            .OrderBy(t => t.shiftDate.Value)
            .GroupBy(t => GetWorkweekStartDate(t.shiftDate.Value));

         foreach (var wk in byWeek) {
            var daily = wk.GroupBy(t => t.shiftDate.Value.Date)
                          .OrderBy(g => g.Key)
                          .Select(g => new { Day = g.Key, Hrs = g.Sum(t => t.totalHrsActual.TotalHours) })
                          .ToList();

            double straight = 0, dailyOT = 0, dailyDT = 0;

            foreach (var d in daily) {
               double r = 0, o = 0, dt = 0;
               if (d.Hrs <= 8) r = d.Hrs;
               else if (d.Hrs <= 12) { r = 8; o = d.Hrs - 8; } else { r = 8; o = 4; dt = d.Hrs - 12; }

               straight += r; dailyOT += o; dailyDT += dt;

               sb.AppendLine($"{identifier},{periodEnd:MM/dd/yyyy},{wk.Key:MM/dd/yyyy},{d.Day:MM/dd/yyyy},{d.Hrs:F2},{r:F2},{o:F2},{dt:F2},,,,");
            }

            double weeklyOT = Math.Max(0.0, straight - 40.0);
            double finalOT = Math.Max(dailyOT, weeklyOT);
            double finalReg = weeklyOT > dailyOT ? Math.Max(0, straight - (weeklyOT - dailyOT)) : straight;

            sb.AppendLine($"{identifier},{periodEnd:MM/dd/yyyy},{wk.Key:MM/dd/yyyy},WEEK_SUMMARY,,,,," +
                          $"{straight:F2},{dailyOT:F2},{dailyDT:F2},{weeklyOT:F2},{finalReg:F2},{finalOT:F2},{dailyDT:F2}");
         }

         File.AppendAllText(file, sb.ToString());
      }
   }
}
