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

      // ===== Shared overtime engine (all pay frequencies) =====
      // Configured by the caller using the employer's established workweek.
      // A 7/14-day payroll period can be used to infer a PROVISIONAL start only
      // when the employer confirms pay-period and workweek boundaries match.
      private List<Timecard> overtimeSourceCards;
      private List<OvertimePeriodCoverage> overtimeSourcePeriods;
      public System.DayOfWeek? OvertimeWorkweekStart { get; private set; }
      public OvertimePeriodResult LastOvertimeResult { get; private set; }
      public string OvertimeError { get; private set; }

      public void ConfigureOvertimeContext(
         List<Timecard> allEmployeeCards,
         List<OvertimePeriodCoverage> periodCoverage,
         System.DayOfWeek workweekStart)
      {
         overtimeSourceCards = allEmployeeCards;
         overtimeSourcePeriods = periodCoverage;
         OvertimeWorkweekStart = workweekStart;
      }

      public int GetWorkWeeks()
      {
         // ADP sheets use the explicit employer workweek once configured.
         if (OvertimeWorkweekStart.HasValue)
            return timeCards.Where(t => t != null && t.shiftDate.HasValue &&
                                      t.totalHrsActual.TotalHours > 0)
               .Select(t => CaliforniaOvertime.WeekStart(t.shiftDate.Value.Date,
                                                         OvertimeWorkweekStart.Value))
               .Distinct().Count();

         // Backward compatibility for callers which never configured the
         // ADP-only overtime context. Preserve the former period-based count;
         // DO NOT throw or infer an employer workweek from it.
         if (!periodBegin.HasValue) return 2;
         var split = periodBegin.Value.Date.AddDays(6);
         bool first = timeCards.Any(t => t != null && t.shiftDate.HasValue &&
                                         t.shiftDate.Value.Date <= split);
         bool second = timeCards.Any(t => t != null && t.shiftDate.HasValue &&
                                          t.shiftDate.Value.Date > split);
         return (first ? 1 : 0) + (second ? 1 : 0);
      }

      // An earlier version silently used periodBegin.DayOfWeek when no
      // employer workweek was configured. A pay-period boundary is not proof
      // of the employer's defined workweek; fail visibly instead of guessing.
      private void EnsureAlignedPeriodContextWhenPossible(List<Timecard> cards)
      {
         if (!OvertimeWorkweekStart.HasValue || overtimeSourcePeriods == null ||
             overtimeSourceCards == null)
            throw new InvalidOperationException(
               "Workweek context has not been configured. Call " +
               "ConfigureOvertimeContext with the employee timecards, " +
               "period coverage, and confirmed employer workweek start.");
      }

      private OvertimePeriodResult RunSharedOvertime(List<Timecard> source)
      {
         EnsureAlignedPeriodContextWhenPossible(source);
         if (!periodBegin.HasValue || !periodEnd.HasValue)
            throw new InvalidOperationException("Pay-period dates are required for OT allocation.");
         if (!OvertimeWorkweekStart.HasValue || overtimeSourcePeriods == null)
            throw new InvalidOperationException(
               "Overtime context is not configured. Supply all employee cards, " +
               "pay-period coverage, and the employer workweek start.");
         return CaliforniaOvertime.ForPayPeriod(source, overtimeSourcePeriods,
            periodBegin.Value, periodEnd.Value, OvertimeWorkweekStart.Value);
      }

      // Keep the public legacy method callable without ADP configuration.
      // For ADP, the configured context always uses the shared calculator.
      // For non-ADP legacy callers, retain the prior Sunday-week behavior,
      // with the corrected arithmetic in CaliforniaOvertime.
      public OvertimeResult CalculateOvertime(List<Timecard> timecards, bool trace = false, string tag = "")
      {
         if (timecards == null || timecards.Count == 0) return new OvertimeResult();
         if (OvertimeWorkweekStart.HasValue && overtimeSourcePeriods != null &&
             overtimeSourceCards != null) {
            var calculated = RunSharedOvertime(overtimeSourceCards);
            LastOvertimeResult = calculated;
            if (!calculated.IsValid) {
               OvertimeError = string.Join("; ", calculated.Errors);
               invalid = true;
               return new OvertimeResult();
            }
            return calculated.Hours;
         }

         // Non-ADP backward-compatible route. This does NOT establish a
         // workweek for ADP; AnalyzeADPHours requires explicit context.
         if (timecards.Any(t => t == null || !t.shiftDate.HasValue || t.invalid ||
                               t.totalHrsActual.TotalHours < 0)) {
            invalid = true;
            OvertimeError = "Missing date or invalid hours in legacy overtime input.";
            return new OvertimeResult();
         }
         var first = CaliforniaOvertime.WeekStart(
            timecards.Min(t => t.shiftDate.Value.Date), System.DayOfWeek.Sunday);
         var last = CaliforniaOvertime.WeekStart(
            timecards.Max(t => t.shiftDate.Value.Date), System.DayOfWeek.Sunday).AddDays(6);
         var coverage = new List<OvertimePeriodCoverage> {
            new OvertimePeriodCoverage { Begin = first, End = last }
         };
         var legacy = CaliforniaOvertime.ForPayPeriod(
            timecards, coverage, first, last, System.DayOfWeek.Sunday);
         LastOvertimeResult = legacy;
         if (!legacy.IsValid) {
            invalid = true;
            OvertimeError = string.Join("; ", legacy.Errors);
            return new OvertimeResult();
         }
         return legacy.Hours;
      }

      // Legacy method name retained, but no separate (incorrect) calculation.
      public (double regularHours, double overtimeHours, double doubleTimeHours)
         CalculateOTHoursBiMonthly(Timesheet timesheet)
      {
         var result = timesheet.CalculateOvertime(timesheet.timeCards);
         return (result.RegularHours, result.OvertimeHours, result.DoubletimeHours);
      }

      public void CalculateNonStub()
      {
         actualHours = TimeSpan.Zero;
         actualOT = TimeSpan.Zero;
         actualDblOT = TimeSpan.Zero;
         actualTotalHours = 0;
         SetListedHoursFromTimecard();
         if (timeCards == null || timeCards.Count == 0) return;

         // CalculateOvertime handles both preexisting legacy callers and
         // explicitly configured ADP callers. No new initialization contract.
         var result = CalculateOvertime(timeCards);
         if (invalid) return;
         ApplyCalculatedHours(result);

         var dates = timeCards.Where(t => t != null && t.shiftDate.HasValue &&
                                   t.totalHrsActual.TotalHours > 0)
             .Select(t => t.shiftDate.Value.Date).Distinct().OrderBy(d => d).ToList();
         sevenInArow = dates.Skip(6).Select((d, i) => (d - dates[i]).Days == 6).Any(x => x);
      }

      private void ApplyCalculatedHours(OvertimeResult hours)
      {
         actualHours = TimeSpan.FromHours(hours.RegularHours);
         actualOT = TimeSpan.FromHours(hours.OvertimeHours);
         actualDblOT = TimeSpan.FromHours(hours.DoubletimeHours);
         actualTotalHours = hours.RegularHours + hours.OvertimeHours + hours.DoubletimeHours;
      }

      // All ADP paths use the same engine. Pay frequency does not select a
      // different overtime formula. No workweek weekday is hardcoded.
      public void AnalyzeADPHours(
         List<Timecard> allEmployeeCards,
         List<OvertimePeriodCoverage> payPeriodCoverage,
         System.DayOfWeek workweekStart)
      {
         ConfigureOvertimeContext(allEmployeeCards, payPeriodCoverage, workweekStart);
         AnalyzeADPHours();
      }

      public void AnalyzeADPHours()
      {
         actualHours = TimeSpan.Zero;
         actualOT = TimeSpan.Zero;
         actualDblOT = TimeSpan.Zero;
         actualTotalHours = 0;
         OvertimeError = null;
         LastOvertimeResult = null;

         if (stub == null)
            throw new InvalidOperationException("AnalyzeADPHours requires a PayStub.");
         listHours = TimeSpan.FromHours(stub.regHrs);
         listOT = TimeSpan.FromHours(stub.otHrs);
         regPay = stub.regPay;
         regRate = stub.regRate;
         otPay = stub.otPay;
         otRate = stub.otRate;
         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;
         checkTotalHours = stub.regHrs + stub.otHrs + stub.doubleOtHrs;
         SetListedHoursFromTimecard();

         if (!periodBegin.HasValue || !periodEnd.HasValue) {
            invalid = true;
            OvertimeError = "Pay-period begin/end is missing.";
            return;
         }
         if (!OvertimeWorkweekStart.HasValue || overtimeSourceCards == null ||
             overtimeSourcePeriods == null) {
            invalid = true;
            OvertimeError = "Overtime context/workweek start not configured.";
            return;
         }

         LastOvertimeResult = RunSharedOvertime(overtimeSourceCards);
         OvertimeError = string.Join("; ", LastOvertimeResult.Errors);
         if (!LastOvertimeResult.IsValid) {
            invalid = true; // exclude affected period, not entire dataset
            return;
         }
         ApplyCalculatedHours(LastOvertimeResult.Hours);

         double listedWorked = timeCards.Where(t => t != null)
                                         .Sum(t => t.totalHrsActual.TotalHours);
         if (Math.Abs(listedWorked - actualTotalHours) > 0.0001) {
            invalid = true;
            OvertimeError = string.Format(
               "Period timecard hours {0:F5} do not equal classified hours {1:F5}.",
               listedWorked, actualTotalHours);
         }

         if (EnableCsvForSuspectPeriods && periodEnd.HasValue &&
             SuspectPeriodEnds.Contains(periodEnd.Value.Date))
            WriteCompactOtCsvForPeriod();
      }

      // Original diagnostic option now writes allocations from the SAME engine.
      private void WriteCompactOtCsvForPeriod()
      {
         if (LastOvertimeResult == null || !LastOvertimeResult.IsValid) return;
         string dir = string.IsNullOrWhiteSpace(CsvOutputDirectory)
                       ? AppDomain.CurrentDomain.BaseDirectory : CsvOutputDirectory;
         string file = Path.Combine(dir,
             string.Format("OT_Debug_{0}_{1:yyyyMMdd}.csv", identifier, periodEnd));
         var sb = new StringBuilder();
         sb.AppendLine("EmpID,PeriodEnd,WeekStart,Date,Worked,Regular,Overtime,Doubletime");
         foreach (var day in LastOvertimeResult.Days.OrderBy(d => d.Date))
            sb.AppendLine(string.Format(System.Globalization.CultureInfo.InvariantCulture,
               "{0},{1:yyyy-MM-dd},{2:yyyy-MM-dd},{3:yyyy-MM-dd},{4:F5},{5:F5},{6:F5},{7:F5}",
               identifier, periodEnd, day.WeekStart, day.Date, day.Worked,
               day.Regular, day.Overtime, day.Doubletime));
         File.WriteAllText(file, sb.ToString());
      }
   }
}
