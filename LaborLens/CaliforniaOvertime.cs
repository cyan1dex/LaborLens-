using System;
using System.Collections.Generic;
using System.Linq;

namespace LaborLens
{
    // Period coverage is evidence that a date was included in the records produced,
    // including a day on which the employee did not work. An absent timecard alone
    // is NOT evidence that the employee was off.
    public sealed class OvertimePeriodCoverage
    {
        public DateTime Begin { get; set; }
        public DateTime End { get; set; }
    }

    public sealed class OvertimeDayResult
    {
        public DateTime Date { get; set; }
        public DateTime WeekStart { get; set; }
        public double Worked { get; set; }
        public double Regular { get; set; }
        public double Overtime { get; set; }
        public double Doubletime { get; set; }
    }

    public sealed class OvertimePeriodResult
    {
        public OvertimeResult Hours { get; set; } = new OvertimeResult();
        public List<OvertimeDayResult> Days { get; } = new List<OvertimeDayResult>();
        public List<string> Errors { get; } = new List<string>();
        public List<string> Warnings { get; } = new List<string>();
        public bool IsValid { get { return Errors.Count == 0; } }
    }

    // Standard CA nonexempt schedule (8/12 daily, 40 weekly, seventh day).
    // One source of calculations for weekly, biweekly and semimonthly payroll.
    // Treats each timecard's hours as attributable to timecard.shiftDate.
    // NOT an alternative workweek, special wage-order, or midnight-splitting engine.
    public static class CaliforniaOvertime
    {
        private const double EPS = 0.000001;

        public static DateTime WeekStart(DateTime day, System.DayOfWeek start)
        {
            int delta = ((int)day.DayOfWeek - (int)start + 7) % 7;
            return day.Date.AddDays(-delta);
        }

        public static OvertimePeriodResult ForPayPeriod(
            IEnumerable<Timecard> allEmployeeCards,
            IEnumerable<OvertimePeriodCoverage> producedPeriods,
            DateTime begin, DateTime end,
            System.DayOfWeek workweekStart)
        {
            var answer = new OvertimePeriodResult();
            if (allEmployeeCards == null || producedPeriods == null)
            {
                answer.Errors.Add("Employee timecards or payroll coverage is missing.");
                return answer;
            }
            begin = begin.Date;
            end = end.Date;
            if (end < begin)
            {
                answer.Errors.Add("Pay period ends before it begins.");
                return answer;
            }
            var first = WeekStart(begin, workweekStart);
            var last = WeekStart(end, workweekStart).AddDays(6);

            var periods = producedPeriods.Where(p => p != null).ToList();
            var uncovered = new List<DateTime>();
            // Verify coverage of EVERY day in every workweek touched by the pay period.
            for (var date = first; date <= last; date = date.AddDays(1))
            {
                if (!periods.Any(p => p.Begin.Date <= date && p.End.Date >= date))
                    uncovered.Add(date);
            }
            if (uncovered.Count > 0)
                answer.Errors.Add("Incomplete workweek coverage: " +
                    string.Join(", ", uncovered.Select(d => d.ToString("yyyy-MM-dd"))));

         // Include cards that start the day before the first week, since an
         // overnight shift can carry hours into that week.
         var candidates = allEmployeeCards.Where(t => t != null && t.shiftDate.HasValue &&
                                    t.shiftDate.Value.Date >= first.AddDays(-1) &&
                                    t.shiftDate.Value.Date <= last).ToList();
         if (candidates.Any(t => t.invalid ||
                     double.IsNaN(t.totalHrsActual.TotalHours) ||
                     double.IsInfinity(t.totalHrsActual.TotalHours) ||
                     t.totalHrsActual.TotalHours < 0))
            answer.Errors.Add("Invalid or negative-hour timecard in a relevant workweek.");
         if (!answer.IsValid) return answer;

         var dayHours = HoursByDay(candidates);

         for (var weekStart = first; weekStart <= last; weekStart = weekStart.AddDays(7)) {
            bool seventhDay = Enumerable.Range(0, 7).All(i =>
                dayHours.ContainsKey(weekStart.AddDays(i)) && dayHours[weekStart.AddDays(i)] > EPS);
            double regularSoFar = 0;

            for (int i = 0; i < 7; i++) {
               DateTime date = weekStart.AddDays(i);
               double hrs;
               if (!dayHours.TryGetValue(date, out hrs) || hrs <= EPS) continue;

               double straightEligible, ot, dt;
               if (seventhDay && i == 6) {
                  straightEligible = 0;
                  ot = Math.Min(8, hrs);
                  dt = Math.Max(0, hrs - 8);
               } else {
                  straightEligible = Math.Min(8, hrs);
                  ot = Math.Min(4, Math.Max(0, hrs - 8));
                  dt = Math.Max(0, hrs - 12);
               }

               double reg = Math.Min(straightEligible, Math.Max(0, 40 - regularSoFar));
               ot += straightEligible - reg;
               regularSoFar += reg;

               if (date < begin || date > end) continue;
               answer.Days.Add(new OvertimeDayResult {
                  Date = date, WeekStart = weekStart, Worked = hrs,
                  Regular = reg, Overtime = ot, Doubletime = dt
               });
               answer.Hours.RegularHours += reg;
               answer.Hours.OvertimeHours += ot;
               answer.Hours.DoubletimeHours += dt;
            }
         }

         double worked = dayHours.Where(kv => kv.Key >= begin && kv.Key <= end).Sum(kv => kv.Value);
         double classified = answer.Hours.RegularHours + answer.Hours.OvertimeHours +
                             answer.Hours.DoubletimeHours;
         if (Math.Abs(classified - worked) > 0.0001)
            answer.Errors.Add(string.Format("Hour reconciliation failed: worked {0:F5}, classified {1:F5}.",
                                           worked, classified));
         return answer;
      }

      // Splits each card's hours across the calendar days its punches actually cover.
      // Total hours per card stay exactly equal to totalHrsActual.
      private static Dictionary<DateTime, double> HoursByDay(IEnumerable<Timecard> cards)
      {
         var byDay = new Dictionary<DateTime, double>();
         foreach (var card in cards) {
            double cardHours = card.totalHrsActual.TotalHours;
            var p = (card.timepunches ?? new List<Timepunch>())
                    .Select(x => x.datetime).OrderBy(x => x).ToList();

            // Punched minutes on each date, from in/out pairs
            var punched = new Dictionary<DateTime, double>();
            for (int i = 0; i + 1 < p.Count; i += 2) {
               DateTime a = p[i], b = p[i + 1];
               while (a < b) {
                  DateTime cut = b < a.Date.AddDays(1) ? b : a.Date.AddDays(1);
                  double cur;
                  punched.TryGetValue(a.Date, out cur);
                  punched[a.Date] = cur + (cut - a).TotalHours;
                  a = cut;
               }
            }

            double totalPunched = punched.Values.Sum();
            if (totalPunched <= EPS) {
               // No usable punch pairs: keep the old behavior for this card
               double cur;
               byDay.TryGetValue(card.shiftDate.Value.Date, out cur);
               byDay[card.shiftDate.Value.Date] = cur + cardHours;
               continue;
            }
            foreach (var kv in punched) {
               double cur;
               byDay.TryGetValue(kv.Key, out cur);
               byDay[kv.Key] = cur + cardHours * kv.Value / totalPunched;
            }
         }
         return byDay;
      }
   }
}
