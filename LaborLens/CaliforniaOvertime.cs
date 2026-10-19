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

            var candidates = allEmployeeCards.Where(t => t != null && t.shiftDate.HasValue &&
                                       t.shiftDate.Value.Date >= first &&
                                       t.shiftDate.Value.Date <= last).ToList();
            if (candidates.Any(t => t.invalid ||
                        double.IsNaN(t.totalHrsActual.TotalHours) ||
                        double.IsInfinity(t.totalHrsActual.TotalHours) ||
                        t.totalHrsActual.TotalHours < 0))
                answer.Errors.Add("Invalid or negative-hour timecard in a relevant workweek.");

            // Without complete data, do not return an apparently valid OT figure.
            if (!answer.IsValid) return answer;

            // Two cards on one date can legitimately represent two shifts; retain both.
            // Identical hours on the same date are ambiguous, so warn rather than
            // silently delete them (there is no reliable unique shift identifier).
            var potentialDuplicates = candidates.GroupBy(t => new {
                Date = t.shiftDate.Value.Date,
                Ticks = t.totalHrsActual.Ticks
            }).Where(g => g.Count() > 1 && g.Key.Ticks > 0).ToList();
            if (potentialDuplicates.Count > 0)
                answer.Warnings.Add("Potential duplicate same-day/same-hours cards: " +
                    string.Join(", ", potentialDuplicates.Select(g =>
                        g.Key.Date.ToString("yyyy-MM-dd") + " (" + g.Count() + ")")));

            foreach (var week in candidates.GroupBy(t => WeekStart(t.shiftDate.Value.Date, workweekStart))
                                          .OrderBy(g => g.Key))
            {
                var daily = week.GroupBy(t => t.shiftDate.Value.Date)
                    .ToDictionary(g => g.Key, g => g.Sum(t => t.totalHrsActual.TotalHours));
                bool seventhDay = Enumerable.Range(0, 7).All(i =>
                    daily.ContainsKey(week.Key.AddDays(i)) && daily[week.Key.AddDays(i)] > EPS);
                double regularSoFar = 0;

                for (int i = 0; i < 7; i++)
                {
                    DateTime date = week.Key.AddDays(i);
                    double hrs;
                    if (!daily.TryGetValue(date, out hrs) || hrs <= EPS) continue;

                    double straightEligible, ot, dt;
                    if (seventhDay && i == 6)
                    {
                        straightEligible = 0;
                        ot = Math.Min(8, hrs);
                        dt = Math.Max(0, hrs - 8);
                    }
                    else
                    {
                        straightEligible = Math.Min(8, hrs);
                        ot = Math.Min(4, Math.Max(0, hrs - 8));
                        dt = Math.Max(0, hrs - 12);
                    }

                    // Daily-premium hours are excluded from the 40-hour straight
                    // bucket. Remaining straight hours over 40 become weekly OT.
                    double reg = Math.Min(straightEligible, Math.Max(0, 40 - regularSoFar));
                    ot += straightEligible - reg;
                    regularSoFar += reg;

                    if (date < begin || date > end) continue;
                    var row = new OvertimeDayResult {
                        Date = date, WeekStart = week.Key, Worked = hrs,
                        Regular = reg, Overtime = ot, Doubletime = dt
                    };
                    answer.Days.Add(row);
                    answer.Hours.RegularHours += reg;
                    answer.Hours.OvertimeHours += ot;
                    answer.Hours.DoubletimeHours += dt;
                }
            }

            double worked = candidates.Where(t => t.shiftDate.Value.Date >= begin &&
                                                 t.shiftDate.Value.Date <= end)
                                      .Sum(t => t.totalHrsActual.TotalHours);
            double classified = answer.Hours.RegularHours + answer.Hours.OvertimeHours +
                                answer.Hours.DoubletimeHours;
            if (Math.Abs(classified - worked) > 0.0001)
                answer.Errors.Add(string.Format("Hour reconciliation failed: worked {0:F5}, classified {1:F5}.",
                                               worked, classified));
            return answer;
        }
    }
}
