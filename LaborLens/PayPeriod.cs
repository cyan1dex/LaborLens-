using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
   class Period {
      public DateTime start;
      public DateTime end;

      public Period(bool populateEnds)
      {
         start = DateTime.MaxValue;
         end = DateTime.MinValue;
      }

      public Period(DateTime start, DateTime end)
      {
         this.start = start;
         this.end = end;
      }

      public Period()
      {

      }
   }

   public class  PayPeriods {
      double regPayTotal;
      double regHrsTotal;

      double otPayTotal;
      double otHrsTotal;

      int totalSheets;

      #region paga numbers
      public int perW3n5 = 0;
      public int perW5andViol = 0;
      public int perW6andViol = 0;
      public int perW5n6Viol = 0;
      public int perW6n10Viol = 0;
      public int perW10Viol = 0;
      public int perW10n12Viol = 0;
      public int perW12Viol = 0;


      public int perW5andCnt = 0;
      public int perW6andCnt = 0;
      public int perW5n6Cnt = 0;
      public int per6n10Cnt = 0;
      public int perW10Cnt = 0;
      public int perW10n12Cnt = 0;
      public int perW12Cnt = 0;

      int pagaSplit1hr = 0;
      int pagaSplit2hr;

      int pagaPeriods = 0;
      #endregion

      int violationCount = 0;

      public int validPeriods = 0;
      public int invalidPayperiods = 0;

      int totalPeriods = 0;

      int HoursMatch;
      int HoursDoNotMatch;
      int missingTimeCards;
      double swingHrsPaid = 0;

      Dictionary<string, int> missingCards = new Dictionary<string, int>();

      public void CalculatePagaPeriods(Dictionary<string, List<Timesheet>> sheets, DateTime pagaDate)
      {
         int roundedOtUnpaid = 0;
         foreach (KeyValuePair<string, List<Timesheet>> entry in sheets) {
            foreach (Timesheet sheet in entry.Value) {
               totalPeriods++;

               foreach (Timecard c in sheet.timeCards) {
                  if (c.totalHrsActual.TotalHours > 12) {
                     int pause = 0;
                  }

               }

               if (sheet.periodEnd >= pagaDate) {

                  if (sheet.ProcessRounding())
                     roundedOtUnpaid++;
                  pagaPeriods++;

                  bool p35 = false;
                  bool p5 = false;
                  bool p6 = false;
                  bool p56 = false;
                  bool p610 = false;
                  bool p10 = false;
                  bool p12 = false;
                  bool p1012 = false;

                  bool ps1 = false;
                  bool ps2 = false;


                  bool cnt5 = false;
                  bool cnt6 = false;
                  bool cnt56 = false;
                  bool cnt610 = false;
                  bool cnt10 = false;
                  bool cnt12 = false;
                  bool cnt1012 = false;


                  foreach (Timecard c in sheet.timeCards) {
                     if (c.splitShiftLenth.Hours >= 1)
                        ps1 = true;
                     if (c.splitShiftLenth.Hours >= 2)
                        ps2 = true;
                     if (c.totalHrsActual.TotalHours > 3.5)
                        p35 = true;

                     if (c.totalHrsActual.TotalHours >= 5 && c.HasViolation())
                        p5 = true;
                     if (c.totalHrsActual.TotalHours >= 5)
                        cnt5 = true;


                     if (c.totalHrsActual.TotalHours > 6 && c.HasViolation())
                        p6 = true;
                     if (c.totalHrsActual.TotalHours > 6)
                        cnt6 = true;

                     if (c.totalHrsActual.TotalHours >= 10 && c.HasSecondMealViolation())
                        p10 = true;
                     if (c.totalHrsActual.TotalHours >= 10)
                        cnt10 = true;

                     if (c.totalHrsActual.TotalHours > 12 && c.HasSecondMealViolation())
                        p12 = true;
                     if (c.totalHrsActual.TotalHours > 12)
                        cnt12 = true;

                     if (c.totalHrsActual.TotalHours >= 5 && c.totalHrsActual.TotalHours <= 6 && c.HasViolation())
                        p56 = true;
                     if (c.totalHrsActual.TotalHours >= 5 && c.totalHrsActual.TotalHours <= 6)
                        cnt56 = true;


                     if (c.totalHrsActual.TotalHours > 6 && c.totalHrsActual.TotalHours < 10 && c.HasViolation())
                        p610 = true;
                     if (c.totalHrsActual.TotalHours > 6 && c.totalHrsActual.TotalHours < 10)
                        cnt610 = true;

                     if (c.totalHrsActual.TotalHours >= 10 && c.totalHrsActual.TotalHours <= 12 && c.HasSecondMealViolation())
                        p1012 = true;
                     if (c.totalHrsActual.TotalHours >= 10 && c.totalHrsActual.TotalHours <= 12)
                        cnt1012 = true;

                  }



                  perW5andCnt += cnt5 == true ? 1 : 0;
                  perW6andCnt += cnt6 == true ? 1 : 0;
                  perW5n6Cnt += cnt56 == true ? 1 : 0;
                  per6n10Cnt += cnt610 == true ? 1 : 0;
                  perW10Cnt += cnt10 == true ? 1 : 0;
                  perW10n12Cnt += cnt1012 == true ? 1 : 0;
                  perW12Cnt += cnt12 == true ? 1 : 0;


                  perW3n5 += p35 == true ? 1 : 0;
                  perW5andViol += p5 == true ? 1 : 0;
                  perW6andViol += p6 == true ? 1 : 0;
                  perW5n6Viol += p56 == true ? 1 : 0;
                  perW6n10Viol += p610 == true ? 1 : 0;
                  perW10Viol += p10 == true ? 1 : 0;
                  perW10n12Viol += p1012 == true ? 1 : 0;
                  perW12Viol += p12 == true ? 1 : 0;

                  pagaSplit1hr += ps1 == true ? 1 : 0;
                  pagaSplit2hr += ps2 == true ? 1 : 0;
               }
            }
         }
      }

      public void CalculateRounding(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         int totFavorEmp = 0;
         int totFavorCmpny = 0;

         double moneyFavoerEmp = 0;
         double moneyFavorComp = 0;
         int autoDeduct = 0;
         double unpaid = 0;

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            foreach (Timesheet s in employee.Value) {
               if (s.invalid)
                  continue;

               if (Math.Abs(s.actualTotalHours - (s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs)) > 4)
                  continue;

               foreach (Timecard t in s.timeCards) {
                  if (t.possibleAutoDeduct)
                     autoDeduct++;

                  if (Math.Abs(t.regHrsListed - t.totalHrsActual.TotalHours) > .01) {
                     if (t.regHrsListed > t.totalHrsActual.TotalHours) totFavorEmp++;
                     else {
                        //     totFavorCmpny++;
                        if (t.totalHrsActual.TotalHours > 12)
                           unpaid += t.totalHrsActual.TotalHours - t.regHrsListed;
                     }
                  }
               }

               var act = s.actualTotalHours;// s.actualHours.TotalHours;// - s.actualOT.TotalHours; //Actual hours are hrs worked minus OT

               var tot = s.listedTotalHours;

               var check = s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs; //Check hours can be different from the timecard hours

               var otList = s.stub.otHrs;

               var otACt = s.actualOT.TotalHours;

               var otCheck = s.stub.doubleOtHrs;

               var difference = Math.Abs(act - check);

               if (difference < .25) {
                  if (check > act) {
                     totFavorEmp++;
                     moneyFavoerEmp += difference;
                  } else {
                     totFavorCmpny++;
                     moneyFavorComp += difference;
                  }

               }
               // else same++;
            }
         }
      }


      private System.DayOfWeek? AdpEmployerWorkweekStart = null;

      private System.DayOfWeek ResolveAdpWorkweek(
          Dictionary<string, List<PayStub>> stubs)
      {
         if (AdpEmployerWorkweekStart.HasValue)
            return AdpEmployerWorkweekStart.Value;

         var starts = stubs.Values.SelectMany(v => v)
             .Where(p => p != null && !p.invalid &&
                         p.periodBegin.HasValue && p.periodEnd.HasValue)
             .Where(p => {
                int days = (p.periodEnd.Value.Date - p.periodBegin.Value.Date).Days + 1;
                return days == 7 || days == 14;
             })
             .Select(p => p.periodBegin.Value.DayOfWeek)
             .Distinct().ToList();

         if (starts.Count != 1)
            throw new InvalidOperationException(
                "Workweek start not established. Set AdpEmployerWorkweekStart " +
                "to the employer-defined start day. The pay stubs " +
                (starts.Count == 0 ? "contain no 7-/14-day periods." :
                 "contain conflicting 7-/14-day period start days."));
         return starts[0];
      }

      public Dictionary<string, List<Timesheet>> PopulateADPTimesheets(
          Dictionary<string, List<PayStub>> stubs,
          Dictionary<string, List<Timecard>> timeCards)
      {
         System.DayOfWeek workweekStart = ResolveAdpWorkweek(stubs);
         var timesheets = new Dictionary<string, List<Timesheet>>();
         // Avoid counting a semimonthly workweek twice when it spans two pay periods.
         var uniqueEmployeeWeeks = new HashSet<string>();
         var failures = new List<string>();
         failures.Add("Employee,PeriodBegin,PeriodEnd,Type,Message");
         // Payroll date coverage is not independent proof that timecards were
         // produced for those dates. Without source production metadata, we
         // cannot safely infer whether a boundary day was truly a non-workday.
         failures.Add(CsvAuditLine("ALL", DateTime.MinValue, DateTime.MinValue,
             "LIMITATION", "Payroll-period coverage is not proof of timecard " +
             "production completeness. Verify missing timecard periods, especially " +
             "at production, hire, and termination boundaries."));
         if (!AdpEmployerWorkweekStart.HasValue)
            failures.Add(CsvAuditLine("ALL", DateTime.MinValue, DateTime.MinValue,
                "ASSUMPTION", "Workweek start inferred from 7/14-day payroll " +
                "period boundaries. Confirm employer's actual defined workweek."));

         foreach (KeyValuePair<string, List<PayStub>> checks in stubs) {
            timesheets[checks.Key] = new List<Timesheet>();
            List<Timecard> cards;
            if (!timeCards.TryGetValue(checks.Key, out cards) || cards == null) {
               foreach (PayStub check in checks.Value) {
                  if (check.invalid || !check.periodBegin.HasValue || !check.periodEnd.HasValue)
                     continue;
                  if (!missingCards.ContainsKey(checks.Key)) missingCards[checks.Key] = 0;
                  missingCards[checks.Key]++;
               }
               continue;
            }

            // Keep every legitimate timecard, even when an employee worked two
            // shifts on the same date. Do not use shiftID or punch times.
            var datedCards = cards.Where(c => c != null && c.shiftDate.HasValue)
                                  .OrderBy(c => c.shiftDate.Value.Date).ToList();
            bool undatedWorkedCards = cards.Any(c => c != null && !c.shiftDate.HasValue &&
                                                      c.totalHrsActual.TotalHours > 0);

            // Coverage comes from the employee's produced pay periods, including
            // zero-work dates; absence of a shift does not prove absence of work.
            var coverage = checks.Value
                .Where(c => c != null && !c.invalid && c.periodBegin.HasValue && c.periodEnd.HasValue)
                .Select(c => new OvertimePeriodCoverage {
                   Begin = c.periodBegin.Value.Date, End = c.periodEnd.Value.Date
                }).ToList();

            foreach (PayStub check in checks.Value) {
               if (check.invalid || !check.periodBegin.HasValue || !check.periodEnd.HasValue)
                  continue;

               var sheet = new Timesheet {
                  periodBegin = check.periodBegin,
                  periodEnd = check.periodEnd,
                  identifier = check.identifier,
                  stub = check
               };
               totalSheets++;

               var b = check.periodBegin.Value.Date;
               var e = check.periodEnd.Value.Date;
               foreach (var card in datedCards) {
                  DateTime d = card.shiftDate.Value.Date;
                  if (d < b) continue;
                  if (d > e) break;
                  sheet.timeCards.Add(card);
                  if (card.invalid) sheet.invalid = true;
               }

               try {
                  // A correction/off-cycle check can have dates overlapping a
                  // regular check. Both would otherwise be assigned the same
                  // timecards and produce duplicated OT. Without metadata that
                  // identifies the proper check, exclude both periods' OT from
                  // aggregate output and report them for manual resolution.
                  bool overlapsAnotherCheck = checks.Value.Any(other =>
                      other != null && !object.ReferenceEquals(other, check) &&
                      !other.invalid && other.periodBegin.HasValue &&
                      other.periodEnd.HasValue &&
                      other.periodBegin.Value.Date <= e &&
                      other.periodEnd.Value.Date >= b);
                  if (overlapsAnotherCheck) {
                     sheet.invalid = true;
                     failures.Add(CsvAuditLine(checks.Key, b, e, "ERROR",
                         "Overlapping paystub periods would classify the same " +
                         "timecards more than once. Resolve regular/off-cycle " +
                         "check association before counting this period's OT."));
                  }
                  if (undatedWorkedCards) {
                     sheet.invalid = true;
                     failures.Add(CsvAuditLine(checks.Key, b, e, "ERROR",
                         "Employee has positive-hour cards with missing shiftDate; " +
                         "cannot determine which pay period they belong to."));
                  }
                  if (!undatedWorkedCards && !overlapsAnotherCheck)
                     sheet.AnalyzeADPHours(datedCards, coverage, workweekStart);
                  if (!string.IsNullOrEmpty(sheet.OvertimeError))
                     failures.Add(CsvAuditLine(checks.Key, b, e, "ERROR", sheet.OvertimeError));
                  if (sheet.LastOvertimeResult != null)
                     foreach (string warning in sheet.LastOvertimeResult.Warnings)
                        failures.Add(CsvAuditLine(checks.Key, b, e, "WARNING", warning));
               } catch (Exception ex) {
                  // A bad employee-period cannot abort all employee analyses.
                  sheet.invalid = true;
                  failures.Add(CsvAuditLine(checks.Key, b, e, "EXCEPTION", ex.Message));
               }

               if (!sheet.invalid && sheet.timeCards.Count > 0) {
                  if (Math.Abs(sheet.actualTotalHours - sheet.listedTotalHours) < 1.5)
                     HoursMatch++;
                  else {
                     HoursDoNotMatch++;
                     if (sheet.actualTotalHours == 0 || sheet.listedTotalHours == 0)
                        missingTimeCards++;
                  }

                  if (sheet.HasViolation()) violationCount++;
                  timesheets[checks.Key].Add(sheet);
                  foreach (var card in sheet.timeCards.Where(t => t.totalHrsActual.TotalHours > 0)) {
                     var wk = CaliforniaOvertime.WeekStart(card.shiftDate.Value.Date, workweekStart);
                     if (uniqueEmployeeWeeks.Add(checks.Key + "|" + wk.ToString("yyyyMMdd")))
                        Timesheet.totalWorkweeks++;
                  }
               } else
                  invalidPayperiods++;

               // Preserve your historical pay aggregates; still count pay even
               // when the corresponding time-period cannot be classified.
               regPayTotal += check.regPay;
               regHrsTotal += check.regHrs;
               otHrsTotal += check.otHrs;
               otPayTotal += check.otPay;
            }
         }

         // Keep diagnostics outside the data output, but ALWAYS create an audit
         // file to distinguish reviewed exceptions from successful calculations.
         string auditPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
             "ADP_OT_Exceptions.csv");
         // Retain a unique historical copy; the familiar path remains the latest.
         string archivedAuditPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
             "ADP_OT_Exceptions_" + DateTime.UtcNow.ToString("yyyyMMdd_HHmmss_fff") +
             "_" + Guid.NewGuid().ToString("N") + ".csv");
         System.IO.File.WriteAllLines(archivedAuditPath, failures);
         System.IO.File.WriteAllLines(auditPath, failures);
         return timesheets;
      }

      private static string CsvAuditLine(string emp, DateTime start, DateTime end,
                                         string type, string message)
      {
         // Escape CSV quotes / line breaks; employee ID contains no commas ordinarily.
         Func<string, string> csv = s => "\"" + (s ?? "").Replace("\"", "\"\"")
             .Replace("\r", " ").Replace("\n", " ") + "\"";
         return csv(emp) + "," + start.ToString("yyyy-MM-dd") + "," +
                end.ToString("yyyy-MM-dd") + "," + csv(type) + "," + csv(message);
      }

      public Dictionary<string, List<Timesheet>> PopulateADPTimesheets(Dictionary<string, List<PayStub>> stubs, Dictionary<string, List<Timesheet>> empSheets)
      {
         Dictionary<string, List<Timesheet>> timesheets = new Dictionary<string, List<Timesheet>>();

         foreach (KeyValuePair<string, List<PayStub>> checks in stubs) {
            timesheets[checks.Key] = new List<Timesheet>(); //Create a new timesheet, all timecards with paystubs

            foreach (PayStub check in checks.Value) {
               if (check.invalid || check.periodsMissing)
                  continue;

               Timesheet sheet = new Timesheet();
               sheet.stub = check;
               totalSheets++;

               if (empSheets.ContainsKey(checks.Key)) //locate all of employee's timecards
               {
                  foreach (Timesheet eSheet in empSheets[checks.Key]) {
                     if (!eSheet.missingPeriodDates && (eSheet.periodEnd == check.periodEnd || eSheet.periodBegin == check.periodBegin))
                        foreach (Timecard card in eSheet.timeCards) {

                           //if (card.shiftDate == null)
                           //    continue;

                           //if (card.shiftDate.Value.Date < check.periodBegin)
                           //    continue;
                           //else if (card.shiftDate.Value.Date > check.periodEnd)
                           //    break; //can be break if the cards are in order

                           sheet.timeCards.Add(card);
                           if (card.invalid)
                              sheet.invalid = true;
                        }
                  }

                  if (!sheet.invalid && sheet.timeCards.Count > 0) {
                     sheet.AnalyzeADPHours();

                     if (Math.Abs(sheet.actualTotalHours - sheet.listedTotalHours) < 1.5)
                        HoursMatch++;
                     else {
                        HoursDoNotMatch++;
                        if (sheet.actualTotalHours == 0 || sheet.listedTotalHours == 0)
                           missingTimeCards++;
                     }


                     if (!sheet.invalid && sheet.timeCards.Count > 0) {
                        if (sheet.HasViolation())
                           violationCount++;
                        timesheets[checks.Key].Add(sheet);
                        //if (sheet.oneThreeNHalfShift)
                        //    perW3n5++;
                     } else
                        invalidPayperiods++;

                     //Add in Pay in Hours for all pay types
                     regPayTotal += check.regPay;
                     regHrsTotal += check.regHrs;

                     otHrsTotal += check.otHrs;
                     otPayTotal += check.otPay;
                  }
               } else {
                  if (!missingCards.ContainsKey(checks.Key))
                     missingCards[checks.Key] = 0;
                  missingCards[checks.Key]++;
               }

               if (!sheet.invalid && sheet.timeCards.Count > 0)
                  validPeriods++;
               else
                  invalidPayperiods++;
               // throw new Exception("Employee time cards not found");
            }
         }

         return timesheets;
      }
      #region notUpdated


      #endregion
   }
}
