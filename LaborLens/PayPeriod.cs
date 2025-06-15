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


      public Dictionary<string, List<Timesheet>> PopulateADPTimesheets(Dictionary<string, List<PayStub>> stubs, Dictionary<string, List<Timecard>> timeCards)
      {
         Dictionary<string, List<Timesheet>> timesheets = new Dictionary<string, List<Timesheet>>();


         foreach (KeyValuePair<string, List<PayStub>> checks in stubs) {
            timesheets[checks.Key] = new List<Timesheet>(); //Create a new timesheet, all timecards with paystubs

            if(checks.Key == "4") {
               int p = 0;
            }

            foreach (PayStub check in checks.Value) {
               if (check.invalid || check.periodEnd is null)
                  continue;

               Timesheet sheet = new Timesheet() { periodBegin = check.periodBegin, periodEnd = check.periodEnd, identifier = check.identifier };
               sheet.stub = check;
               totalSheets++;

               if (timeCards.ContainsKey(checks.Key)) //locate all of employee's timecards
               {
                  //sort timecards
                  timeCards[checks.Key] = timeCards[checks.Key].OrderBy(o => o.shiftDate).ToList();
                  var v = timeCards[checks.Key];

                  foreach (Timecard card in timeCards[checks.Key]) {
                     if (card.shiftDate.Value.Date < check.periodBegin.Value.Date)
                        continue;
                     else if (card.shiftDate.Value.Date > check.periodEnd.Value.Date)
                        continue; //can be break if the cards are in order

                     if (sheet.timeCards.Count == 0)
                        sheet.timeCards.Add(card);
                     else if (sheet.timeCards[sheet.timeCards.Count - 1].shiftDate != card.shiftDate)
                        sheet.timeCards.Add(card);

                     if (card.invalid)
                        sheet.invalid = true;
                  }

                  sheet.AnalyzeADPHours();
                  if (sheet.timeCards.Count > 20) { //If OT from 6th day
                     int pause = 0;
                  }

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
                  } else
                     invalidPayperiods++;

                  //Add in Pay in Hours for all pay types
                  regPayTotal += check.regPay;
                  regHrsTotal += check.regHrs;

                  otHrsTotal += check.otHrs;
                  otPayTotal += check.otPay;
               } else {
                  if (!missingCards.ContainsKey(checks.Key))
                     missingCards[checks.Key] = 0;
                  missingCards[checks.Key]++;
               }
               // throw new Exception("Employee time cards not found");
            }
         }


         return timesheets;
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

      public Dictionary<string, List<Timesheet>> PopulateTimesheets(Dictionary<string, List<PayStub>> stubs, Dictionary<string, List<Timecard>> timeCards)
      {
         Dictionary<string, List<Timesheet>> timesheets = new Dictionary<string, List<Timesheet>>();
         double total = 0;
         foreach (KeyValuePair<string, List<PayStub>> checks in stubs) {
            timesheets[checks.Key] = new List<Timesheet>();

            foreach (PayStub check in checks.Value) {
               Timesheet sheet = new Timesheet();
               sheet.stub = check;
               totalSheets++;

               if (timeCards.ContainsKey(checks.Key)) {
                  foreach (Timecard card in timeCards[checks.Key]) {
                     if (card.shiftDate.Value.Date < check.periodBegin)
                        continue;
                     else if (card.shiftDate.Value.Date > check.periodEnd)
                        continue; //can be break if the cards are in order

                     sheet.timeCards.Add(card);
                  }

                  sheet.AnalyzeDelunaHours(); //RUN ANALYSIS ON THE SHEET

                  total += sheet.listedTotalHours;
                  // if (sheet.mealMissedOrUnder30 > 0)
                  //    mealMissedorUnder++;

                  // double differnceHours = sheet.actualCombinedHours.TotalHours - (sheet.listedTotalHours);
                  //if (Math.Abs(differnceHours) < .5)
                  //   total += differnceHours;

                  timesheets[checks.Key].Add(sheet);
                  regPayTotal += check.regPay;
                  regHrsTotal += check.regHrs;
                  otHrsTotal += check.otHrs;
                  otPayTotal += check.otPay;
               } else
                  throw new Exception("Empployee time cards not found");
            }
         }

         return timesheets;
      }
      #endregion
   }
}
