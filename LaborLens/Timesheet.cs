using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {

   public class OvertimeResult {
      public double RegularHours { get; set; }
      public double OvertimeHours { get; set; }
      public double DoubletimeHours { get; set; }
   }

   public class Timesheet {
      public DateTime? periodBegin;
      public DateTime? periodEnd;

      public TimeSpan listHours;
      public TimeSpan listOT;

      public TimeSpan actualOT;
      public TimeSpan actualDblOT;
      public TimeSpan actualHours;

      public double listedTotalHours;
      public double actualTotalHours;
      public double checkTotalHours;

      public double regPay;
      public double otPay;

      public double regRate;
      public double otRate;

      public PayStub stub;
      public List<Timecard> timeCards = new List<Timecard>();

      public TimeSpan actMinusListedTime = new TimeSpan();
      public string identifier;
      public bool invalid;

      public bool missingPeriodDates;
      public static int totalWorkweeks;

      public int roundedShiftsForCompany;
      public int roundedShiftsForEmployee;
      public Decimal roudingBalance;
      public int roundedWOrksWeeks;

      public int splitShiftOneHr;
      public int splitShiftTwohr;

      public HashSet<DateTime> holidays;
      public bool workedAllDaysOfaWeek;

      public bool sevenInArow;

      public bool PaidPenalties()
      {
         foreach (Timecard t in timeCards)
            if (t.penalties > 0)
               return true;
         return false;
      }

      public bool HasViolation()
      {
         foreach (Timecard t in timeCards)
            if (t.HasViolation())
               return true;
         return false;
      }
      public bool HasFirstMealViolation()
      {
         foreach (Timecard t in timeCards)
            if (t.HasFirstMealViolation())
               return true;
         return false;
      }

      public bool HasSecondMealViolation()
      {
         foreach (Timecard t in timeCards)
            if (t.HasSecondMealViolation())
               return true;
         return false;
      }

      public void PopulateTimeCardShifts()
      {
         if (!missingPeriodDates) {
            foreach (Timecard t in timeCards)
               t.PopulateShiftDate(periodBegin);
         }
      }

      public void Analyze()
      {
         int daysInWeekOneWrkd = 0;
         int daysInWeekTwoWrked = 0;
         if (this.periodBegin == null)
            totalWorkweeks += 2;

         DateTime midweek = this.periodBegin.Value.AddDays(6); //up to midweek inclusive
         Timecard previous = null;
         foreach (Timecard card in timeCards) {

            if (previous != null) {
               double shiftDif = card.timepunches[0].datetime.Subtract(previous.timepunches[previous.timepunches.Count - 1].datetime).TotalHours;

               if (shiftDif <= 1 && shiftDif >= 0)
                  this.splitShiftOneHr++;
               else if (shiftDif <= 2 && shiftDif >= 0)
                  this.splitShiftTwohr++;
            }

            if (card.shiftDate <= midweek)
               daysInWeekOneWrkd++;
            else if (card.shiftDate > midweek)
               daysInWeekTwoWrked++;

            previous = card;
         }
         if (daysInWeekOneWrkd > 0)
            totalWorkweeks++;
         if (daysInWeekTwoWrked > 0)
            totalWorkweeks++;
      }

      public int GetWorkWeeks()
      {
         int daysInWeekOneWrkd = 0;
         int daysInWeekTwoWrked = 0;

         if (this.periodBegin == null)
            return 2;

         DateTime midweek = this.periodBegin.Value.AddDays(6); //up to midweek inclusive

         foreach (Timecard card in timeCards) {

            if (card.shiftDate <= midweek)
               daysInWeekOneWrkd++;
            else if (card.shiftDate > midweek)
               daysInWeekTwoWrked++;

         }
         if (daysInWeekOneWrkd > 0 && daysInWeekTwoWrked > 0)
            return 2;
         else if (daysInWeekOneWrkd > 0 || daysInWeekTwoWrked > 0)
            return 1;
         else
            return 0;
      }


      public bool ProcessRounding()
      {
         foreach (Timecard card in timeCards) {

            if (card.mealsTaken == 0 && card.totalHrsActual.TotalHours > 5) {

               double actHrs = card.totalHrsActual.TotalHours - 1;
               if (actHrs > 8) {
                  if (actHrs - card.regHrsListed > 0)
                     roundedShiftsForCompany++;
                  else if (actHrs - card.regHrsListed < 0)
                     roundedShiftsForEmployee++;

                  roudingBalance += (Decimal)(actHrs - card.regHrsListed);
               }
            }
         }
         if (timeCards.Count > 0 && timeCards[0].shiftDate < new DateTime(2016, 2, 6))
            roundedWOrksWeeks += 2;

         if (roudingBalance > 0)
            return true;

         return false;
      }


      public (double regularHours, double overtimeHours, double doubleTimeHours) CalculateOTHoursBiMonthly(Timesheet timesheet)
      {
         double totalRegularHours = 0;
         double totalOvertimeHours = 0;
         double totalDoubleTimeHours = 0;
         double weeklyRegularLimit = 40.0;
         double dailyRegularLimit = 8.0;
         double dailyOvertimeLimit = 12.0;

         Dictionary<int, double> weeklyTotals = new Dictionary<int, double>(); // Week number -> total hours
         Dictionary<int, int> weeklyWorkDays = new Dictionary<int, int>(); // Week number -> consecutive workdays

         foreach (var timecard in timesheet.timeCards) {
            // Determine the week number relative to the pay period
            int weekNumber = GetWeekNumber(timesheet.periodBegin.Value, timecard.shiftDate.Value);

            // Initialize weekly tracking if not present
            if (!weeklyTotals.ContainsKey(weekNumber)) {
               weeklyTotals[weekNumber] = 0;
               weeklyWorkDays[weekNumber] = 0;
            }

            weeklyWorkDays[weekNumber]++;
            double hoursWorked = timecard.totalHrsActual.TotalHours;

            // Apply 7th Consecutive Day Rule
            if (weeklyWorkDays[weekNumber] == 7) {
               double seventhDayOvertime = Math.Min(hoursWorked, 8.0);
               double seventhDayDoubleTime = Math.Max(0, hoursWorked - 8.0);
               totalOvertimeHours += seventhDayOvertime;
               totalDoubleTimeHours += seventhDayDoubleTime;
               continue;
            }

            double dailyRegularHours = Math.Min(hoursWorked, dailyRegularLimit);
            double dailyOvertimeHours = 0;
            double dailyDoubleTimeHours = 0;

            if (hoursWorked > dailyRegularLimit) {
               dailyOvertimeHours = Math.Min(hoursWorked - dailyRegularLimit, dailyOvertimeLimit - dailyRegularLimit);
               dailyDoubleTimeHours = Math.Max(0, hoursWorked - dailyOvertimeLimit);
            }

            // Weekly Regular Hours Cap
            if (weeklyTotals[weekNumber] < weeklyRegularLimit) {
               double availableWeeklyRegularHours = weeklyRegularLimit - weeklyTotals[weekNumber];
               double regularAllocation = Math.Min(dailyRegularHours, availableWeeklyRegularHours);
               totalRegularHours += regularAllocation;
               dailyRegularHours -= regularAllocation;
            }

            // Remaining daily regular hours go to overtime
            totalOvertimeHours += dailyRegularHours;

            // Add overtime and double time
            totalOvertimeHours += dailyOvertimeHours;
            totalDoubleTimeHours += dailyDoubleTimeHours;

            // Update weekly totals
            weeklyTotals[weekNumber] += hoursWorked;
         }

         return (totalRegularHours, totalOvertimeHours, totalDoubleTimeHours);
      }

      private static int GetWeekNumber(DateTime periodBegin, DateTime date)
      {
         return (int)Math.Floor((date - periodBegin).TotalDays / 7);
      }

      public void AnalyzeADPHours()
      {

         actualHours = new TimeSpan();
         actualOT = new TimeSpan();

         listHours = TimeSpan.FromHours(stub.regHrs);
         listOT = stub.otHrs != 0 ? TimeSpan.FromHours(stub.otHrs) : TimeSpan.FromHours(0);

         regPay = stub.regPay;
         regRate = stub.regRate;

         otPay = stub.otPay;
         otRate = stub.otRate;

         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;

         TimeSpan week = new TimeSpan(0);
         TimeSpan doubleOT = new TimeSpan(0);

         TimeSpan weekHrsLessOT = new TimeSpan(0);

         DateTime midweek = stub.periodBegin.Value.AddDays(6); //up to midweek inclusive
         int daysInWeekOneWrkd = 0;
         int daysInWeekTwoWrked = 0;


         #region Total Workweeks
         foreach (Timecard card in timeCards) {
            if (card.totalHrsActual.TotalHours < 0) {
               this.invalid = true;
               card.invalid = true;
            }

            if (card.shiftDate <= midweek)
               daysInWeekOneWrkd++;
            else if (card.shiftDate > midweek)
               daysInWeekTwoWrked++;

            #region enable auto-deduct comparison
            //  int deductMeal = 0;
            //if (card.timepunches.Count == 2)
            //   deductMeal = 30;// card.listMealLenth;

            //  actualHours += card.totalHrsActual - TimeSpan.FromMinutes(deductMeal);

            //card.totalHrsActual = card.totalHrsActual.Add(TimeSpan.FromMinutes(-deductMeal)); //auto-deduct
            #endregion
         }
         #endregion

         var overtime = CalculateOvertime(timeCards);
         actualOT = TimeSpan.FromHours(overtime.OvertimeHours);
         actualDblOT = TimeSpan.FromHours(overtime.DoubletimeHours);

         #region BI-Monthly
         //var vals = CalculateOTHoursBiMonthly(this);
         //actualOT = TimeSpan.FromHours( vals.overtimeHours);
         //actualDblOT = TimeSpan.FromHours(vals.doubleTimeHours);
         #endregion

         double totalHours = timeCards.Sum(timecard => timecard.totalHrsActual.TotalHours);

         if (daysInWeekOneWrkd > 0)
            totalWorkweeks++;
         if (daysInWeekTwoWrked > 0)
            totalWorkweeks++;
         if (daysInWeekOneWrkd == 7 || daysInWeekTwoWrked == 7)
            workedAllDaysOfaWeek = true;

         //Remove OT from regulars hours
         actualHours = TimeSpan.FromHours(totalHours) - actualOT;
         //add in double OT
         //actualDblOT = doubleOT;

         checkTotalHours = stub.regHrs + stub.otHrs + stub.doubleOtHrs;
         actualTotalHours = actualHours.TotalHours + actualOT.TotalHours;
         double diff = actualTotalHours - listedTotalHours;
      }


      public static OvertimeResult CalculateOvertime(List<Timecard> timecards)
      {
         if (timecards == null || !timecards.Any())
            return new OvertimeResult();

         // Sort timecards by date
         var sortedTimecards = timecards.OrderBy(t => t.shiftDate.Value.Date).ToList();

         var result = new OvertimeResult();

         // Group timecards by workweek (Sunday to Saturday in California)
         // California workweek is defined as 7 consecutive 24-hour periods starting with the same day/time each week
         var workweeks = sortedTimecards
             .GroupBy(t => GetWorkweekStartDate(t.shiftDate.Value.Date))
             .OrderBy(g => g.Key)
             .ToList();

         // Process each workweek separately
         foreach (var workweek in workweeks) {
            double weeklyRegularHours = 0;
            double weeklyOvertimeHours = 0;
            double weeklyDoubletimeHours = 0;

            // Group by day to calculate daily overtime
            var dailyHours = workweek
                .GroupBy(t => t.shiftDate.Value.Date)
                .ToDictionary(g => g.Key, g => g.Sum(t => t.totalHrsActual.TotalHours));

            // Calculate daily overtime and doubletime
            foreach (var day in dailyHours) {
               double hoursForDay = day.Value;

               // Regular hours (up to 8 hours per day)
               double dailyRegularHours = Math.Min(8.0, hoursForDay);
               weeklyRegularHours += dailyRegularHours;
               hoursForDay -= dailyRegularHours;

               // Overtime hours (between 8 and 12 hours per day)
               double dailyOvertimeHours = Math.Min(4.0, hoursForDay);
               weeklyOvertimeHours += dailyOvertimeHours;
               hoursForDay -= dailyOvertimeHours;

               // Doubletime hours (beyond 12 hours per day)
               weeklyDoubletimeHours += hoursForDay;
            }

            // Calculate weekly overtime (over 40 regular hours)
            double adjustedRegularHours = Math.Min(40.0, weeklyRegularHours);
            double weeklyOvertimeFromRegular = weeklyRegularHours - adjustedRegularHours;

            // Add to overall results with weekly adjustment
            result.RegularHours += adjustedRegularHours;
            result.OvertimeHours += weeklyOvertimeHours + weeklyOvertimeFromRegular;
            result.DoubletimeHours += weeklyDoubletimeHours;
         }

         // Check for 7 consecutive days in a row
         CheckConsecutiveDays(sortedTimecards, result);

         return result;
      }



      /// <summary>
      /// Gets the start date of the workweek containing the given date
      /// In California, the workweek is typically defined as Sunday through Saturday
      /// </summary>
      private static DateTime GetWorkweekStartDate(DateTime date)
      {
         // Get previous Sunday
         int daysToSubtract = (int)date.DayOfWeek;
         return date.Date.AddDays(-daysToSubtract);
      }

      /// <summary>
      /// Checks for 7 consecutive days worked and adjusts overtime accordingly
      /// In California, working 7 consecutive days means:
      /// - First 8 hours on the 7th day are overtime (1.5x)
      /// - Hours beyond 8 on the 7th day are doubletime (2x)
      /// </summary>
      private static void CheckConsecutiveDays(List<Timecard> timecards, OvertimeResult result)
      {
         var workDays = timecards
             .Select(t => t.shiftDate.Value.Date)
             .Distinct()
             .OrderBy(d => d)
             .ToList();

         for (int i = 0; i <= workDays.Count - 7; i++) {
            // Check if we have 7 consecutive days
            bool isConsecutive = true;
            for (int j = 0; j < 6; j++) {
               if ((workDays[i + j + 1] - workDays[i + j]).Days != 1) {
                  isConsecutive = false;
                  break;
               }
            }

            if (isConsecutive) {
               // Found 7 consecutive days
               DateTime seventhDay = workDays[i + 6];

               // Get hours worked on 7th day
               var seventhDayTimecards = timecards
                   .Where(t => t.shiftDate.Value.Date == seventhDay)
                   .ToList();

               double seventhDayHours = seventhDayTimecards.Sum(t => t.totalHrsActual.TotalHours);

               // First 8 hours on 7th consecutive day are overtime
               double seventhDayRegularToOT = Math.Min(8.0, seventhDayHours);

               // Hours beyond 8 on 7th consecutive day are doubletime
               double seventhDayOTToDT = Math.Max(0, seventhDayHours - 8.0);

               // Adjust the hours in the result
               // 1. Convert regular hours to overtime for first 8 hours of 7th day
               result.RegularHours -= seventhDayRegularToOT;
               result.OvertimeHours += seventhDayRegularToOT;

               // 2. Convert overtime hours to doubletime for hours beyond 8 on 7th day
               result.OvertimeHours -= seventhDayOTToDT;
               result.DoubletimeHours += seventhDayOTToDT;
            }
         }
      }


      public void CalculateNonStub()
      {
         actualHours = new TimeSpan();
         actualOT = new TimeSpan();

         TimeSpan week = new TimeSpan(0);
         TimeSpan doubleOT = new TimeSpan(0);

         TimeSpan weekHrsLessOT = new TimeSpan(0);

         int daysInArow = 1;
         DateTime previousDay = timeCards.Count > 0 ? timeCards[0].shiftDate.Value : DateTime.Now;


         foreach (Timecard card in timeCards) {

            if (card.shiftDate.Value.AddDays(-1).Day == previousDay.Day)
               daysInArow++;
            else
               daysInArow = 1;

            if (daysInArow >= 7)
               sevenInArow = true;

            previousDay = card.shiftDate.Value;

            if (card.totalHrsActual.TotalHours < 0) {
               // throw new Exception("not possible");
               this.invalid = true;
               card.invalid = true;
               //timeHasWrongEnd++;
            }

            week += card.totalHrsActual;

            int deductMeal = 0;
            if (card.timepunches.Count == 2)
               deductMeal = card.listMealLenth;

            actualHours += card.totalHrsActual - TimeSpan.FromMinutes(deductMeal);
            listedTotalHours += card.regHrsListed + card.otListed + card.dtListed;

            card.totalHrsActual = card.totalHrsActual.Add(TimeSpan.FromMinutes(-deductMeal)); //auto-deduct


            //((daysInWeekOneWrkd >= 7 && daysInWeekTwoWrked == 0 && card.totalHrsActual.TotalHours > 8) || (daysInWeekTwoWrked == 7 && card.totalHrsActual.TotalHours > 8)) {
            if ((daysInArow >= 7 && card.totalHrsActual.TotalHours > 8)) {
               if (card.totalHrsActual.TotalHours < 8) //under 8 hrs on 7th consecutive is regular OT
                  actualOT += card.totalHrsActual;
               else {
                  actualOT += new TimeSpan(8, 0, 0);
                  doubleOT += card.totalHrsActual - new TimeSpan(8, 0, 0); //anything over 8 hrs is dbl ot
               }
            } else if (weekHrsLessOT.TotalHours + card.totalHrsActual.TotalHours > 40) { //anything over 40 hrs is OT
               actualOT += weekHrsLessOT + card.totalHrsActual - new TimeSpan(40, 0, 0);
               weekHrsLessOT = new TimeSpan(40, 0, 0);
               // actualOT += TimeSpan.FromHours(week.TotalHours) - new TimeSpan(40, 0, 0);

               if (card.totalHrsActual.TotalHours > 12) //If more than 12 hours was worked in the day
                  doubleOT += card.totalHrsActual - new TimeSpan(12, 0, 0); //time minus 8 hours
            } else if (card.totalHrsActual.TotalHours > 8) { //anything on a daily basis over  8 is OT
               if (card.totalHrsActual.TotalHours > 8 && card.totalHrsActual.TotalHours <= 12) {//IF more than 8 hours was worked in the day
                  actualOT += card.totalHrsActual - new TimeSpan(8, 0, 0); //time minus 8 hours
                  weekHrsLessOT += new TimeSpan(8, 0, 0);
               } else if (card.totalHrsActual.TotalHours > 12) { //If more than 12 hours was worked in the day
                  doubleOT += card.totalHrsActual - new TimeSpan(12, 0, 0); //time minus 8 hours
                  actualOT += new TimeSpan(4, 0, 0);
                  weekHrsLessOT += new TimeSpan(8, 0, 0);
               }
            } else
               weekHrsLessOT += card.totalHrsActual;


         }
      }

      #region NotLatest
      public void AnalyzeDelunaPost15Hours()
      {
         actualHours = new TimeSpan();
         actualOT = new TimeSpan();

         listHours = TimeSpan.FromHours(stub.regHrs);
         if (stub.otHrs != 0)
            listOT = TimeSpan.FromHours(stub.otHrs);

         regPay = stub.regPay;
         otPay = stub.otPay;
         otRate = stub.otRate;
         regRate = stub.regRate;
         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;

         TimeSpan week = new TimeSpan(0);
         TimeSpan doubleOT = new TimeSpan(0);

         listedTotalHours = stub.regHrs + stub.otHrs;
         foreach (Timecard card in timeCards) {
            week += card.totalHrsActual.TotalHours > 8 ? new TimeSpan(8, 0, 0) : card.totalHrsActual;

            actualHours += card.totalHrsActual;

            //Daily OT analysis
            if (card.totalHrsActual.TotalHours > 8) //IF more than 8 hours was worked in the day
               actualOT += card.totalHrsActual - new TimeSpan(8, 0, 0); //time minus 8 hours

            //TODO: cannot just swap these types of hours
            if (card.totalHrsActual.TotalHours > 12) //IF more than 12 hours was worked in the day
               doubleOT += card.totalHrsActual - new TimeSpan(12, 0, 0); //time minus 8 hours

            // if (card.mealAfter5hrs || card.mealUnder30) //Caclulate meal penalties
            //    mealMissedOrUnder30 = 1;
         }

         //If more than 40 hours was worked in either work week
         if (week.TotalHours > 40) //TODO: determine how to properly to do the over 40 OT
         {
            TimeSpan ot = week - new TimeSpan(40, 0, 0);
            actualOT += ot;
         }

         //Remove OT from regulars hours
         actualHours = actualHours - actualOT;
         //add in double OT
         actualOT += doubleOT;
      }

      public void AnalyzeDelunaHours()
      {
         actualHours = new TimeSpan();
         actualOT = new TimeSpan();

         listHours = TimeSpan.FromHours(stub.regHrs);
         if (stub.otHrs != 0)
            listOT = TimeSpan.FromHours(stub.otHrs);

         regPay = stub.regPay;
         otPay = stub.otPay;
         otRate = stub.otRate;
         regRate = stub.regRate;
         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;

         TimeSpan week = new TimeSpan(0);
         // TimeSpan weekTwo = new TimeSpan(0);
         TimeSpan doubleOT = new TimeSpan(0);

         listedTotalHours = stub.regHrs + stub.otHrs;
         foreach (Timecard card in timeCards) {
            week += card.totalHrsActual.TotalHours > 8 ? new TimeSpan(8, 0, 0) : card.totalHrsActual;

            actualHours += card.totalHrsActual;

            //Daily OT analysis
            if (card.totalHrsActual.TotalHours > 8) //IF more than 8 hours was worked in the day
               actualOT += card.totalHrsActual - new TimeSpan(8, 0, 0); //time minus 8 hours

            //TODO: cannot just swap these types of hours
            if (card.totalHrsActual.TotalHours > 12) //IF more than 12 hours was worked in the day
               doubleOT += card.totalHrsActual - new TimeSpan(12, 0, 0); //time minus 8 hours

            // if (card.mealAfter5hrs || card.mealUnder30) //Caclulate meal penalties
            //     mealMissedOrUnder30 = 1;
         }

         //If more than 40 hours was worked in either work week
         if (week.TotalHours > 40) //TODO: determine how to properly to do the over 40 OT
         {
            TimeSpan ot = week - new TimeSpan(40, 0, 0);
            actualOT += ot;
         }

         //Remove OT from regulars hours
         actualHours = actualHours - actualOT;
         //add in double OT
         actualOT += doubleOT;
      }

      public void AnalyzeHours()
      {
         //TODO: if CARD IS SATURDAY push the hours to OT, this is not standard, have it commented out normally
         actualHours = new TimeSpan();
         actualOT = new TimeSpan();

         listHours = TimeSpan.FromHours(stub.regHrs);
         if (stub.otHrs != 0)
            listOT = TimeSpan.FromHours(stub.otHrs);

         regPay = stub.regPay;
         otPay = stub.otPay;
         otRate = stub.otRate;
         regRate = stub.regRate;
         periodBegin = stub.periodBegin;
         periodEnd = stub.periodEnd;


         TimeSpan weekOne = new TimeSpan(0);
         TimeSpan weekTwo = new TimeSpan(0);
         TimeSpan doubleOT = new TimeSpan(0);

         foreach (Timecard card in timeCards) {
            if (card.identifier == "101687") {
               int x = 0;
            }

            //TODO: check for conssective working days overtime

            //if (card.shiftDate.Value.Date <= midweek.Date)
            //    weekOne += card.totalHrsActual.TotalHours > 8 ? new TimeSpan(8, 0, 0) : card.totalHrsActual;
            //else
            //    weekTwo += card.totalHrsActual.TotalHours > 8 ? new TimeSpan(8, 0, 0) : card.totalHrsActual;

            actualHours += card.totalHrsActual;

            //Daily OT analysis
            if (card.totalHrsActual.TotalHours > 8) //IF more than 8 hours was worked in the day
               actualOT += card.totalHrsActual - new TimeSpan(8, 0, 0); //time minus 8 hours

            //TODO: cannot just swap these types of hours
            if (card.totalHrsActual.TotalHours > 12) //IF more than 12 hours was worked in the day
               doubleOT += card.totalHrsActual - new TimeSpan(12, 0, 0); //time minus 8 hours

         }

         //If more than 40 hours was worked in either work week
         if (weekOne.TotalHours > 40) //TODO: determine how to properly to do the over 40 OT
         {
            TimeSpan ot = weekOne - new TimeSpan(40, 0, 0);
            actualOT += ot;
         }

         if (weekTwo.TotalHours > 40) {
            TimeSpan ot = weekTwo - new TimeSpan(40, 0, 0);
            actualOT += ot;
         }

         //Remove OT from regulars hours
         actualHours = actualHours - actualOT;
         //add in double OT
         actualOT += doubleOT;
      }
      #endregion
   }
}
