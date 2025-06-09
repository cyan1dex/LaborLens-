using System;
using System.Collections.Generic;
using System.Linq;

namespace LaborLens {
   public enum DayOfWeek { NONE, MO, TU, WE, TH, FR, SA, SU, MON, TUE, WED, THU, FRI, SAT, SUN, MONDAY, TUESDAY, WEDNESDAY, THURSDAY, FRIDAY, SATURDAY, SUNDAY }
   public enum Month { NONE, JAN, FEB, MAR, APR, MAY, JUN, JUL, AUG, SEP, OCT, NOV, DEC }

   public class Timecard {
      public int dayOfMonth;
      public DayOfWeek dayOfWeek;

      public string identifier;
      public string name;
      public string ssn;
      public DateTime? shiftDate;
      public List<Timepunch> timepunches = new List<Timepunch>();

      public TimeSpan totalHrsActual = new TimeSpan(0);
      public double regHrsListed = 0;
      public double otListed = 0;
      public double dtListed = 0;
      public double totalHrsListed = 0;
      public int listMealLenth;

      public TimeSpan mealDuration = new TimeSpan(0);

      public int mealsTaken;
      public bool mealUnder30 = false;

      public bool invalid;
      public bool clockBackIn = true;
      public bool mealIs30 = false;
      public bool mealIs60 = false;
      public bool mealIs30AtTopOfHr = false;

      public bool mealTakenbBtwn5and6 = false;
      public bool mealTakenbtwn6and10 = false;
      public bool mealTakenAfter10 = false;

      public int breaksAndMeals;
      public int breaksORMealsUnder30;
      public int breaksORMealsUnder30Before5th;
      public int breaksORMealsUnder30Between5and10;
      public bool lateMeal = false;

      public bool fullMealBefore5th = false;
      public static int wrongTime = 0;

      public TimeSpan splitShiftLenth;

      public static DateTime earliest = DateTime.MaxValue;
      public static DateTime latest = DateTime.MinValue;

      public int penalties;
      public bool overr;
      public bool possibleAutoDeduct = false;
      public bool paidMealPremium = false;
      public int fiveHourSegments = 0;
      public bool movedDate = false;

      public int restBreaksTaken = 0;
      public int restZero = 0;
      public int restShort = 0;
      public bool firstRest = false;
      public bool secondRest = false;
      public bool firstRestShort = false;
      public bool secondRestShort = false;

      public bool firstMealTaken = false;
      public bool secondMealTaken = false;
      public bool firstMealLate = false;
      public bool secondMealLate = false;



      public void AnalyzeTimeCard()
      {
         if (invalid) return;

         AddDayIfPriorDayIsEarlier();

         for (int i = 0; i < timepunches.Count - 1; i += 2) {
            totalHrsActual += timepunches[i + 1].datetime - timepunches[i].datetime;

            regHrsListed += timepunches[i].hrsListed;
            otListed += timepunches[i].otHrsListed;
            dtListed += timepunches[i].dblOtListed;
         }

         DateTime shiftStart = timepunches[0].datetime;
         DateTime? lastMealEnd = null;

         for (int i = 1; i < timepunches.Count - 2; i += 2) {
            var breakStart = timepunches[i];
            var breakEnd = timepunches[i + 1];
            double mins = (breakEnd.datetime - breakStart.datetime).TotalMinutes;

            // If this break starts immediately after last break ended, skip it
            if (lastMealEnd.HasValue && breakStart.datetime == lastMealEnd.Value) {
               continue;
            }

            TimeSpan breakStartOffset = breakStart.datetime - shiftStart;
            TimeSpan breakEndOffset = breakEnd.datetime - shiftStart;

            if (mins > 900 || mins < 0) {
               invalid = true;
               continue;
            }

            // ------------------------
            // Case: FULL meal (≥ 30)
            // ------------------------
            if (mins >= 30) {
               if (!firstMealTaken) {
                  firstMealTaken = true;
                  if (breakStartOffset.TotalHours > 5) firstMealLate = true;
                  mealsTaken++;
               } else if (!secondMealTaken) {
                  secondMealTaken = true;
                  mealsTaken++;

                  double hoursOffset = breakStartOffset.TotalHours;
                  secondMealLate = breakStartOffset.TotalHours > 10.0;
                  mealTakenAfter10 = breakStartOffset.TotalHours > 10.0;
               }

               lastMealEnd = breakEnd.datetime;

               if (mins == 30) mealIs30 = true;
               if (mins == 60) mealIs60 = true;

               if (breakEnd.datetime.Minute == 0 || breakEnd.datetime.Minute == 15 || breakEnd.datetime.Minute == 30)
                  mealIs30AtTopOfHr = true;

               if (breakStartOffset.TotalHours <= 5)
                  fullMealBefore5th = true;

               if (breakStartOffset.TotalHours > 5 && breakStartOffset.TotalHours <= 6 && mealsTaken == 1)
                  mealTakenbBtwn5and6 = true;
               else if (breakStartOffset.TotalHours > 6 && breakStartOffset.TotalHours <= 10 && mealsTaken == 1)
                  mealTakenbtwn6and10 = true;
            }

            // ------------------------
            // Case: SHORT meal (18–30)
            // ------------------------
            else if (mins >= 18 && mins < 30) {

               if (!firstMealTaken) {
                  breaksORMealsUnder30Before5th++;
               } else if (!secondMealTaken) {
                  breaksORMealsUnder30Between5and10++;
               }
               breaksORMealsUnder30++;

               if (breakStartOffset.TotalHours > 10.0)
                  mealTakenAfter10 = true; // <-- ADD THIS
            }
            // ------------------------
            // Case: Break (Under 15)
            // ------------------------
            else if (mins <= 15) {

               totalHrsActual= totalHrsActual.Add(TimeSpan.FromMinutes(mins));
            }

            #region logging
            //Console.WriteLine("==== TIME CARD ANALYSIS ====");
            //Console.WriteLine($"Identifier: {identifier}");
            //Console.WriteLine($"Total Hrs Actual: {totalHrsActual.TotalHours}");
            //Console.WriteLine($"Meals Taken: {mealsTaken}");
            //Console.WriteLine($"First Meal Taken: {firstMealTaken}, Late: {firstMealLate}");
            //Console.WriteLine($"Second Meal Taken: {secondMealTaken}, Late: {secondMealLate}");
            //Console.WriteLine($"Meal Taken After 10: {mealTakenAfter10}");
            //Console.WriteLine($"Breaks/Meals <30 Between 5–10: {breaksORMealsUnder30Between5and10}");
            //Console.WriteLine($"Late Meal Flag: {lateMeal}");
            //Console.WriteLine($"Possible Auto Deduct: {possibleAutoDeduct}");
            #endregion
         }


         if (mealTakenbBtwn5and6 || mealTakenbtwn6and10 || mealTakenAfter10)
            lateMeal = true;

         if (totalHrsActual.TotalHours >= 5 && mealsTaken == 0)
            possibleAutoDeduct = true;
      }




      public bool HasFirstMealViolation()
      {
         if (totalHrsActual.TotalHours <= 5) return false;
         if (!firstMealTaken || firstMealLate) return true;
         return false;
      }

      public bool HasSecondMealViolation()
      {
         if (totalHrsActual.TotalHours <= 10)
            return false;

         // If no first meal or first was late, second meal required anyway
         if (!firstMealTaken || firstMealLate)
            return true;

         // If second meal was not taken, regardless of short breaks — it's a violation
         if (!secondMealTaken)
            return true;

         // If second meal was taken but late
         if (secondMealLate)
            return true;

         return false;
      }



      public bool HasViolation()
      {
         if ((breaksORMealsUnder30 > 0 && totalHrsActual.TotalHours > 5 && mealsTaken == 0) || lateMeal)
            return true;
         if ((mealsTaken == 0 && totalHrsActual.TotalHours > 5) || (mealsTaken == 1 && totalHrsActual.TotalHours > 10))
            return true;
         return false;
      }

      public bool No30minMealTakenOn5hrShift()
      {
         return mealsTaken == 0 && totalHrsActual.TotalHours > 5;
      }

      public void PopulateShiftDate(DateTime? begin)
      {
         try {
            if (this.dayOfMonth > 0) {
               var temp = new DateTime(begin.Value.Year, begin.Value.Month, this.dayOfMonth);
               if (temp.DayOfWeek.ToString().ToUpper().IndexOf(this.dayOfWeek.ToString().ToUpper()) == 0)
                  shiftDate = temp;
            }
         } catch (Exception) { }
      }

      public void SortTimeCards()
      {
         timepunches = timepunches.OrderBy(o => o.datetime).ToList();
      }

      public void SwapCardsIfBreakIsTooLong()
      {
         bool moved = false;
         for (int i = 0; i < timepunches.Count - 1; i++) {
            if (timepunches[i + 1].datetime < timepunches[i].datetime) {
               timepunches[i + 1].datetime = timepunches[i + 1].datetime.AddDays(1);
               if (!moved) {
                  moved = true;
               }
            }
         }
      }

      public void AddDayIfPriorDayIsEarlier()
      {
         bool moved = false;
         for (int i = 0; i < timepunches.Count - 1; i++) {
            if (timepunches[i + 1].datetime < timepunches[i].datetime) {
               TimeSpan tDiff = timepunches[i + 1].datetime.Subtract(timepunches[i].datetime);
               if (Math.Abs(tDiff.TotalHours) >= 12)
                  timepunches[i + 1].datetime = timepunches[i + 1].datetime.AddHours(24);
               else
                  timepunches[i + 1].datetime = timepunches[i + 1].datetime.AddHours(12);

               if (!moved) {
                  moved = true;
                  movedDate = true;
               }
            }
         }
      }
   }
}