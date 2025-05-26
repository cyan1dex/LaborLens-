using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
  public class Shift {
      public string identifier;
      public DateTime shiftDate;
      public bool shortMeal;
      public bool shortBtwn5n10;
      public bool missedFirstMeal;
      public bool missedSecondMeal;
      public bool lateMeal;
      public bool lateAfter10hrs;

      public bool firstMealViolation;
      public bool secondMealViolation;

      public double shiftLength;
      public bool hasViolation;

      bool invalid;
      public static int totalMeals;
      public static int mealIs30AtTopOfHr;
      public static int splitShiftsOneHour;
      public static int splitShiftsTwoHours;
      public static int mealIs30;
      public static int mealIs60;
      public static int shiftIs8;
      //public static int workWeeks;
      public static HashSet<string> employees = new HashSet<string>();

      public static int over35;
      public int violations;

      public bool endsMod30;
      public bool over;
      public bool meal30;

      public void AnalyzeShift(Timecard card)
      {
         if (card.invalid) this.invalid = true;

         if (card.timepunches.Count > 2)
            if (card.timepunches[2].datetime.Minute % 30 == 0)
               endsMod30 = true;
         if (card.overr)
            this.over = true;

         if (card.totalHrsListed == 0 ? card.totalHrsActual.TotalHours == 8 : card.totalHrsListed == 8)
            Shift.shiftIs8++;

         violations += card.penalties;

         this.identifier = card.identifier;
         // shiftDate = card.shiftDate.Value;

         if (card.splitShiftLenth != null && card.splitShiftLenth.TotalMinutes > 120)
            Shift.splitShiftsTwoHours++;
         if (card.splitShiftLenth != null && card.splitShiftLenth.TotalMinutes > 60)
            Shift.splitShiftsOneHour++;

         Shift.totalMeals += card.mealsTaken > 0 ? 1 : 0;

         if (card.mealIs30) //synthetic meal
            Shift.mealIs30++;
         //if (card.mealIs30AtTopOfHr)
         //   Shift.mealIs30AtTopOfHr++;
         else if (card.mealIs60)
            Shift.mealIs60++;
         bool missedMeal = false;
         bool missed2ndMeal = false;

         //Breaks or meals under 30 needs to be 0, so it does not count when a short meal does, and can not have a late meal either
         if (card.totalHrsActual.TotalHours > 5 && card.mealsTaken == 0 && card.breaksORMealsUnder30Before5th == 0 && !card.lateMeal)
            missedMeal = true;
         if (card.totalHrsActual.TotalHours > 10 && card.mealsTaken == 1 && card.breaksORMealsUnder30Between5and10 == 0)
            missed2ndMeal = true;

         if (card.totalHrsActual.TotalHours > 5 && card.mealsTaken == 0 && card.breaksORMealsUnder30Before5th > 0 && !card.lateMeal)  //short meal before 5th
            shortMeal = true; //short meal before 5th
         else if (missedMeal)
            missedFirstMeal = true; //missed 1st meal

         if (card.totalHrsActual.TotalHours > 10 && card.mealsTaken == 1 && card.breaksORMealsUnder30Between5and10 > 0)//short meal between 5 and 10
            shortBtwn5n10 = true; //short meal between 5 and 10
         else if (missed2ndMeal)
            missedSecondMeal = true; //misses 2nd meal

         lateMeal = (card.firstMealLate && !card.secondMealLate);     
         lateAfter10hrs = card.secondMealLate;                     

         if (lateMeal) {
            int pause = 0;
         }

         if (lateMeal || shortMeal || missedFirstMeal)
            firstMealViolation = true;
         if (lateAfter10hrs || shortBtwn5n10 || missedSecondMeal)
            secondMealViolation = true;

         hasViolation = (firstMealViolation || secondMealViolation);
         shiftLength = card.totalHrsActual.TotalHours;

         if (shiftLength > 3.5)
            over35++;
      }
   }
}
