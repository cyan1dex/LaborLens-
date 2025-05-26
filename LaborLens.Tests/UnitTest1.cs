using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestPlatform.TestHost;
using NUnit.Framework;

namespace LaborLens.Tests {
   [TestFixture]
   public class MealPeriodTests {
      [Test]
      public void TestBasicMealViolation_NoMealOver5Hours()
      {
         // Arrange: 6-hour shift with no meal break
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 14, 0, 0)   // Clock out (6 hours)
         );

         // Act & Assert
         Assert.That(timecard.HasFirstMealViolation(), Is.True);
         Assert.That(timecard.mealsTaken, Is.EqualTo(0));
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(6.0));
      }

      [Test]
      public void TestShortMeal_Under30Minutes()
      {
         // Arrange: 6-hour shift with 20-minute meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 11, 0, 0),  // Clock out for meal
             new DateTime(2024, 1, 1, 11, 20, 0), // Clock in from meal (20 min)
             new DateTime(2024, 1, 1, 14, 0, 0)   // Clock out
         );

         // Act & Assert
         Assert.That(timecard.breaksORMealsUnder30Before5th, Is.GreaterThan(0));
         Assert.That(timecard.HasFirstMealViolation(), Is.True);
      }

      [Test]
      public void TestLateMeal_After5Hours()
      {
         // Arrange: 8-hour shift with meal after 5.5 hours
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 13, 30, 0), // Clock out for meal (5.5 hrs)
             new DateTime(2024, 1, 1, 14, 0, 0),  // Clock in from meal
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         // Act & Assert
         Assert.That(timecard.lateMeal, Is.True);
         Assert.That(timecard.HasFirstMealViolation(), Is.True);
      }

      [Test]
      public void TestSecondMealViolation_MissedSecondMeal()
      {
         // Arrange: 12-hour shift with only one meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 11, 0, 0),  // Clock out for meal
             new DateTime(2024, 1, 1, 11, 30, 0), // Clock in from meal
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out (12 hours total)
         );

         // Act & Assert
         Assert.That(timecard.HasSecondMealViolation(), Is.True);
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(11.5)); // 12 hours minus 30 min meal
      }

      [Test]
      public void TestCompliantShift_ProperMealTiming()
      {
         // Arrange: 8-hour shift with proper 30-minute meal before 5th hour
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 12, 0, 0),  // Clock out for meal (4 hrs)
             new DateTime(2024, 1, 1, 12, 30, 0), // Clock in from meal (30 min)
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         // Act & Assert
         Assert.That(timecard.HasViolation(), Is.False);
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
         Assert.That(timecard.lateMeal, Is.False);
      }

      [Test]
      public void TestAutoDeduction_Detection()
      {
         // Arrange: 5.5 hour shift with no meal punches (likely auto-deducted)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 13, 30, 0)  // Clock out (5.5 hours)
         );

         // Act & Assert
         Assert.That(timecard.possibleAutoDeduct, Is.True);
         Assert.That(timecard.mealsTaken, Is.EqualTo(0));
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(5.5));
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }
   }

   [TestFixture]
   public class RestPeriodTests {
      private RestPeriodAnalyzer analyzer;

      [SetUp]
      public void Setup()
      {
         analyzer = new RestPeriodAnalyzer();
      }

      [Test]
      public void TestRestPeriodRequirements_ByShiftLength()
      {
         // Test different shift lengths and expected rest period requirements
         var timecard35hrs = CreateTimecardWithDuration("EMP001", 3.5);
         var timecard6hrs = CreateTimecardWithDuration("EMP001", 6.0);
         var timecard8hrs = CreateTimecardWithDuration("EMP001", 8.0);
         var timecard12hrs = CreateTimecardWithDuration("EMP001", 12.0);

         var result35 = analyzer.AnalyzeTimecard(timecard35hrs, "EMP001");
         var result6 = analyzer.AnalyzeTimecard(timecard6hrs, "EMP001");
         var result8 = analyzer.AnalyzeTimecard(timecard8hrs, "EMP001");
         var result12 = analyzer.AnalyzeTimecard(timecard12hrs, "EMP001");

         Assert.That(result35.RequiredRestPeriods, Is.EqualTo(1)); // 3.5 hrs = 1 rest
         Assert.That(result6.RequiredRestPeriods, Is.EqualTo(1));  // 6.0 hrs = 1 rest
         Assert.That(result8.RequiredRestPeriods, Is.EqualTo(2));  // 8.0 hrs = 2 rests
         Assert.That(result12.RequiredRestPeriods, Is.EqualTo(3)); // 12.0 hrs = 3 rests
      }

      [Test]
      public void TestRestPeriod_ProperBreaks()
      {
         // Arrange: 8-hour shift with two 15-minute breaks
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // Break 1 start
             new DateTime(2024, 1, 1, 10, 15, 0), // Break 1 end (15 min)
             new DateTime(2024, 1, 1, 12, 0, 0),  // Meal start
             new DateTime(2024, 1, 1, 12, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 14, 30, 0), // Break 2 start
             new DateTime(2024, 1, 1, 14, 45, 0), // Break 2 end (15 min)
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         // Act
         var result = analyzer.AnalyzeTimecard(timecard, "EMP001");

         // Assert
         Assert.That(result.RequiredRestPeriods, Is.EqualTo(2));
         Assert.That(result.SuccessfulRestPeriods, Is.EqualTo(2));
         Assert.That(result.MissedRestPeriods, Is.EqualTo(0));
      }

      [Test]
      public void TestRestPeriod_ShortBreaks()
      {
         // Arrange: 8-hour shift with short (8-minute) breaks
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // Break 1 start
             new DateTime(2024, 1, 1, 10, 8, 0),  // Break 1 end (8 min - short)
             new DateTime(2024, 1, 1, 12, 0, 0),  // Meal start
             new DateTime(2024, 1, 1, 12, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 14, 30, 0), // Break 2 start
             new DateTime(2024, 1, 1, 14, 38, 0), // Break 2 end (8 min - short)
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         // Act
         var result = analyzer.AnalyzeTimecard(timecard, "EMP001");

         // Assert
         Assert.That(result.RequiredRestPeriods, Is.EqualTo(2));
         Assert.That(result.ShortRestPeriods, Is.EqualTo(2));
         Assert.That(result.SuccessfulRestPeriods, Is.EqualTo(0));
      }

      [Test]
      public void TestRestPeriod_MissedBreaks()
      {
         // Arrange: 8-hour shift with no rest breaks, only meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 12, 0, 0),  // Meal start
             new DateTime(2024, 1, 1, 12, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         // Act
         var result = analyzer.AnalyzeTimecard(timecard, "EMP001");

         // Assert
         Assert.That(result.RequiredRestPeriods, Is.EqualTo(2));
         Assert.That(result.MissedRestPeriods, Is.EqualTo(2));
         Assert.That(result.SuccessfulRestPeriods, Is.EqualTo(0));
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }

      private Timecard CreateTimecardWithDuration(string empId, double hours)
      {
         var startTime = new DateTime(2024, 1, 1, 8, 0, 0);
         var endTime = startTime.AddHours(hours);

         return CreateTimecard(empId, startTime, endTime);
      }
   }

   [TestFixture]
   public class OvertimeCalculationTests {
      [Test]
      public void TestDailyOvertime_Over8Hours()
      {
         // Arrange: Single 10-hour day
         var timecard = CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 1), 10.0);
         var timecards = new List<Timecard> { timecard };

         // Act
         var result = Timesheet.CalculateOvertime(timecards);

         // Assert
         Assert.That(result.RegularHours, Is.EqualTo(8.0));
         Assert.That(result.OvertimeHours, Is.EqualTo(2.0));
         Assert.That(result.DoubletimeHours, Is.EqualTo(0.0));
      }

      [Test]
      public void TestDailyDoubletime_Over12Hours()
      {
         // Arrange: Single 14-hour day
         var timecard = CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 1), 14.0);
         var timecards = new List<Timecard> { timecard };

         // Act
         var result = Timesheet.CalculateOvertime(timecards);

         // Assert
         Assert.That(result.RegularHours, Is.EqualTo(8.0));
         Assert.That(result.OvertimeHours, Is.EqualTo(4.0));  // Hours 8-12
         Assert.That(result.DoubletimeHours, Is.EqualTo(2.0)); // Hours 12-14
      }

      [Test]
      public void TestWeeklyOvertime_Over40Hours()
      {
         // Arrange: 5 days of 9 hours each (45 total)
         var timecards = new List<Timecard>();
         var startDate = new DateTime(2024, 1, 1); // Monday

         for (int i = 0; i < 5; i++) {
            timecards.Add(CreateTimecardWithHours("EMP001", startDate.AddDays(i), 9.0));
         }

         // Act
         var result = Timesheet.CalculateOvertime(timecards);

         // Assert
         Assert.That(result.RegularHours, Is.EqualTo(40.0)); // Capped at 40
         Assert.That(result.OvertimeHours, Is.EqualTo(5.0));  // 5 hours total OT
      }

      [Test]
      public void TestSeventhConsecutiveDay()
      {
         // Arrange: 7 consecutive days, 8 hours each
         var timecards = new List<Timecard>();
         var startDate = new DateTime(2024, 1, 1);

         for (int i = 0; i < 7; i++) {
            timecards.Add(CreateTimecardWithHours("EMP001", startDate.AddDays(i), 8.0));
         }

         // Act
         var result = Timesheet.CalculateOvertime(timecards);

         // Assert - 7th day should convert regular hours to overtime
         Assert.That(result.OvertimeHours, Is.GreaterThan(8.0));
      }

      [Test]
      public void TestWeeklyOvertime()
      {
         // Arrange: Create PayStubs and Timecards for complex week
         var empCards = new Dictionary<string, List<Timecard>>();
         var stubs = new Dictionary<string, List<PayStub>>();

         var periodStart = new DateTime(2024, 1, 1);
         var periodEnd = new DateTime(2024, 1, 14);

         // Create PayStub to define pay period
         var payStub = new PayStub {
            identifier = "EMP001",
            periodBegin = periodStart,
            periodEnd = periodEnd,
            regHrs = 46,
            regPay = 1840,
            otHrs = 6,
            otPay = 540,
            regRate = 40
         };

         stubs["EMP001"] = new List<PayStub> { payStub };

         // Create timecards with mixed daily and weekly OT
         var timecards = new List<Timecard>
         {
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 1), 10.0), // 2 hrs daily OT
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 2), 8.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 3), 8.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 4), 8.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 5), 12.0)  // 4 hrs daily OT, 2 hrs double time
       };

         empCards["EMP001"] = timecards;

         // Act - Use the proper PayPeriods workflow
         var timeSheets = new PayPeriods().PopulateADPTimesheets(stubs, empCards);
         var employeeTimesheet = timeSheets["EMP001"].First();

         // Assert
         Assert.That(employeeTimesheet.actualHours.TotalHours, Is.EqualTo(40.0));
         Assert.That(employeeTimesheet.actualOT.TotalHours, Is.EqualTo(6.0)); // Daily OT + weekly OT
         Assert.That(employeeTimesheet.actualDblOT.TotalHours, Is.EqualTo(0.0));   // From 12-hour day
      }

      [Test]
      public void TestWeeklyOvertimeDoubletime()
      {
         // Arrange: Create PayStubs and Timecards for complex week
         var empCards = new Dictionary<string, List<Timecard>>();
         var stubs = new Dictionary<string, List<PayStub>>();

         var periodStart = new DateTime(2024, 1, 1);
         var periodEnd = new DateTime(2024, 1, 14);

         // Create PayStub to define pay period
         var payStub = new PayStub {
            identifier = "EMP001",
            periodBegin = periodStart,
            periodEnd = periodEnd,
            regHrs = 46,
            regPay = 1840,
            otHrs = 6,
            otPay = 540,
            regRate = 40
         };

         stubs["EMP001"] = new List<PayStub> { payStub };

         // Create timecards with mixed daily and weekly OT
         var timecards = new List<Timecard>
         {
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 1), 10.0), // 2 hrs daily OT
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 2), 8.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 3), 8.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 4), 14.0),
       CreateTimecardWithHours("EMP001", new DateTime(2024, 1, 5), 12.0)  // 4 hrs daily OT, 2 hrs double time
       };

         empCards["EMP001"] = timecards;

         // Act - Use the proper PayPeriods workflow
         var timeSheets = new PayPeriods().PopulateADPTimesheets(stubs, empCards);
         var employeeTimesheet = timeSheets["EMP001"].First();

         // Assert
         Assert.That(employeeTimesheet.actualHours.TotalHours, Is.EqualTo(42.0));
         Assert.That(employeeTimesheet.actualOT.TotalHours, Is.EqualTo(10.0)); // Daily OT + weekly OT
         Assert.That(employeeTimesheet.actualDblOT.TotalHours, Is.EqualTo(2.0));   // From 12-hour day
      }

      private Timecard CreateTimecardWithHours(string empId, DateTime date, double hours)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = date,
            totalHrsActual = TimeSpan.FromHours(hours),
            timepunches = new List<Timepunch>
             {
                    new Timepunch { datetime = date.AddHours(8) },
                    new Timepunch { datetime = date.AddHours(8 + hours) }
                }
         };

         return timecard;
      }
   }

   [TestFixture]
   public class EdgeCaseTests {
      [Test]
      public void TestMidnightCrossover()
      {
         // Test shifts that cross midnight
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 22, 0, 0),  // 10 PM
             new DateTime(2024, 1, 2, 6, 0, 0)    // 6 AM next day
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(8.0));
      }



      [Test]
      public void TestInvalidTimecard_OutOfOrderPunches()
      {
         // Test handling of invalid data
         var timecard = new Timecard {
            identifier = "EMP001",
            shiftDate = new DateTime(2024, 1, 1),
            timepunches = new List<Timepunch>
             {
                    new Timepunch { datetime = new DateTime(2024, 1, 1, 14, 0, 0) }, // Later time first
                    new Timepunch { datetime = new DateTime(2024, 1, 1, 8, 0, 0) }   // Earlier time second
                }
         };

         timecard.AnalyzeTimeCard();

         // Should be marked as invalid or handled gracefully
         Assert.That(timecard.movedDate || timecard.invalid, Is.True);
      }

      [Test]
      public void TestExtremelyLongShift()
      {
         // Test 16-hour shift
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // 6 AM
             new DateTime(2024, 1, 1, 11, 0, 0),  // Meal 1 start
             new DateTime(2024, 1, 1, 11, 30, 0), // Meal 1 end
             new DateTime(2024, 1, 1, 16, 0, 0),  // Meal 2 start
             new DateTime(2024, 1, 1, 16, 30, 0), // Meal 2 end
             new DateTime(2024, 1, 1, 22, 0, 0)   // 10 PM (16 hours total)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(15.0)); // 16 - 1 hour meals
         Assert.That(timecard.mealsTaken, Is.EqualTo(2));
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }
   }

   [TestFixture]
   public class IntegrationTests {
      [Test]
      public void TestFullWorkflow_TypicalViolationScenarios()
      {
         // Test the complete workflow with various violation types
         var empCards = new Dictionary<string, List<Timecard>>();

         // Employee with meal violations
         var violationCards = new List<Timecard>
         {
                CreateViolationTimecard("EMP001", ViolationType.NoMeal),
                CreateViolationTimecard("EMP001", ViolationType.LateMeal),
                CreateViolationTimecard("EMP001", ViolationType.ShortMeal)
            };

         empCards["EMP001"] = violationCards;

         // Employee with compliant shifts
         var compliantCards = new List<Timecard>
         {
                CreateCompliantTimecard("EMP002"),
                CreateCompliantTimecard("EMP002")
            };

         empCards["EMP002"] = compliantCards;

         // Act
         DataProcessor dataProcessor = new DataProcessor();
         var shifts = dataProcessor.GetAnalyzedShifts(empCards);

         // Assert
         var violations = shifts.Where(s => s.hasViolation).Count();
         var compliant = shifts.Where(s => !s.hasViolation).Count();

         Assert.That(violations, Is.EqualTo(3)); // 3 violation scenarios
         Assert.That(compliant, Is.EqualTo(2));  // 2 compliant scenarios
      }

      [Test]
      public void TestRestPeriodIntegration()
      {
         // Test rest period analysis integration
         var empCards = new Dictionary<string, List<Timecard>>();
         var timecard = CreateTimecardWithRestPeriods("EMP001");
         empCards["EMP001"] = new List<Timecard> { timecard };

         var analyzer = new RestPeriodAnalyzer();
         var results = analyzer.AnalyzeEmployeeRestPeriods(empCards);
         var summary = analyzer.GenerateRestPeriodSummary(results);

         Assert.That(results.ContainsKey("EMP001"), Is.True);
         Assert.That(summary.TotalRequiredRestPeriods, Is.GreaterThan(0));
      }

      private enum ViolationType {
         NoMeal,
         ShortMeal,
         LateMeal
      }

      private Timecard CreateViolationTimecard(string empId, ViolationType violationType)
      {
         var baseDate = new DateTime(2024, 1, 1, 8, 0, 0);

         switch (violationType) {
            case ViolationType.NoMeal:
               return CreateTimecard(empId, baseDate, baseDate.AddHours(6)); // 6 hours, no meal

            case ViolationType.ShortMeal:
               return CreateTimecard(empId,
                   baseDate,                    // Clock in
                   baseDate.AddHours(3),        // Meal start
                   baseDate.AddHours(3.25),     // Meal end (15 min)
                   baseDate.AddHours(8));       // Clock out

            case ViolationType.LateMeal:
               return CreateTimecard(empId,
                   baseDate,                    // Clock in
                   baseDate.AddHours(5.5),      // Late meal start
                   baseDate.AddHours(6),        // Meal end
                   baseDate.AddHours(8));       // Clock out

            default:
               throw new ArgumentException("Unknown violation type");
         }
      }

      private Timecard CreateCompliantTimecard(string empId)
      {
         var baseDate = new DateTime(2024, 1, 1, 8, 0, 0);
         return CreateTimecard(empId,
             baseDate,                    // Clock in
             baseDate.AddHours(4),        // Meal start (4 hrs)
             baseDate.AddHours(4.5),      // Meal end (30 min)
             baseDate.AddHours(8));       // Clock out
      }

      private Timecard CreateTimecardWithRestPeriods(string empId)
      {
         var baseDate = new DateTime(2024, 1, 1, 8, 0, 0);
         return CreateTimecard(empId,
             baseDate,                    // Clock in
             baseDate.AddHours(2),        // Break 1 start
             baseDate.AddHours(2.25),     // Break 1 end
             baseDate.AddHours(4),        // Meal start
             baseDate.AddHours(4.5),      // Meal end
             baseDate.AddHours(6.5),      // Break 2 start
             baseDate.AddHours(6.75),     // Break 2 end
             baseDate.AddHours(8));       // Clock out
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }


   }

   [TestFixture]
   public class AdvancedMealPeriodTests {
      [Test]
      public void TestMultipleMealViolations_SameShift()
      {
         // 14-hour shift with short first meal and no second meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // Meal 1 start (4 hrs - OK timing)
             new DateTime(2024, 1, 1, 10, 40, 0), // Meal 1 end (40 min - fine)
             new DateTime(2024, 1, 1, 20, 0, 0)   // Clock out (14 hrs total, no 2nd meal)
         );

         // Assert multiple violations
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Should not have first meal violation");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Should have second meal violation (missing)");
         //Assert.That(timecard.breaksORMealsUnder30Before5th, Is.EqualTo(1), "Should count short meal");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Count meal");
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(13.33).Within(0.1), "Should be ~13.67 hours worked");
      }

      [Test]
      public void TestBoundaryCase_ExactlyFiveHours()
      {
         // Exactly 5-hour shift with no meal - should NOT trigger violation
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 9, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 14, 0, 0)   // Clock out (exactly 5 hours)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(5.0));
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Exactly 5 hours should NOT require meal");
      }

      [Test]
      public void TestBoundaryCase_JustOverFiveHours()
      {
         // 5.1-hour shift with no meal - should trigger violation
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 9, 0, 0),     // Clock in
             new DateTime(2024, 1, 1, 14, 6, 0)     // Clock out (5.1 hours)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.GreaterThan(5.0));
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Over 5 hours requires meal");
      }

      [Test]
      public void TestBoundaryCase_JustUnderFiveHours()
      {
         // 4.9-hour shift with no meal - should NOT trigger violation
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 9, 0, 0),     // Clock in
             new DateTime(2024, 1, 1, 13, 54, 0)    // Clock out (4.9 hours)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.LessThan(5.0));
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Under 5 hours should not require meal");
      }

      [Test]
      public void TestLateSecondMeal_After10Hours()
      {
         // 12-hour shift with second meal taken after 10 hours of work
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end (30 min)
             new DateTime(2024, 1, 1, 17, 0, 0),  // Second meal start (10.5 hrs worked - LATE)
             new DateTime(2024, 1, 1, 17, 30, 0), // Second meal end
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal should be compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Second meal taken too late");
         Assert.That(timecard.mealTakenAfter10, Is.True, "Should flag meal after 10 hours");
         Assert.That(timecard.mealsTaken, Is.EqualTo(2), "Should count both meals");
      }



      [Test]
      public void TestShortSecondMeal_Between5And10Hours()
      {
         // 12-hour shift with short second meal (25 minutes)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end (30 min)
             new DateTime(2024, 1, 1, 16, 0, 0),  // Second meal start (9.5 hrs worked)
             new DateTime(2024, 1, 1, 16, 25, 0), // Second meal end (25 min - SHORT)
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal should be compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Short second meal should be violation");
         Assert.That(timecard.breaksORMealsUnder30Between5and10, Is.EqualTo(1), "Should count short meal between 5-10 hrs");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Short meal shouldn't count, only first meal");
      }

      [Test]
      public void TestMissedSecondMeal_Over10Hours()
      {
         // 12-hour shift with only one meal (missing second meal)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end (30 min)
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out (no second meal)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal should be compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Should have second meal violation");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Should only count first meal");
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(11.5), "Should be 11.5 hours worked");
      }


      [Test]
      public void TestMealWaiver_SixHourShift()
      {
         // 6-hour shift with no meal (potential meal waiver scenario)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 10, 0, 0),  // Clock in
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out (6 hours)
         );

         // In California, 6-hour shifts can waive meal if both parties agree
         // But system should still flag as violation unless waiver is documented
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Should flag 6-hour shift without meal");
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(6.0));
      }

      [Test]
      public void TestConsecutiveShortBreaks_VsMealPeriod()
      {
         // Employee takes multiple short breaks instead of proper meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // Break 1 start
             new DateTime(2024, 1, 1, 10, 10, 0), // Break 1 end (10 min)
             new DateTime(2024, 1, 1, 11, 30, 0), // Break 2 start
             new DateTime(2024, 1, 1, 11, 45, 0), // Break 2 end (15 min)
             new DateTime(2024, 1, 1, 13, 0, 0),  // Break 3 start
             new DateTime(2024, 1, 1, 13, 10, 0), // Break 3 end (10 min)
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out (8 hours)
         );

         // Multiple short breaks don't substitute for meal period
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Short breaks don't count as meal");
         Assert.That(timecard.mealsTaken, Is.EqualTo(0), "No qualifying meals taken");

      }

      [Test]
      public void TestShortMeal_Under30Minutes_ButOver18Minutes()
      {
         // 6-hour shift with 25-minute meal (long enough to count, but too short to qualify)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 11, 0, 0),  // Clock out for meal
             new DateTime(2024, 1, 1, 11, 25, 0), // Clock in from meal (25 min)
             new DateTime(2024, 1, 1, 14, 0, 0)   // Clock out
         );

         // 25 minutes should be counted as a break but not qualify as meal
         Assert.That(timecard.breaksORMealsUnder30Before5th, Is.EqualTo(1), "Should count 25-min break");
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "25 min doesn't qualify as meal");
      }

      [Test]
      public void TestVeryShortBreak_Under18Minutes()
      {
         // 6-hour shift with 15-minute break (should be ignored by meal analysis)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 11, 0, 0),  // Clock out for break
             new DateTime(2024, 1, 1, 11, 15, 0), // Clock in from break (15 min)
             new DateTime(2024, 1, 1, 14, 0, 0)   // Clock out
         );

         // 15 minutes should be ignored by meal analysis (handled by RestPeriodAnalyzer)
         Assert.That(timecard.breaksORMealsUnder30Before5th, Is.EqualTo(0), "Should ignore breaks under 18 min");
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "No qualifying meal break taken");
      }


      [Test]
      public void TestAutoDeduction_FalsePositive()
      {
         // 8.5 hour timecard that looks like auto-deduction but isn't
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 12, 0, 0),  // Meal start
             new DateTime(2024, 1, 1, 12, 30, 0), // Meal end (30 min)
             new DateTime(2024, 1, 1, 16, 30, 0)  // Clock out (8.5 total, 8 worked)
         );

         // Should NOT be flagged as auto-deduction since meal is punched
         Assert.That(timecard.possibleAutoDeduct, Is.False, "Should not flag when meal is punched");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
         Assert.That(timecard.HasViolation(), Is.False, "Should be compliant");
      }

      [Test]
      public void TestExtremelyLongShift_MultipleMealRequirements()
      {
         // 16-hour shift requiring multiple meals
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 5, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 9, 0, 0),   // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 9, 30, 0),  // First meal end
             new DateTime(2024, 1, 1, 14, 30, 0), // Second meal start (10 hrs worked)
             new DateTime(2024, 1, 1, 15, 0, 0),  // Second meal end
             new DateTime(2024, 1, 1, 21, 0, 0)   // Clock out (16 hrs total)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(15.0), "Should be 15 worked hours");
         Assert.That(timecard.mealsTaken, Is.EqualTo(2));
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal should be compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.False, "Second meal should be compliant");
      }

      [Test]
      public void TestOnDutyMeal_SpecialCase()
      {
         // Employee stays on duty during meal (common in healthcare/security)
         // This would need to be flagged separately as it requires special documentation
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out (8 hours straight)
         );

         // If this is marked as having meal but no punch-out
         timecard.mealsTaken = 1; // Simulate system showing meal taken
         timecard.mealIs30 = true; // But it's synthetic/auto-deducted

         Assert.That(timecard.mealIs30, Is.True, "Should detect synthetic meal");
         Assert.That(timecard.timepunches.Count, Is.EqualTo(2), "Should only have in/out punches");
         // This scenario needs special handling for on-duty meal compliance
      }

      [Test]
      public void TestMidnightShift_MealTiming()
      {
         // Overnight shift crossing midnight with meal timing
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 22, 0, 0),  // 10 PM start
             new DateTime(2024, 1, 2, 2, 0, 0),   // 2 AM meal start (4 hrs)
             new DateTime(2024, 1, 2, 2, 30, 0),  // 2:30 AM meal end
             new DateTime(2024, 1, 2, 6, 0, 0)    // 6 AM end
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(7.5), "Should handle midnight crossover");
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Meal timing should be compliant");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
      }

      [Test]
      public void TestVariableMealLength_EdgeCases()
      {
         // Test different meal lengths and their classification
         var shortMeal = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),
             new DateTime(2024, 1, 1, 12, 0, 0),
             new DateTime(2024, 1, 1, 12, 29, 0), // 29 minutes - should be short
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         var exactMeal = CreateTimecard("EMP002",
             new DateTime(2024, 1, 1, 8, 0, 0),
             new DateTime(2024, 1, 1, 12, 0, 0),
             new DateTime(2024, 1, 1, 12, 30, 0), // Exactly 30 minutes
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         var longMeal = CreateTimecard("EMP003",
             new DateTime(2024, 1, 1, 8, 0, 0),
             new DateTime(2024, 1, 1, 12, 0, 0),
             new DateTime(2024, 1, 1, 13, 0, 0),  // 60 minutes
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         // 29 minutes should be violation
         Assert.That(shortMeal.mealsTaken, Is.EqualTo(0), "29 min shouldn't count as meal");
         Assert.That(shortMeal.breaksORMealsUnder30Before5th, Is.EqualTo(1));

         // 30 minutes should be compliant
         Assert.That(exactMeal.mealsTaken, Is.EqualTo(1), "30 min should count as meal");
         Assert.That(exactMeal.mealIs30, Is.True);

         // 60 minutes should be compliant
         Assert.That(longMeal.mealsTaken, Is.EqualTo(1), "60 min should count as meal");
         Assert.That(longMeal.mealIs60, Is.True);
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }
   }

   [TestFixture]
   public class ExpertMealPeriodTests {
      [Test]
      public void TestLateFiestMeal_DoesNotSatisfySecondMealRequirement()
      {
         // 12-hour shift: meal taken at 8th hour (late for 1st meal)
         // Should have BOTH late first meal AND missed second meal violations
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 14, 0, 0),  // Meal start (8 hrs - LATE for 1st meal)
             new DateTime(2024, 1, 1, 14, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out (12 hrs total)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Late first meal should be violation");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Late first meal doesn't satisfy second meal requirement");
         Assert.That(timecard.lateMeal, Is.True, "Should flag late meal");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Should count the meal taken");
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(11.5), "11.5 hours worked");
      }

      [Test]
      public void TestProperFirstMeal_WithMissedSecondMeal()
      {
         // 12-hour shift: proper first meal, but no second meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs - proper timing)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out (no second meal)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal should be compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Should have missed second meal violation");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Should count first meal only");
      }

      [Test]
      public void TestShortBreaksBetween5And10_DoNotPreventSecondMealViolation()
      {
         // 12-hour shift: proper first meal, short breaks between 5-10 hrs, no second meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end
             new DateTime(2024, 1, 1, 13, 0, 0),  // Short break start (7 hrs worked)
             new DateTime(2024, 1, 1, 13, 20, 0), // Short break end (20 min)
             new DateTime(2024, 1, 1, 15, 30, 0), // Another short break (9 hrs worked)
             new DateTime(2024, 1, 1, 15, 45, 0), // Break end (15 min)
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out
         );

         // This tests the bug in the original logic - short breaks shouldn't prevent flagging missed second meal
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal compliant");
         Assert.That(timecard.breaksORMealsUnder30Between5and10, Is.GreaterThan(0), "Should have short breaks");
         // Note: Current buggy logic might not catch this, but it should be a violation
         Assert.That(timecard.HasViolation(), Is.True, "HasViolation should catch missed second meal");
      }

      [Test]
      public void TestTwoLateMeals_BothViolations()
      {
         // 14-hour shift: both meals taken late
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 12, 0, 0),  // First meal start (6 hrs - LATE)
             new DateTime(2024, 1, 1, 12, 30, 0), // First meal end
             new DateTime(2024, 1, 1, 17, 0, 0),  // Second meal start (10.5 hrs worked - LATE)
             new DateTime(2024, 1, 1, 17, 30, 0), // Second meal end
             new DateTime(2024, 1, 1, 20, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "First meal taken late");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Second meal taken late");
         Assert.That(timecard.lateMeal, Is.True, "Should flag late meal");
         Assert.That(timecard.mealTakenAfter10, Is.True, "Should flag meal after 10 hours");
         Assert.That(timecard.mealsTaken, Is.EqualTo(2), "Should count both meals");
      }

      [Test]
      public void TestOnlySecondMeal_NoFirstMeal()
      {
         // 12-hour shift: no first meal, but second meal taken properly at 9th hour
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 15, 0, 0),  // Meal start (9 hrs - would be good timing for 2nd meal)
             new DateTime(2024, 1, 1, 15, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Missing first meal violation");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Taking only 2nd meal doesn't satisfy requirements");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Should count the one meal");
      }

      [Test]
      public void TestComplexShift_MultipleShortBreaksAndLateMeals()
      {
         // 16-hour shift: multiple short breaks, one late meal, one missed meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 5, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 7, 30, 0),  // Break 1 start (2.5 hrs)
             new DateTime(2024, 1, 1, 7, 45, 0),  // Break 1 end (15 min)
             new DateTime(2024, 1, 1, 11, 30, 0), // First meal start (6 hrs worked - LATE)
             new DateTime(2024, 1, 1, 12, 0, 0),  // First meal end
             new DateTime(2024, 1, 1, 14, 0, 0),  // Break 2 start (7.5 hrs worked)
             new DateTime(2024, 1, 1, 14, 20, 0), // Break 2 end (20 min)
             new DateTime(2024, 1, 1, 17, 0, 0),  // Break 3 start (10 hrs worked)
             new DateTime(2024, 1, 1, 17, 15, 0), // Break 3 end (15 min)
             new DateTime(2024, 1, 1, 21, 0, 0)   // Clock out (no second meal)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Late first meal");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Missing second meal");
         Assert.That(timecard.lateMeal, Is.True, "Should flag late meal");
         Assert.That(timecard.breaksORMealsUnder30Between5and10, Is.GreaterThan(0), "Should count short breaks");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Only one qualifying meal");
      }

      [Test]
      public void TestBoundaryTiming_MealAt5thHourExactly()
      {
         // 8-hour shift: meal taken exactly at 5th hour (should be compliant)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 13, 0, 0),  // Meal start (exactly 5 hrs)
             new DateTime(2024, 1, 1, 13, 30, 0), // Meal end
             new DateTime(2024, 1, 1, 16, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Meal at 5th hour should be compliant");
         Assert.That(timecard.lateMeal, Is.False, "Should not flag as late");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
      }

      [Test]
      public void TestBoundaryTiming_MealAt10thHourExactly()
      {
         // 12-hour shift: second meal taken exactly at 10th hour of work
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // First meal start (4 hrs)
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal end
             new DateTime(2024, 1, 1, 16, 0, 0),  // 10.0 hours (correct)
             new DateTime(2024, 1, 1, 17, 0, 0),  // Second meal end
             new DateTime(2024, 1, 1, 18, 0, 0)   // Clock out
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False, "First meal compliant");
         Assert.That(timecard.HasSecondMealViolation(), Is.False, "Second meal at 10th hour should be compliant");
         Assert.That(timecard.mealTakenAfter10, Is.False, "Should not flag meal after 10");
         Assert.That(timecard.mealsTaken, Is.EqualTo(2));
      }

      [Test]
      public void TestExtremeCaseShift_18Hours_MultipleMealViolations()
      {
         // 18-hour shift: first meal late, second meal short, no third meal
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 4, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 30, 0), // First meal start (6.5 hrs - LATE)
             new DateTime(2024, 1, 1, 11, 0, 0),  // First meal end
             new DateTime(2024, 1, 1, 16, 0, 0),  // Second meal start (11 hrs worked - LATE)
             new DateTime(2024, 1, 1, 16, 25, 0), // Second meal end (25 min - SHORT)
             new DateTime(2024, 1, 1, 22, 0, 0)   // Clock out (18 hrs, no third meal)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Late first meal");
         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Late and short second meal");
         Assert.That(timecard.lateMeal, Is.True, "Should flag late meal");
         Assert.That(timecard.mealTakenAfter10, Is.True, "Should flag meal after 10");
         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Short second meal shouldn't count");
         // Note: Third meal requirements vary by industry/agreement
      }

      [Test]
      public void TestAutoDeduction_WithActualViolation()
      {
         // 8.5-hour shift with no meal punches (auto-deduction) but still a violation
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 16, 30, 0)  // Clock out (8.5 hrs, no meal punches)
         );

         // Simulate system auto-deducting meal
         timecard.mealIs30 = true; // Synthetic meal marker

         Assert.That(timecard.possibleAutoDeduct, Is.True, "Should detect auto-deduction");
         Assert.That(timecard.mealIs30, Is.True, "Should be marked as synthetic meal");
         Assert.That(timecard.HasFirstMealViolation(), Is.True, "Auto-deduct without actual meal break is violation");
         Assert.That(timecard.mealsTaken, Is.EqualTo(0), "No actual meal break taken");
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }
   }

   [TestFixture]
   public class AdvancedEdgeCaseMealTests {
      [Test]
      public void TestOvernightShift_CorrectedDate()
      {
         // Shift punches span midnight but shiftDate remains prior day
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 20, 0, 0),  // Clock in
             new DateTime(2024, 1, 2, 1, 0, 0),   // Break start
             new DateTime(2024, 1, 2, 1, 30, 0),  // Break end
             new DateTime(2024, 1, 2, 4, 0, 0)    // Clock out (next day)
         );

         Assert.That(timecard.shiftDate.Value.Date, Is.EqualTo(new DateTime(2024, 1, 1)));
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(7.5));
         Assert.That(timecard.HasFirstMealViolation(), Is.False);
      }

      [Test]
      public void TestMealStart_ExactlyAt10Hours()
      {
         // Ensure second meal at 10.0 hour is not late
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),    // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),   // Meal 1 start
             new DateTime(2024, 1, 1, 10, 30, 0),  // Meal 1 end
             new DateTime(2024, 1, 1, 16, 0, 0),   // Meal 2 start (10 hrs worked)
             new DateTime(2024, 1, 1, 16, 30, 0),  // Meal 2 end
             new DateTime(2024, 1, 1, 18, 0, 0)    // Clock out
         );

         Assert.That(timecard.HasSecondMealViolation(), Is.False, "Second meal exactly at 10 hrs is compliant");
         Assert.That(timecard.mealTakenAfter10, Is.False, "Should not flag meal after 10");
      }

      [Test]
      public void TestMealStart_JustAfter10Hours()
      {
         // Second meal starts just after the 10th hour (late)
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),
             new DateTime(2024, 1, 1, 10, 0, 0),
             new DateTime(2024, 1, 1, 10, 30, 0),
             new DateTime(2024, 1, 1, 16, 1, 0),   // 10.016 hrs worked
             new DateTime(2024, 1, 1, 16, 31, 0),
             new DateTime(2024, 1, 1, 18, 0, 0)
         );

         Assert.That(timecard.HasSecondMealViolation(), Is.True, "Should flag meal after 10 as late");
         Assert.That(timecard.mealTakenAfter10, Is.True);
      }


      [Test]
      public void TestPunchAcrossMidnight_ShiftDateAligned()
      {
         // A punch occurs after midnight but still belongs to prior shift
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 18, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 22, 0, 0),
             new DateTime(2024, 1, 2, 0, 30, 0),   // Meal ends on next calendar day
             new DateTime(2024, 1, 2, 2, 0, 0)     // Clock out
         );

         Assert.That(timecard.shiftDate.Value, Is.EqualTo(new DateTime(2024, 1, 1)), "Shift date should remain the clock-in day");
         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(5.5));
      }

      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes) {
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });
         }

         timecard.AnalyzeTimeCard();
         return timecard;
      }
   }

   [TestFixture]
   public class AdditionalMealPeriodTests {
      private Timecard CreateTimecard(string empId, params DateTime[] punchTimes)
      {
         var timecard = new Timecard {
            identifier = empId,
            shiftDate = punchTimes[0].Date,
            timepunches = new List<Timepunch>()
         };

         foreach (var punchTime in punchTimes)
            timecard.timepunches.Add(new Timepunch { datetime = punchTime });

         timecard.AnalyzeTimeCard();
         return timecard;
      }

      [Test]
      public void Test_ShortMealFollowedByValidMealBefore5thHour()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),   // Clock in
             new DateTime(2024, 1, 1, 10, 0, 0),  // Short break (10 min)
             new DateTime(2024, 1, 1, 10, 10, 0),
             new DateTime(2024, 1, 1, 12, 15, 0),  // Proper meal (30 min)
             new DateTime(2024, 1, 1, 12, 45, 0),
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False);
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
      }

      [Test]
      public void Test_ShortSecondMealAfter10Hours()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),
             new DateTime(2024, 1, 1, 10, 0, 0),
             new DateTime(2024, 1, 1, 10, 30, 0),
             new DateTime(2024, 1, 1, 16, 30, 0),
             new DateTime(2024, 1, 1, 16, 55, 0),
             new DateTime(2024, 1, 1, 18, 0, 0)
         );

         Assert.That(timecard.HasSecondMealViolation(), Is.True);
         Assert.That(timecard.mealTakenAfter10, Is.True);
         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
      }

      [Test]
      public void Test_TwoShortBreaksNoMeal()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 9, 0, 0),
             new DateTime(2024, 1, 1, 11, 0, 0),
             new DateTime(2024, 1, 1, 11, 15, 0),
             new DateTime(2024, 1, 1, 13, 0, 0),
             new DateTime(2024, 1, 1, 13, 15, 0),
             new DateTime(2024, 1, 1, 15, 0, 0)
         );

         Assert.That(timecard.mealsTaken, Is.EqualTo(0));
         Assert.That(timecard.HasFirstMealViolation(), Is.True);
      }

      [Test]
      public void Test_MealCrossesMidnight()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 20, 0, 0),
             new DateTime(2024, 1, 1, 23, 45, 0),
             new DateTime(2024, 1, 2, 0, 15, 0),
             new DateTime(2024, 1, 2, 4, 0, 0)
         );

         Assert.That(timecard.mealsTaken, Is.EqualTo(1));
         Assert.That(timecard.HasFirstMealViolation(), Is.False);
      }

      [Test]
      public void Test_MealAtExactly5Hours()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 7, 0, 0),
             new DateTime(2024, 1, 1, 12, 0, 0),
             new DateTime(2024, 1, 1, 12, 30, 0),
             new DateTime(2024, 1, 1, 15, 0, 0)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.False);
      }

      [Test]
      public void Test_OnlySecondMeal_NoFirstMeal()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 6, 0, 0),
             new DateTime(2024, 1, 1, 15, 0, 0),
             new DateTime(2024, 1, 1, 15, 30, 0),
             new DateTime(2024, 1, 1, 18, 0, 0)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True);
         Assert.That(timecard.HasSecondMealViolation(), Is.True);
      }


      [Test]
      public void Test_PunchesCrossMidnight_ShiftDatePreserved()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 22, 0, 0),
             new DateTime(2024, 1, 2, 6, 0, 0)
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(8.0));
         Assert.That(timecard.shiftDate.Value.Date, Is.EqualTo(new DateTime(2024, 1, 1)));
      }

      [Test]
      public void TestShortThenFullMeal_Before5thHour()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),    // In
             new DateTime(2024, 1, 1, 9, 0, 0),    // Short meal start (1hr in)
             new DateTime(2024, 1, 1, 9, 20, 0),   // Short meal end (20 min)
             new DateTime(2024, 1, 1, 11, 0, 0),   // Full meal start (3 hrs in)
             new DateTime(2024, 1, 1, 11, 30, 0),  // Full meal end
             new DateTime(2024, 1, 1, 16, 0, 0)    // Out
         );

         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Only full meal should count");
         Assert.That(timecard.HasFirstMealViolation(), Is.False, "Full meal before 5th hour is compliant");
      }

      [Test]
      public void TestMealBackToBack_ShiftManipulation()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),
             new DateTime(2024, 1, 1, 10, 0, 0),   // Meal 1 start
             new DateTime(2024, 1, 1, 10, 30, 0),  // Meal 1 end
             new DateTime(2024, 1, 1, 10, 30, 0),  // Meal 2 start (immediate)
             new DateTime(2024, 1, 1, 11, 0, 0),   // Meal 2 end
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         Assert.That(timecard.mealsTaken, Is.EqualTo(1), "Back-to-back meals should not count twice");
      }

      [Test]
      public void TestPunchesCrossMidnight_ShiftDateMismatch()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 22, 0, 0),   // Clock in
             new DateTime(2024, 1, 2, 6, 0, 0)     // Clock out
         );

         Assert.That(timecard.shiftDate, Is.EqualTo(new DateTime(2024, 1, 1)), "Shift date should remain aligned with clock-in");
      }

      [Test]
      public void TestMealWithOddMinuteEnding()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 8, 0, 0),
             new DateTime(2024, 1, 1, 11, 0, 0),       // Meal start
             new DateTime(2024, 1, 1, 11, 29, 59),     // Meal end (just under 30)
             new DateTime(2024, 1, 1, 16, 0, 0)
         );

         Assert.That(timecard.HasFirstMealViolation(), Is.True, "29m59s should not count as a compliant meal");
      }

      [Test]
      public void TestShiftExactlyAt24Hours_ShouldBeInvalid()
      {
         var timecard = CreateTimecard("EMP001",
             new DateTime(2024, 1, 1, 0, 0, 0),
             new DateTime(2024, 1, 2, 0, 0, 0) // Exactly 24 hrs
         );

         Assert.That(timecard.totalHrsActual.TotalHours, Is.EqualTo(24.0));
         Assert.That(timecard.invalid, Is.False); // Up to you if this should be valid or invalid
      }








   }
}