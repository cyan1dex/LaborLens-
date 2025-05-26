using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
   #region rest
   // Class to represent a break between timepunches
   public class BreakPeriod {
      public DateTime StartTime { get; set; }
      public DateTime EndTime { get; set; }
      public TimeSpan Duration => EndTime - StartTime;
      public bool IsRestPeriod => Duration.TotalMinutes >= 1 && Duration.TotalMinutes <= 20;
      public bool IsShort => IsRestPeriod && Duration.TotalMinutes < 10;
      public bool IsLate { get; set; } // Will be set based on when the break should occur
      public int RestPeriodNumber { get; set; } // 1st, 2nd, or 3rd rest period
   }

   // Class to hold the results of the analysis
   public class RestPeriodAnalysisResult {
      public string EmployeeId { get; set; }
      public DateTime ShiftDate { get; set; }
      public int RequiredRestPeriods { get; set; }
      public int SuccessfulRestPeriods { get; set; }
      public int MissedRestPeriods { get; set; }
      public int ShortRestPeriods { get; set; }
      public int LateRestPeriods { get; set; }
      public List<BreakPeriod> BreakPeriods { get; set; } = new List<BreakPeriod>();
   }

   // Class to hold the overall statistics for all employees
   public class RestPeriodStatistics {
      public int TotalRequiredRestPeriods { get; set; }
      public int TotalSuccessfulRestPeriods { get; set; }
      public int TotalMissedRestPeriods { get; set; }
      public int TotalShortRestPeriods { get; set; }
      public int TotalLateRestPeriods { get; set; }

      // Statistics by rest period number
      public int FirstRestRequired { get; set; }
      public int FirstRestSuccessful { get; set; }
      public int FirstRestMissed { get; set; }
      public int FirstRestShort { get; set; }
      public int FirstRestLate { get; set; }

      public int SecondRestRequired { get; set; }
      public int SecondRestSuccessful { get; set; }
      public int SecondRestMissed { get; set; }
      public int SecondRestShort { get; set; }
      public int SecondRestLate { get; set; }

      public int ThirdRestRequired { get; set; }
      public int ThirdRestSuccessful { get; set; }
      public int ThirdRestMissed { get; set; }
      public int ThirdRestShort { get; set; }
      public int ThirdRestLate { get; set; }
   }

   public class RestPeriodAnalyzer {
      // Method to export shift details to CSV for validation
      private void ExportShiftDetailsToCSV(List<(string EmployeeId, DateTime ShiftDate, int RequiredRests,
                                          bool FirstRestTaken, bool FirstRestSuccessful, bool FirstRestShort, bool FirstRestLate,
                                          bool SecondRestTaken, bool SecondRestSuccessful, bool SecondRestShort, bool SecondRestLate)> shiftDetails,
                                          string filePath)
      {
         try {
            using (var writer = new System.IO.StreamWriter(filePath)) {
               // Write header
               writer.WriteLine("EmployeeId,ShiftDate,RequiredRests," +
                              "FirstRestTaken,FirstRestSuccessful,FirstRestShort,FirstRestLate," +
                              "SecondRestTaken,SecondRestSuccessful,SecondRestShort,SecondRestLate");

               // Write data
               foreach (var detail in shiftDetails) {
                  writer.WriteLine($"{detail.EmployeeId},{detail.ShiftDate:yyyy-MM-dd},{detail.RequiredRests}," +
                                 $"{detail.FirstRestTaken},{detail.FirstRestSuccessful},{detail.FirstRestShort},{detail.FirstRestLate}," +
                                 $"{detail.SecondRestTaken},{detail.SecondRestSuccessful},{detail.SecondRestShort},{detail.SecondRestLate}");
               }
            }

            Console.WriteLine($"Shift details exported to {filePath} for validation");
         } catch (Exception ex) {
            Console.WriteLine($"Error exporting shift details: {ex.Message}");
         }
      }

      // Method to determine how many rest periods are required based on shift length
      private int GetRequiredRestPeriods(TimeSpan shiftDuration)
      {
         double hours = shiftDuration.TotalHours;

         // Using strict boundaries from the chart
         if (hours < 3.5)
            return 0;
         else if (hours >= 3.5 && hours <= 6.0)  // 3.5-6.0 hours = 1 rest
            return 1;
         else if (hours > 6.0 && hours <= 10.0)  // 6.1-10.0 hours = 2 rests
            return 2;
         else if (hours > 10.0 && hours <= 14.0) // 10.1-14.0 hours = 3 rests
            return 3;
         else // 14+ hours = 4 rests
            return 4;
      }

      // Method to analyze a single timecard
      // Method to analyze a single timecard
      // Method to analyze a single timecard - with late check removed
      public RestPeriodAnalysisResult AnalyzeTimecard(Timecard timecard, string employeeId)
      {
         var result = new RestPeriodAnalysisResult();
         result.EmployeeId = employeeId;

         // Ensure we have valid timepunches
         if (timecard?.timepunches == null || timecard.timepunches.Count < 2) {
            return result;
         }

         // Set the shift date
         result.ShiftDate = timecard.timepunches.First().datetime.Date;

         // Identify all breaks in the timecard
         List<BreakPeriod> allBreaks = new List<BreakPeriod>();

         // First find all the break periods (any gap between a clockout and the next clockin)
         for (int i = 1; i < timecard.timepunches.Count; i += 2) {
            // Each break is from an odd-indexed punch (clock-out) to the next even-indexed punch (clock-in)
            if (i + 1 < timecard.timepunches.Count) {
               DateTime clockOut = timecard.timepunches[i].datetime;
               DateTime nextClockIn = timecard.timepunches[i + 1].datetime;

               // Create the break period
               var breakPeriod = new BreakPeriod {
                  StartTime = clockOut,
                  EndTime = nextClockIn
               };

               allBreaks.Add(breakPeriod);
            }
         }

         // Store all breaks in the result
         result.BreakPeriods = allBreaks;

         // Separate rest periods (1-20 minutes) from meal periods
         var restPeriods = allBreaks.Where(b => b.IsRestPeriod).ToList();
         var mealPeriods = allBreaks.Where(b => !b.IsRestPeriod).ToList();

         // Calculate shift duration (from first clock in to last clock out)
         DateTime firstClockIn = timecard.timepunches.First().datetime;
         DateTime lastClockOut = timecard.timepunches.Last().datetime;

         TimeSpan rawShiftDuration = lastClockOut - firstClockIn;

         // Subtract meal periods from shift duration (breaks > 20 minutes)
         TimeSpan adjustedShiftDuration = rawShiftDuration;
         foreach (var meal in mealPeriods) {
            adjustedShiftDuration = adjustedShiftDuration.Subtract(meal.Duration);
         }

         // Now determine required rest periods based on adjusted shift duration
         result.RequiredRestPeriods = GetRequiredRestPeriods(adjustedShiftDuration);

         // Sort rest periods by start time to ensure we're analyzing them in chronological order
         restPeriods = restPeriods.OrderBy(r => r.StartTime).ToList();

         // Count short and successful rest periods
         int shortRests = 0;
         int successfulRests = 0;

         // Process each rest period up to the required number
         for (int i = 0; i < Math.Min(restPeriods.Count, result.RequiredRestPeriods); i++) {
            var rest = restPeriods[i];

            // Only check if short - no late check
            if (rest.IsShort)
               shortRests++;
            else
               successfulRests++;
         }

         // Store the results
         result.SuccessfulRestPeriods = successfulRests;
         result.ShortRestPeriods = shortRests;
         result.LateRestPeriods = 0; // No late check, so always 0

         // Calculate missed rest periods
         result.MissedRestPeriods = result.RequiredRestPeriods - (successfulRests + shortRests);

         return result;
      }


      // Main method to analyze all timecards for all employees
      public Dictionary<string, List<RestPeriodAnalysisResult>> AnalyzeEmployeeRestPeriods(Dictionary<string, List<Timecard>> empCards)
      {
         var results = new Dictionary<string, List<RestPeriodAnalysisResult>>();

         foreach (var empEntry in empCards) {
            string employeeId = empEntry.Key;
            List<Timecard> timecards = empEntry.Value;

            var employeeResults = new List<RestPeriodAnalysisResult>();

            foreach (var timecard in timecards) {
               var result = AnalyzeTimecard(timecard, employeeId);
               employeeResults.Add(result);
            }

            results.Add(employeeId, employeeResults);
         }

         return results;
      }

      // Method to generate a summary report across all employees
      public RestPeriodStatistics GenerateRestPeriodSummary(Dictionary<string, List<RestPeriodAnalysisResult>> analysisResults)
      {
         var stats = new RestPeriodStatistics();

         // Process each employee's results
         foreach (var empResults in analysisResults.Values) {
            foreach (var result in empResults) {
               // Calculate overall stats
               stats.TotalRequiredRestPeriods += result.RequiredRestPeriods;
               stats.TotalSuccessfulRestPeriods += result.SuccessfulRestPeriods;
               stats.TotalMissedRestPeriods += result.MissedRestPeriods;
               stats.TotalShortRestPeriods += result.ShortRestPeriods;
               stats.TotalLateRestPeriods += result.LateRestPeriods;

               // Get all rest periods
               var restPeriods = result.BreakPeriods
                   .Where(b => b.IsRestPeriod)
                   .OrderBy(b => b.StartTime)
                   .ToList();

               // First rest period
               if (result.RequiredRestPeriods >= 1) {
                  stats.FirstRestRequired++;

                  if (restPeriods.Count >= 1) {
                     var firstRest = restPeriods[0];

                     if (firstRest.IsShort)
                        stats.FirstRestShort++;
                     else if (firstRest.IsLate)
                        stats.FirstRestLate++;
                     else
                        stats.FirstRestSuccessful++;
                  } else {
                     stats.FirstRestMissed++;
                  }
               }

               // Second rest period
               if (result.RequiredRestPeriods >= 2) {
                  stats.SecondRestRequired++;

                  if (restPeriods.Count >= 2) {
                     var secondRest = restPeriods[1];

                     if (secondRest.IsShort)
                        stats.SecondRestShort++;
                     else if (secondRest.IsLate)
                        stats.SecondRestLate++;
                     else
                        stats.SecondRestSuccessful++;
                  } else {
                     stats.SecondRestMissed++;
                  }
               }

               // Third rest period
               if (result.RequiredRestPeriods >= 3) {
                  stats.ThirdRestRequired++;

                  if (restPeriods.Count >= 3) {
                     var thirdRest = restPeriods[2];

                     if (thirdRest.IsShort)
                        stats.ThirdRestShort++;
                     else if (thirdRest.IsLate)
                        stats.ThirdRestLate++;
                     else
                        stats.ThirdRestSuccessful++;
                  } else {
                     stats.ThirdRestMissed++;
                  }
               }
            }
         }

         return stats;
      }

   }
   #endregion
}
