using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
   class Analysis {
      HashSet<string> totalPagaEligible = new HashSet<string>(); //from 3/13/18 to present
      HashSet<string> totalLC226Emps = new HashSet<string>();  //from 10/6/16 to present
      HashSet<string> totalCurrentEmps = new HashSet<string>();

      double otPay = 0;
      double otHrs = 0;

      List<Payment> payments = new List<Payment>();
      int totalWorkWeeksGroupA = 0;
      int totalWorkWeeksPAGA = 0;

      Dictionary<string, Period> empPeriodRanges = new Dictionary<string, Period>();
      Dictionary<DateTime, List<string>> periodCounts = new Dictionary<DateTime, List<string>>();

      int terminatedPAGA = 0; //term from present to 3 years back
      int avgEmpsPerPeriod; //avg employees per pay period
      int totalClassPeriods = 0; //total pay periods from 3/13/18 to present
      int totalPagaPeriods = 0; //total pay periods fro 10/6/16 to present
      int totalLLCperiods = 0;
      public int totalEmployeesTimesheets = 0;
      Decimal avgAllYears = 0;

      public DateTime minMealViolPayDate = new DateTime(2100, 12, 1);
      public double paidMealViolationsAmt = 0; // total paid in meal violations
      public double hrsPaidMealViolations = 0;

      public int totalEmployeesPaydata = 0;
      public int totalEmployeesTimedata = 0;
      public int totalShifts;
      public int mealViolations;
      public int totMealViolShifts;
      public int totalWorkweeks;
      public double avgShiftlength;
      public int meal30;
      public int mealsTaken;
      public int shift8;

      public int totalPayPeriods;

      List<Shift> shifts;
      Dictionary<string, List<PayStub>> empStubs;

      public Analysis(List<Shift> shifts, Dictionary<string, List<PayStub>> empStubs)
      {
         this.shifts = shifts;
         this.empStubs = empStubs;
      }


      public void CompleteCalcuations()
      {
         totalShifts = shifts.Count;
         mealViolations = shifts.Where(x => x.shiftLength > 5).Where(x => x.hasViolation).Count();
         totMealViolShifts = shifts.Where(x => x.shiftLength > 5).Count();

         //int mealIs60 = Shift.mealIs60;

         hrsPaidMealViolations = 0;
         paidMealViolationsAmt = 0;


         foreach (KeyValuePair<string, List<PayStub>> es in empStubs) {
            totalEmployeesPaydata++;

            foreach (PayStub stub in es.Value) {

               if (stub.regHrs > 0) {
                  totalPayPeriods++;
               }
            }
         }

         PaymentAnalysis();
      }


      public void UnpaidOvertime(Dictionary<string, List<Timesheet>> empSheets)
      {
         int affectedEmployees = 0;
         int affectedPayPeriods = 0;
         double totalUnpaidOTHours = 0.0;
         double totalUnpaidDTHours = 0.0;

         var employeesWithUnderpayment = new HashSet<string>();
         int totalPeriods = 0;

         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            string employeeId = entry.Key;
            List<Timesheet> timesheets = entry.Value;

            foreach (Timesheet sheet in timesheets) {
               // Get actual values that are already calculated and stored in timesheet
               double actualTotalHours = sheet.actualTotalHours;
               double actualOTHours = Math.Round(sheet.actualOT.TotalHours,1);
               double actualDTHours = Math.Round(sheet.actualDblOT.TotalHours, 1);

               // Get paystub values
               PayStub stub = sheet.stub;
               double stubTotalHours = stub.regHrs + stub.otHrs + stub.doubleOtHrs;
               double stubOTHours = Math.Round(stub.otHrs, 1);
               double stubDTHours = Math.Round(stub.doubleOtHrs, 1);

               // Check if total hours are within 0.25 tolerance
               if (Math.Abs(actualTotalHours - stubTotalHours) <= 0.25) {

                  totalPeriods++;
                  // Calculate underpayment (actual should be >= stub, so underpayment is positive)
                  double otUnderpayment = Math.Max(0, actualOTHours - stubOTHours);
                  double dtUnderpayment = Math.Max(0, actualDTHours - stubDTHours);

                  // If there's any underpayment, count this pay period
                  if (otUnderpayment > 0 || dtUnderpayment > 0) {
                     affectedPayPeriods++;
                     totalUnpaidOTHours += otUnderpayment;
                     totalUnpaidDTHours += dtUnderpayment;

                     // Track this employee (use HashSet to avoid duplicates)
                     employeesWithUnderpayment.Add(employeeId);

                     // Optional: Log details for debugging
                     Console.WriteLine($"Employee {employeeId}, Period {sheet.periodBegin:MM/dd/yyyy}-{sheet.periodEnd:MM/dd/yyyy}:");
                     Console.WriteLine($"  Total Hours - Actual: {actualTotalHours:F2}, Stub: {stubTotalHours:F2}");
                     Console.WriteLine($"  OT Hours - Actual: {actualOTHours:F2}, Stub: {stubOTHours:F2}, Underpaid: {otUnderpayment:F2}");
                     Console.WriteLine($"  DT Hours - Actual: {actualDTHours:F2}, Stub: {stubDTHours:F2}, Underpaid: {dtUnderpayment:F2}");
                     Console.WriteLine();
                  }
               } else {
                  // Optional: Log when total hours don't match
                  Console.WriteLine($"Skipping Employee {employeeId} - Total hours mismatch: Actual {actualTotalHours:F2} vs Stub {stubTotalHours:F2}");
               }
            }
         }

         // Get final count of affected employees
         affectedEmployees = employeesWithUnderpayment.Count;

         // Output final results
         Console.WriteLine("=== UNPAID OVERTIME ANALYSIS RESULTS ===");
         Console.WriteLine($"Affected Employees: {affectedEmployees}");
         Console.WriteLine($"Affected Pay Periods: {affectedPayPeriods}");
         Console.WriteLine($"Total Unpaid Overtime Hours: {totalUnpaidOTHours:F2}");
         Console.WriteLine($"Total Unpaid Double-Time Hours: {totalUnpaidDTHours:F2}");
         Console.WriteLine($"Total Unpaid Premium Hours: {(totalUnpaidOTHours + totalUnpaidDTHours):F2}");

         // Optional: Calculate monetary impact (if you have hourly rates)
         // Console.WriteLine($"Estimated Unpaid OT Pay (at $25/hr base): ${(totalUnpaidOTHours * 25 * 1.5):F2}");
         // Console.WriteLine($"Estimated Unpaid DT Pay (at $25/hr base): ${(totalUnpaidDTHours * 25 * 2.0):F2}");
      }


      public int PeriodAnalysis(Dictionary<string, List<Timesheet>> empSheets)
      {
         int total = 0;
         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            bool workedAfter = false;
            if (entry.Value.Count() > 0)
               totalEmployeesTimesheets++;

            foreach (Timesheet sheet in entry.Value) {
               if (sheet.invalid)
                  continue;
               if (sheet.periodBegin > new DateTime(2018, 1, 1))
                  workedAfter = true;
               if (sheet.periodBegin != null && sheet.periodBegin.Value >= Globals.pagaInitDt && sheet.periodEnd.Value <= Globals.pagaEndDt)
                  totalWorkWeeksPAGA += sheet.GetWorkWeeks();

               #region Employee worked Date Ranges
               if (!empPeriodRanges.ContainsKey(sheet.identifier))
                  empPeriodRanges[sheet.identifier] = new Period(true);

               if (sheet.periodEnd > empPeriodRanges[sheet.identifier].end)
                  empPeriodRanges[sheet.identifier].end = sheet.periodEnd.Value;
               if (sheet.periodEnd < empPeriodRanges[sheet.identifier].start)
                  empPeriodRanges[sheet.identifier].start = sheet.periodEnd.Value;
               #endregion

               #region Period Counts
               if (!periodCounts.ContainsKey(sheet.periodEnd.Value))
                  periodCounts[sheet.periodEnd.Value] = new List<string>();

               //If that employee ID has not been added for that check date already
               if (!periodCounts[sheet.periodEnd.Value].Contains(sheet.identifier))
                  periodCounts[sheet.periodEnd.Value].Add(sheet.identifier);
               #endregion

               #region totalPeriods
               if (sheet.periodEnd.Value >= Globals.classStartDt)
                  totalClassPeriods++;

               if (sheet.periodBegin.Value >= Globals.pagaInitDt && sheet.periodEnd.Value <= Globals.pagaEndDt)
                  totalPagaPeriods++;
               #endregion

               #region Total Employees
               //if (sheet.periodEnd >= Globals.classStartDt)
               //   if (!totalCurrentEmps.Contains(sheet.identifier)) {
               //      totalCurrentEmps.Add(sheet.identifier);
               //   }

               if (sheet.periodEnd >= Globals.pagaInitDt)
                  if (!totalPagaEligible.Contains(sheet.identifier)) {
                     totalPagaEligible.Add(sheet.identifier);
                  }
               #endregion
            }
            if (workedAfter)
               total++;
         }

         #region Terminated Employees      
         //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\kebin\Documents\Timecards\Quezada-Sanchez v Halkat Inc\output.txt", true)) {
         //   foreach (KeyValuePair<string, Period> entry in empPeriodRanges) {
         //      file.WriteLine(String.Format("{0},{1},{2}", entry.Key, entry.Value.start, entry.Value.end));
         //      ////Individual has a paycheck before the last pay period, but not one for the last pay period
         //      if (entry.Value.end >= Globals.pagaInitDt && entry.Value.end < Globals.pagaEndDt)
         //         terminatedPAGA++;

         //   }                 
         //}

         foreach (KeyValuePair<string, Period> entry in empPeriodRanges) {

            ////Individual has a paycheck before the last pay period, but not one for the last pay period
            if (entry.Value.end >= Globals.pagaInitDt && entry.Value.end < Globals.pagaEndDt)
               terminatedPAGA++;

         }
         #endregion

         int totalInperiods = 0;
         foreach (KeyValuePair<DateTime, List<string>> entry in periodCounts) {
            totalInperiods += entry.Value.Count;
         }

         periodCounts.OrderBy(o => o.Key).ToList();
         //avgEmpsPerPeriod = totalInperiods / periodCounts.Count();
         return totalWorkWeeksPAGA;
      }

      public Payment p2013;
      public Payment p2014;
      public Payment p2015;
      public Payment p2016;
      public Payment p2017;
      public Payment p2018;
      public Payment p2019;
      public Payment p2020;
      public Payment p2021;
      public Payment p2022;
      public Payment p2023;
      public Payment p2024;
      public Payment p2025;
      public double otRate;
      public decimal regRate;

      public void PaymentAnalysis()
      {

         p2013 = new Payment();
         p2014 = new Payment();
         p2015 = new Payment();
         p2016 = new Payment();
         p2017 = new Payment();
         p2018 = new Payment();
         p2019 = new Payment();
         p2020 = new Payment();
         p2021 = new Payment();
         p2022 = new Payment();
         p2023 = new Payment();
         p2024 = new Payment();
         p2025 = new Payment();

         payments.Add(p2013);
         payments.Add(p2014);
         payments.Add(p2015);
         payments.Add(p2016);
         payments.Add(p2017);
         payments.Add(p2018);
         payments.Add(p2019);
         payments.Add(p2020);
         payments.Add(p2021);
         payments.Add(p2022);
         payments.Add(p2023);
         payments.Add(p2024);
         payments.Add(p2025);

         foreach (KeyValuePair<string, List<PayStub>> entry in empStubs) {


            foreach (PayStub stub in entry.Value) {
               if (stub.invalid)
                  continue;
               if (stub.regHrs < 0 || stub.regPay < 0 || stub.otHrs < 0 || stub.otPay < 0)
                  continue;
               #region Avg Pay
               // if card is not invalid from bad payment details
               if (stub.regHrs > 0) {
                  if (stub.periodEnd.Value.Year == 2013) { p2013.hrs += (Decimal)stub.regHrs; p2013.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2014) { p2014.hrs += (Decimal)stub.regHrs; p2014.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2015) { p2015.hrs += (Decimal)stub.regHrs; p2015.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2016) { p2016.hrs += (Decimal)stub.regHrs; p2016.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2017) { p2017.hrs += (Decimal)stub.regHrs; p2017.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2018) { p2018.hrs += (Decimal)stub.regHrs; p2018.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2019) { p2019.hrs += (Decimal)stub.regHrs; p2019.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2020) { p2020.hrs += (Decimal)stub.regHrs; p2020.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2021) { p2021.hrs += (Decimal)stub.regHrs; p2021.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2022) { p2022.hrs += (Decimal)stub.regHrs; p2022.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2023) { p2023.hrs += (Decimal)stub.regHrs; p2023.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2024) { p2024.hrs += (Decimal)stub.regHrs; p2024.pay += (Decimal)stub.regPay; };
                  if (stub.periodEnd.Value.Year == 2025) { p2025.hrs += (Decimal)stub.regHrs; p2025.pay += (Decimal)stub.regPay; };
               }

               //Ot with Dbl OT
               //otHrs += stub.otHrs + stub.doubleOtHrs;
               //otPay += stub.otPay + stub.doubleOtPay;

               if (stub.otPay > 0 && stub.otHrs > 0) {
                  otHrs += stub.otHrs;
                  otPay += stub.otPay;
               }

               #endregion

               #region Meal violations
               // meal violation pay
               if (stub.penaltyPay > 0) {
                  hrsPaidMealViolations += stub.penaltyHrs;
                  paidMealViolationsAmt += stub.penaltyPay;

                  if (stub.periodEnd < minMealViolPayDate)
                     minMealViolPayDate = stub.periodBegin.Value;
               }
               #endregion

            }
         }

         Decimal totalPay = 0;
         Decimal totalHrs = 0;
         foreach (Payment p in payments) {

            if (p.hrs > 0) {
               p.rate = p.pay / p.hrs;
               totalPay += p.pay;
               totalHrs += p.hrs;
            }
         }

         if (totalHrs != 0)
            regRate = totalPay / totalHrs;

         otRate = otPay / otHrs;
      }
   }
}
