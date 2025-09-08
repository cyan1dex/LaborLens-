using Microsoft.IdentityModel.Protocols;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlTypes;
using System.Linq;
using static LaborLens.DataProcessor;


namespace LaborLens {
   class ExcelWriter {

      public string currentDir;

      public ExcelWriter()
      {
          currentDir = Path.GetDirectoryName(Path.GetDirectoryName(Path.GetDirectoryName(Directory.GetCurrentDirectory())));
      }

      public void WriteUnpaidBonus(Dictionary<string, List<PayStub>> empStubs)
      {
         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\Desktop\UnpaidBonus.xlsx";

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
         #endregion

         #region vars
         int row = 1;
         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "Period Start Date";
         xlWorkSheet.Cells[row, 3] = "Bonus";
         xlWorkSheet.Cells[row, 4] = "Reg Hrs";
         xlWorkSheet.Cells[row, 5] = "Reg Rate";
         xlWorkSheet.Cells[row, 6] = "Reg Pay";
         xlWorkSheet.Cells[row, 7] = "OT Hrs";
         xlWorkSheet.Cells[row, 8] = "OT Rate";
         xlWorkSheet.Cells[row, 9] = "OT Pay";
         xlWorkSheet.Cells[row, 10] = "Unpaid Bonus";
         row++;
         #endregion

         #region Excel Writing
         foreach (KeyValuePair<string, List<PayStub>> employee in empStubs) {
            foreach (PayStub stub in employee.Value) {
               stub.AnalyzeUnpaidBonus();

               if (stub.unpaidBonusOT > 0) {

                  xlWorkSheet.Cells[row, 1] = stub.identifier;
                  xlWorkSheet.Cells[row, 2] = stub.periodBegin;
                  xlWorkSheet.Cells[row, 3] = stub.bonus;
                  xlWorkSheet.Cells[row, 4] = stub.regHrs;
                  xlWorkSheet.Cells[row, 5] = stub.regRate;
                  xlWorkSheet.Cells[row, 6] = stub.regPay;
                  xlWorkSheet.Cells[row, 7] = stub.otHrs;
                  xlWorkSheet.Cells[row, 8] = stub.otRate;
                  xlWorkSheet.Cells[row, 9] = stub.otPay;
                  xlWorkSheet.Cells[row, 10] = stub.unpaidBonusOT;
                  row++;
               }
            }
         }
         #endregion

         #region Excel Close/Release Doc
         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
         #endregion
      }

      public void PoulateSummaryPayData(Dictionary<string, List<PayStub>> stubs, Analysis salary)
      {
         #region vars
         // int workWeeks = Program.GetTotalWorkweeks(empSheets);
         #endregion

         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx");

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
         #endregion

         salary.PaymentAnalysis(); //Get salary analysis

         double cnt = 0;
         DateTime min = DateTime.MaxValue;
         //var firstPremiumPaid = GetFirstPremiumPeriod(stubs);

         foreach (KeyValuePair<string, List<PayStub>> entry in stubs)
            cnt += entry.Value.Where(x => x.regHrs > 0).Count();

         #region Columns Headers
         int row = 1;
         xlWorkSheet.Cells[3, 3] = stubs.Count; //total emps from pay
         xlWorkSheet.Cells[4, 3] = cnt; //total pay periods from pay

         xlWorkSheet.Cells[7, 3] = salary.p2016.rate;
         xlWorkSheet.Cells[8, 3] = salary.p2017.rate;
         xlWorkSheet.Cells[9, 3] = salary.p2018.rate;
         xlWorkSheet.Cells[10, 3] = salary.p2019.rate;
         xlWorkSheet.Cells[11, 3] = salary.p2020.rate;
         xlWorkSheet.Cells[12, 3] = salary.p2021.rate;
         xlWorkSheet.Cells[13, 3] = salary.p2022.rate;
         xlWorkSheet.Cells[14, 3] = salary.p2023.rate;
         xlWorkSheet.Cells[15, 3] = salary.p2024.rate;
         xlWorkSheet.Cells[16, 3] = salary.p2025.rate;
         xlWorkSheet.Cells[17, 3] = salary.regRate;
         xlWorkSheet.Cells[18, 3] = salary.otRate;
         #endregion


         #region Excel Close/Release Doc
         //xlWorkBook.Save();

         // xlWorkBook.SaveAs(newPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
         xlWorkBook.Close(true, Path.Combine(currentDir, "PaySummary.xlsx") , misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);

         #endregion
      }

      public void PoulateRoster(Dictionary<string, EmployeeDateRanges> dateRanges)
      {
         #region vars
         // int workWeeks = Program.GetTotalWorkweeks(empSheets);
         #endregion

         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx");

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(6);
         #endregion

         int row = 1;
         xlWorkSheet.Cells[row, 1] = "EID";
         xlWorkSheet.Cells[row, 2] = "Shift StartDt";
         xlWorkSheet.Cells[row, 3] = "Shift EndDt";
         xlWorkSheet.Cells[row, 4] = "Wage StartDt";
         xlWorkSheet.Cells[row, 5] = "Wage EndDt";
         row++;

         #region Write Roster
         foreach (KeyValuePair<string, EmployeeDateRanges> entry in dateRanges) {

            xlWorkSheet.Cells[row, 1] = entry.Key;
            xlWorkSheet.Cells[row, 2] = entry.Value.ShiftStart;
            xlWorkSheet.Cells[row, 3] = entry.Value.ShiftEnd;
            xlWorkSheet.Cells[row, 4] = entry.Value.PayStart;
            xlWorkSheet.Cells[row, 5] = entry.Value.PayEnd;
            row++;
         }
         #endregion

         #region Excel Close/Release Doc

         xlWorkBook.Close(true, Path.Combine(currentDir, "Roster.xlsx") , misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);

         #endregion
      }


      public void PoulateGraphData(List<Shift> shifts, int totEmps, PayPeriods pagaData, int over35, Analysis analysis)
      {
        // double avgShift, int totalWorkWeeks
         #region vars
         // int workWeeks = Program.GetTotalWorkweeks(empSheets);
         #endregion

         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx");

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
         #endregion

         #region Columns Headers
         int row = 1;
         xlWorkSheet.Cells[row, 1] = "Shift Length";
         xlWorkSheet.Cells[row, 2] = "Missed 1st";
         xlWorkSheet.Cells[row, 3] = "Late (after 5th)";
         xlWorkSheet.Cells[row, 4] = "Short before 5th";
         xlWorkSheet.Cells[row, 5] = "1st Meal Violation";
         xlWorkSheet.Cells[row, 6] = "2nd Meal Violation";
         xlWorkSheet.Cells[row, 7] = "Total Violations";
         xlWorkSheet.Cells[row, 8] = "Total Shifts";
         xlWorkSheet.Cells[row, 9] = "Violation Rate";
         #endregion

         #region Graph
         row = 2;
         xlWorkSheet.Cells[row, 1] = "Between 5 and 6";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shiftLength <= 6).Count();

         xlWorkSheet.Cells[row, 1] = "Between 6 and 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shiftLength <= 10).Count();

         xlWorkSheet.Cells[row, 1] = "Between 10 and 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shiftLength <= 12).Count();
         row++;

         xlWorkSheet.Cells[row, 1] = "Over 3.5";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 3.5).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 3.5).Count();

         xlWorkSheet.Cells[row, 1] = "Over 5";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 5).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 5).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 5).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 5).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 5).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 5).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 5).Count();

         xlWorkSheet.Cells[row, 1] = "Over 6";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 6).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 6).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 6).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 6).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 6).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 6).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 6).Count();

         xlWorkSheet.Cells[row, 1] = "Over 8";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 8).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 8).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 8).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 8).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 8).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 8).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 8).Count();

         xlWorkSheet.Cells[row, 1] = "Over 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 10).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 10).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 10).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 10).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 10).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 10).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 10).Count();

         xlWorkSheet.Cells[row, 1] = "Over 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength > 12).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength > 12).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength > 12).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength > 12).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength > 12).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength > 12).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength > 12).Count();

         //xlWorkSheet.Cells[row, 1] = "All valid shifts";
         //xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.missedFirstMeal).Count();
         //xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.lateMeal).Count();
         //xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shortMeal).Count();
         //xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.firstMealViolation).Count();
         //xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.secondMealViolation).Count();
         //xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Count();
         row++;


         //long shift latter violations
         xlWorkSheet.Cells[row, 1] = "Shift Length";
         xlWorkSheet.Cells[row, 2] = "Late (after 10th)";
         xlWorkSheet.Cells[row, 3] = "Short between 5 and 10";
         xlWorkSheet.Cells[row++, 4] = "Missed 2nd meal";

         xlWorkSheet.Cells[row, 1] = "Between 10 and 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.missedSecondMeal).Count();

         xlWorkSheet.Cells[row, 1] = "Over 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.missedSecondMeal).Count();

         xlWorkSheet.Cells[row, 1] = "Over 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.missedSecondMeal).Count();
         row += 2;
         #endregion

         #region Tallied Totals
         //xlWorkSheet.Cells[row, 1] = "Total Employees";
         //xlWorkSheet.Cells[row, 2] = "Total Meals";
         //xlWorkSheet.Cells[row, 3] = "Syn. Meal";
         //xlWorkSheet.Cells[row, 4] = "Syn. Meal at @30 interval";
         xlWorkSheet.Cells[row, 5] = "Avg shift length";
         xlWorkSheet.Cells[row, 6] = "Total Workweeks";
         xlWorkSheet.Cells[row++, 7] = "Syn. Meal 60 min";



         // row++;
         xlWorkSheet.Cells[row, 1] = totEmps;
         xlWorkSheet.Cells[row, 2] = Shift.totalMeals;
         xlWorkSheet.Cells[row, 3] = Shift.mealIs30;
         xlWorkSheet.Cells[row, 4] = Shift.shiftIs8;
         xlWorkSheet.Cells[row, 5] = analysis.avgShiftlength;
         xlWorkSheet.Cells[row, 6] = analysis.totalWorkweeks;
         xlWorkSheet.Cells[row, 7] = Shift.mealIs60;
         row += 2;

         xlWorkSheet.Cells[row, 1] = "First Day Worked";
         xlWorkSheet.Cells[row++, 2] = "Last Day Worked";

         xlWorkSheet.Cells[row, 1] = Timecard.earliest.ToShortDateString();
         xlWorkSheet.Cells[row, 2] = Timecard.latest.ToShortDateString();


         xlWorkSheet.Cells[28, 12] = "$"+ analysis.paidMealViolationsAmt.ToString("##,###,###.##");
         xlWorkSheet.Cells[29, 12] = analysis.hrsPaidMealViolations.ToString("###,###");
         xlWorkSheet.Cells[30, 12] = analysis.minMealViolPayDate.ToShortDateString();

         //PAGA numbers
         xlWorkSheet.Cells[38, 1] = pagaData.perW3n5;
         xlWorkSheet.Cells[38, 2] = pagaData.perW5andViol;
         xlWorkSheet.Cells[38, 3] = pagaData.perW5n6Viol;
         xlWorkSheet.Cells[38, 4] = pagaData.perW6andViol;
         xlWorkSheet.Cells[38, 5] = pagaData.perW6n10Viol;
         xlWorkSheet.Cells[38, 6] = pagaData.perW10Viol;
         xlWorkSheet.Cells[38, 7] = pagaData.perW10n12Viol;
         xlWorkSheet.Cells[38, 8] = pagaData.perW12Viol;

         #endregion

         #region Excel Close/Release Doc
         xlWorkBook.Close(true, Path.Combine(currentDir, Program.project + " Analysis") , misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
         #endregion
      }

      public void PoulateGraphData(List<Shift> shifts, Dictionary<string, List<Timesheet>> empSheets)
      {
         #region DateTotals
         Dictionary<DateTime, int> dateTotals = Program.GetPayPeriodTotals(empSheets);
         List<DateTime> keyList = new List<DateTime>(dateTotals.Keys);
         keyList = keyList.OrderBy(x => x.Date).ToList();
         #endregion

         #region vars
         int workWeeks = Program.GetTotalWorkweeks(empSheets);
         #endregion

         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx");

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
         #endregion

         #region Columns Headers
         int row = 1;
         xlWorkSheet.Cells[row, 1] = "Shift Length";
         xlWorkSheet.Cells[row, 2] = "Missed 1st";
         xlWorkSheet.Cells[row, 3] = "Late (after 5th)";
         xlWorkSheet.Cells[row, 4] = "Short before 5th";
         xlWorkSheet.Cells[row, 5] = "First Meal Violation";
         xlWorkSheet.Cells[row, 6] = "Second Meal Violation";
         xlWorkSheet.Cells[row, 7] = "Total Violations";
         xlWorkSheet.Cells[row, 8] = "Total Shifts";
         xlWorkSheet.Cells[row, 9] = "Violation Rate";
         #endregion

         #region Graph
         row = 2;
         xlWorkSheet.Cells[row, 1] = "Between 5 and 6";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shiftLength <= 6).Count();

         xlWorkSheet.Cells[row, 1] = "Between 6 and 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shiftLength <= 10).Count();

         xlWorkSheet.Cells[row, 1] = "Between 10 and 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Count();
         row++;

         xlWorkSheet.Cells[row, 1] = "Over 3.5";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 3.5).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 3.5).Count();

         xlWorkSheet.Cells[row, 1] = "Over 5";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 5).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 5).Count();

         xlWorkSheet.Cells[row, 1] = "Over 6";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 6).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 6).Count();

         xlWorkSheet.Cells[row, 1] = "Over 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 10).Count();

         xlWorkSheet.Cells[row, 1] = "Over 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Where(x => x.shiftLength >= 12).Count();

         xlWorkSheet.Cells[row, 1] = "All valid shifts";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.missedFirstMeal).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.lateMeal).Count();
         xlWorkSheet.Cells[row, 4] = shifts.Where(x => x.shortMeal).Count();
         xlWorkSheet.Cells[row, 5] = shifts.Where(x => x.firstMealViolation).Count();
         xlWorkSheet.Cells[row, 6] = shifts.Where(x => x.secondMealViolation).Count();
         xlWorkSheet.Cells[row, 7] = shifts.Where(x => x.hasViolation).Count();
         xlWorkSheet.Cells[row++, 8] = shifts.Count();
         row++;

         //long shift latter violations
         xlWorkSheet.Cells[row, 1] = "Shift Length";
         xlWorkSheet.Cells[row, 2] = "Late (after 10th)";
         xlWorkSheet.Cells[row, 3] = "Short btwn 5 and 10";
         xlWorkSheet.Cells[row++, 4] = "Missed 2nd meal";

         xlWorkSheet.Cells[row, 1] = "Between 10 and 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shiftLength <= 12).Where(x => x.missedSecondMeal).Count();

         xlWorkSheet.Cells[row, 1] = "Over 10";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 10).Where(x => x.missedSecondMeal).Count();

         xlWorkSheet.Cells[row, 1] = "Over 12";
         xlWorkSheet.Cells[row, 2] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.lateAfter10hrs).Count();
         xlWorkSheet.Cells[row, 3] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.shortBtwn5n10).Count();
         xlWorkSheet.Cells[row++, 4] = shifts.Where(x => x.shiftLength >= 12).Where(x => x.missedSecondMeal).Count();
         row += 2;
         #endregion

         #region Tallied Totals
         xlWorkSheet.Cells[row, 1] = "Total Employees";
         xlWorkSheet.Cells[row, 2] = "Total Meals";
         xlWorkSheet.Cells[row, 3] = "Syn. Meal";
         xlWorkSheet.Cells[row, 4] = "Syn. Meal at @30 interval";
         xlWorkSheet.Cells[row, 5] = "Split shift 1 hr";
         xlWorkSheet.Cells[row, 6] = "Split shift 2 hr";
         xlWorkSheet.Cells[row++, 7] = "Total workweeks";

         xlWorkSheet.Cells[row, 1] = Shift.employees.Count;
         xlWorkSheet.Cells[row, 2] = Shift.totalMeals;
         xlWorkSheet.Cells[row, 3] = Shift.mealIs30;
         xlWorkSheet.Cells[row, 4] = Shift.mealIs30AtTopOfHr;
         xlWorkSheet.Cells[row, 5] = Shift.splitShiftsOneHour;
         xlWorkSheet.Cells[row, 6] = Shift.splitShiftsTwoHours;
         xlWorkSheet.Cells[row++, 7] = Timesheet.totalWorkweeks;
         #endregion

         #region Shift Output/Analysis
         ////row = 1;
         //xlWorkSheet.Cells[row, 12] = "Period Start";
         //xlWorkSheet.Cells[row++, 13] = "Emp Total";
         //foreach (DateTime d in keyList) {
         //   xlWorkSheet.Cells[row, 12] = d.ToShortDateString();
         //   xlWorkSheet.Cells[row++, 13] = dateTotals[d];
         //}
         #endregion


         #region Excel Close/Release Doc
         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
         #endregion
      }

      public void TimecardGraphData(Dictionary<string, List<Timecard>> empObj, Dictionary<string, List<Timesheet>> empSheets)
      {
         #region DateTotals
         Dictionary<DateTime, int> dateTotals = Program.GetPayPeriodTotals(empSheets);
         List<DateTime> keyList = new List<DateTime>(dateTotals.Keys);
         keyList = keyList.OrderBy(x => x.Date).ToList();
         #endregion

         #region vars
         int workWeeks = Program.GetTotalWorkweeks(empSheets);
         int totalMeals = 0;
         int syntheticMeals = 0;
         int splitShiftsTwoHours = 0;
         int splitShiftsOneHour = 0;
         int mealIs30AtTopOfHr = 0;
         #endregion

         #region Excel Doc Creation
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx") ;

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
         #endregion

         #region Columns Headers
         int row = 1;
         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "Shift Date";
         xlWorkSheet.Cells[row, 3] = "Short";
         xlWorkSheet.Cells[row, 4] = "Missed";
         xlWorkSheet.Cells[row, 5] = "Missed 2nd";
         xlWorkSheet.Cells[row, 6] = "Late";
         xlWorkSheet.Cells[row, 7] = "Violation";
         xlWorkSheet.Cells[row, 8] = "Shift Length (hrs)";
         xlWorkSheet.Cells[row, 9] = "Shrt btwn 5 and 10";
         xlWorkSheet.Cells[row, 10] = "Late aftr 10th";

         xlWorkSheet.Cells[row, 12] = "Total Meals";
         xlWorkSheet.Cells[row, 13] = "Syntethic Meals";
         xlWorkSheet.Cells[row, 14] = "splitShift 1hr";
         xlWorkSheet.Cells[row, 15] = "splitShift 2hr";
         xlWorkSheet.Cells[row, 16] = "Syn Meal at 0 or 30";
         xlWorkSheet.Cells[row, 17] = "Total Employees";
         xlWorkSheet.Cells[row, 18] = "Total Workweeks";
         xlWorkSheet.Cells[row, 19] = "Period Start";
         xlWorkSheet.Cells[row, 20] = "Period Quantity";
         #endregion

         #region Period Totals
         row = 2;
         foreach (DateTime d in keyList) {
            xlWorkSheet.Cells[row, 19] = d.ToShortDateString();
            xlWorkSheet.Cells[row++, 20] = dateTotals[d];
         }
         #endregion

         #region Shift Output/Analysis
         row = 2;
         foreach (KeyValuePair<string, List<Timecard>> employee in empObj) {
            foreach (Timecard card in employee.Value) {
               //Only process timecards that are not invalid
               if (card.invalid) continue;

               if (card.splitShiftLenth != null && card.splitShiftLenth.TotalMinutes > 120)
                  splitShiftsTwoHours++;
               if (card.splitShiftLenth != null && card.splitShiftLenth.TotalMinutes > 60)
                  splitShiftsOneHour++;

               totalMeals += card.mealsTaken;

               if (card.mealIs30) //synthetic meal
                  syntheticMeals++;
               if (card.mealIs30AtTopOfHr)
                  mealIs30AtTopOfHr++;

               bool missedMeal = false;
               bool missed2ndMeal = false;

               //Breaks or meals under 30 needs to be 0, so it does not count when a short meal does, and can not have a late meal either
               if (card.totalHrsActual.TotalHours > 5 && card.mealsTaken == 0 && card.breaksORMealsUnder30Before5th == 0 && !card.lateMeal)
                  missedMeal = true;
               if (card.totalHrsActual.TotalHours > 10 && card.mealsTaken == 1 && card.breaksORMealsUnder30Between5and10 == 0)
                  missed2ndMeal = true;

               xlWorkSheet.Cells[row, 1] = card.identifier;
               xlWorkSheet.Cells[row, 2] = card.shiftDate != null ? card.shiftDate.Value.ToShortDateString() : "";// + sheet.beginDate == null ? "" : sheet.beginDate.Value.Year.ToString();

               if (card.totalHrsActual.TotalHours > 5 && card.mealsTaken == 0 && card.breaksORMealsUnder30Before5th > 0 && !card.lateMeal)  //short meal before 5th
                  xlWorkSheet.Cells[row, 3] = 1; //short meal before 5th
               else if (missedMeal)
                  xlWorkSheet.Cells[row, 4] = 1; //missed 1st meal

               if (card.totalHrsActual.TotalHours > 10 && card.mealsTaken == 1 && card.breaksORMealsUnder30Between5and10 > 0)//short meal between 5 and 10
                  xlWorkSheet.Cells[row, 9] = 1; //short meal between 5 and 10
               else if (missed2ndMeal)
                  xlWorkSheet.Cells[row, 5] = 1; //misses 2nd meal

               xlWorkSheet.Cells[row, 6] = (card.lateMeal && !card.mealTakenAfter10) == true ? 1 : 0;
               xlWorkSheet.Cells[row, 10] = card.mealTakenAfter10 == true ? 1 : 0;

               xlWorkSheet.Cells[row, 7] = card.HasViolation() == true ? 1 : 0;
               xlWorkSheet.Cells[row++, 8] = card.totalHrsActual.TotalHours;
            }
         }
         #endregion

         #region Misc Analysis
         xlWorkSheet.Cells[2, 12] = totalMeals;
         xlWorkSheet.Cells[2, 13] = syntheticMeals;
         xlWorkSheet.Cells[2, 14] = splitShiftsOneHour;
         xlWorkSheet.Cells[2, 15] = splitShiftsTwoHours;
         xlWorkSheet.Cells[2, 16] = mealIs30AtTopOfHr;
         xlWorkSheet.Cells[2, 17] = empObj.Count;
         xlWorkSheet.Cells[2, 18] = workWeeks;
         #endregion

         #region Excel Close/Release Doc
         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
         #endregion
      }
      public void WriteRoundingActualVsListed(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                   // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;
         int total = 0;

         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "Pay Period Start";
         xlWorkSheet.Cells[row, 3] = "Pay Period End";
         xlWorkSheet.Cells[row, 4] = "Hrs actually wrkd";
         xlWorkSheet.Cells[row++, 5] = "Hrs listed wrkd";


         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            total++;


            foreach (Timesheet s in employee.Value) {

               if (s.invalid)
                  continue;

               if (Math.Abs(s.actualHours.TotalHours - s.listedTotalHours) > 2)
                  continue;

               int pos = 1;
               xlWorkSheet.Cells[row, pos++] = s.identifier;
               xlWorkSheet.Cells[row, pos++] = s.periodBegin.Value.ToShortDateString();
               xlWorkSheet.Cells[row, pos++] = s.periodEnd.Value.ToShortDateString();
               xlWorkSheet.Cells[row, pos++] = s.actualHours.TotalHours;
               xlWorkSheet.Cells[row++, pos] = s.listedTotalHours;//s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs; //Check hours can be different from the timecard hours

            }
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WriteTimesheetViolationsWithoutStubs(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                   // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;
         int total = 0;

         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "Shift day";
         xlWorkSheet.Cells[row, 3] = "Violation";
         xlWorkSheet.Cells[row, 4] = "Hrs actually wrkd";
         xlWorkSheet.Cells[row, 5] = "Hrs listed wrkd";

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {

            int count = 0;
            total++;
            // if (total > 100)
            //    break;
            double listOT = 0;
            double actOT = 0;
            double listDblOt = 0;
            double actDblOT = 0;


            foreach (Timesheet s in employee.Value) {


               // if (!s.sevenInArow)
               //   continue;
               int autoDeduct = 0;
               count++;

               if (s.invalid)
                  continue;

               // listOT += s.stub.otHrs;
               actOT += s.actualOT.TotalHours;
               actDblOT += s.actualDblOT.TotalHours; ;
               // listDblOt += s.stub.doubleOtHrs;

               double totalHrs = 0;

               foreach (Timecard t in s.timeCards) {
                  if (t.possibleAutoDeduct)
                     autoDeduct++;

                  if (Math.Abs(t.totalHrsActual.TotalHours - t.regHrsListed) > .5)
                     continue;

                  //totalListed += t.totalHrsListed + t.otListed + t.dblOTListed;
                  xlWorkSheet.Cells[row, 1] = t.identifier;
                  xlWorkSheet.Cells[row, 2] = t.shiftDate;
                  if (t.HasViolation()) {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 3];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                     xlWorkSheet.Cells[row, 3] = "Y";
                  } else
                     xlWorkSheet.Cells[row, 3] = "N";

                  // if (t.penalties > 0)
                  xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

                  xlWorkSheet.Cells[row, 5] = t.regHrsListed;
                  // xlWorkSheet.Cells[row, 5] = t.totalHrsListed + t.otListed + t.dblOTListed;

                  totalHrs += t.totalHrsActual.TotalHours;
                  int col = 6;
                  for (int i = 0; i < t.timepunches.Count; i++) {
                     xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                     col++;

                     if (i % 2 == 1 && i + 1 < t.timepunches.Count) {
                        double timeIn = t.timepunches[i].datetime.Subtract(t.timepunches[i - 1].datetime).TotalHours;
                        if (timeIn > 5) {
                           Microsoft.Office.Interop.Excel.Range  cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col - 1];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        double minutesOut = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalMinutes;
                        xlWorkSheet.Cells[row, col] = minutesOut;

                        if (minutesOut < 30) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }
                        col++;
                     }
                  }
                  if (t.mealsTaken == 1 && t.totalHrsActual.TotalHours > 10)//one meal taken over 10 hours
                  {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 10];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                  }
                  row++;
               }

               int pos = 3;

            }
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }
      public void WriteTimesheetViolations(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         //string newPath = @"C:\Users\CYAN1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         string newPath = @"C:\Users\CYAN1\OneDrive\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;
         int total = 0;

         double otUnder = 0;

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            int count = 0;
            total++;
            // if (total > 100)
            //    break;
            double listOT = 0;
            double actOT = 0;
            double listDblOt = 0;
            double actDblOT = 0;


            foreach (Timesheet s in employee.Value) {
               //if (!s.sevenInArow)
               //   continue;
               int autoDeduct = 0;
               count++;
               //if (count > 10)
               //    break;
               if (s.invalid)
                  continue;

               //if (Math.Abs(s.actualTotalHours - (s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs)) > 4)
               //   continue;

               listOT += s.stub.otHrs;
               actOT += s.actualOT.TotalHours;
               actDblOT += s.actualDblOT.TotalHours; ;
               listDblOt += s.stub.doubleOtHrs;


               xlWorkSheet.Cells[row, 1] = "EmpID";
               xlWorkSheet.Cells[row, 2] = "Shift day";
               xlWorkSheet.Cells[row, 3] = "Violation";
               xlWorkSheet.Cells[row, 4] = "Hrs actually wrkd";
               xlWorkSheet.Cells[row, 5] = "Hrs listed wrkd";
               xlWorkSheet.Cells[row, 6] = s.periodBegin.Value.ToShortDateString() + " -- " + s.periodEnd.Value.ToShortDateString();

               row++;
               double totalHrs = 0;

               foreach (Timecard t in s.timeCards) {
                  if (t.possibleAutoDeduct)
                     autoDeduct++;

                  //totalListed += t.totalHrsListed + t.otListed + t.dblOTListed;
                  xlWorkSheet.Cells[row, 1] = t.identifier;
                  xlWorkSheet.Cells[row, 2] = t.shiftDate.Value.ToShortDateString();
                  if (t.HasViolation()) {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 3];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                     xlWorkSheet.Cells[row, 3] = "Y";
                  } else
                     xlWorkSheet.Cells[row, 3] = "N";

                  // if (t.penalties > 0)
                  xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

                  //xlWorkSheet.Cells[row, 5] = t.totalHrsListed;
                  xlWorkSheet.Cells[row, 5] = t.regHrsListed + t.otListed;

                  totalHrs += t.totalHrsActual.TotalHours;
                  int col = 6;
                  for (int i = 0; i < t.timepunches.Count; i++) {
                     xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                     col++;

                     if (i % 2 == 1 && i + 1 < t.timepunches.Count) {
                        double timeIn = t.timepunches[i].datetime.Subtract(t.timepunches[i - 1].datetime).TotalHours;
                        if (timeIn >= 5) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col - 1];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        double minutesOut = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalMinutes;
                        xlWorkSheet.Cells[row, col] = minutesOut;

                        if (minutesOut < 30) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }
                        col++;
                     }
                  }
                  if (t.mealsTaken == 1 && t.totalHrsActual.TotalHours >= 10)//one meal taken over 10 hours
                  {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 10];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                  }
                  row++;
               }
               row++;

               int pos = 3;
               xlWorkSheet.Cells[row - 1, pos] = "Penalties Paid";
               xlWorkSheet.Cells[row, pos++] = s.stub.penaltyHrs;//s.stub.penaltyPay > 0 ? s.stub.penaltyPay / (s.stub.regPay / s.stub.regHrs) : 0;
               Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];

               xlWorkSheet.Cells[row - 1, pos] = "Tot act hrs";
               xlWorkSheet.Cells[row, pos] = s.actualTotalHours; //-- > calc for auto-deduct 5.5 ? s.actualTotalHours -.5 : s.actualTotalHours;// s.actualHours.TotalHours;// - s.actualOT.TotalHours; //Actual hours are hrs worked minus OT
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];
               // cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               //DISPLAY TIME SHEET SUMMARY
               xlWorkSheet.Cells[row - 1, pos] = "Total listed hours";
               xlWorkSheet.Cells[row, pos] = s.listedTotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];

               //DISPLAY TIME SHEET SUMMARY
               xlWorkSheet.Cells[row - 1, pos] = "Tot check hrs";
               xlWorkSheet.Cells[row, pos] = s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs; //Check hours can be different from the timecard hours
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];


               xlWorkSheet.Cells[row - 1, pos] = "OT HRS (CHECK)";
               xlWorkSheet.Cells[row, pos] = s.stub.otHrs; //Check hours can be different from the timecard hours
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];

               xlWorkSheet.Cells[row - 1, pos] = "OT HRS (ACT)";
               xlWorkSheet.Cells[row, pos] = s.actualOT.TotalHours; //Check hours can be different from the timecard hours
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];

               if (s.actualOT.TotalHours > s.stub.doubleOtHrs)
                  otUnder += s.actualOT.TotalHours - s.stub.doubleOtHrs;

               xlWorkSheet.Cells[row - 1, pos] = "DBL OT HRS (CHECK)";
               xlWorkSheet.Cells[row, pos] = s.stub.doubleOtHrs; //Check hours can be different from the timecard hours
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];

               xlWorkSheet.Cells[row - 1, pos] = "DBL OT HRS (ACT)";
               xlWorkSheet.Cells[row, pos] = s.actualDblOT == null ? 0 : s.actualDblOT.TotalHours; //Check hours can be different from the timecard hours
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, pos++];


               //    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               //     xlWorkSheet.Cells[row - 1, 6] = "List OT";
               //     xlWorkSheet.Cells[row, 6] = s.listOT.TotalHours;
               //     cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 6];
               // //    cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

               //     xlWorkSheet.Cells[row - 1, 7] = "Act OT";
               //     xlWorkSheet.Cells[row, 7] = s.actualOT.TotalHours;
               //     cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 7];
               ////     cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
               row += 2;

               //sheetsSeen++;
               //if (sheetsSeen > 20)
               //    break;
            }

            //ot hours HERE
            //using (System.IO.StreamWriter file = new System.IO.StreamWriter(@"C:\Users\CYAN1\Desktop\output.txt", true)) {
            //   file.WriteLine(String.Format("EMP: {0} ListOT {1} ActualOT {2} ListDblOT {3} ActDblOT {4}", employee.Key, listOT, actOT, listDblOt, actDblOT));
            //}
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }
      public void WriteTimesheetViolationsNoSummary(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                   // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;
         int total = 0;


         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "Shift day";
         xlWorkSheet.Cells[row, 3] = "Violation";
         xlWorkSheet.Cells[row, 4] = "Hrs actually wrkd";
         xlWorkSheet.Cells[row, 5] = "Hrs listed wrkd";

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            int count = 0;
            total++;
            // if (total > 100)
            //    break;
            double listOT = 0;
            double actOT = 0;
            double listDblOt = 0;
            double actDblOT = 0;


            foreach (Timesheet s in employee.Value) {
               // if (!s.sevenInArow)
               //   continue;
               int autoDeduct = 0;
               count++;
               //if (count > 10)
               //    break;
               if (s.invalid)
                  continue;

               // if (Math.Abs(s.actualTotalHours - (s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs)) > 4)
               //    continue;

               listOT += s.stub.otHrs;
               actOT += s.actualOT.TotalHours;
               actDblOT += s.actualDblOT.TotalHours; ;
               listDblOt += s.stub.doubleOtHrs;


               // xlWorkSheet.Cells[row, 6] = s.periodBegin.Value.ToShortDateString() + " -- " + s.periodEnd.Value.ToShortDateString();

               row++;
               double totalHrs = 0;

               foreach (Timecard t in s.timeCards) {
                  if (t.possibleAutoDeduct)
                     autoDeduct++;

                  //totalListed += t.totalHrsListed + t.otListed + t.dblOTListed;
                  xlWorkSheet.Cells[row, 1] = t.identifier;
                  xlWorkSheet.Cells[row, 2] = t.shiftDate;
                  if (t.HasViolation()) {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 3];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                     xlWorkSheet.Cells[row, 3] = "Y";
                  } else
                     xlWorkSheet.Cells[row, 3] = "N";

                  // if (t.penalties > 0)
                  xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

                  xlWorkSheet.Cells[row, 5] = t.totalHrsListed;
                  // xlWorkSheet.Cells[row, 5] = t.totalHrsListed + t.otListed + t.dblOTListed;

                  totalHrs += t.totalHrsActual.TotalHours;
                  int col = 6;
                  for (int i = 0; i < t.timepunches.Count; i++) {
                     xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                     col++;

                     if (i % 2 == 1 && i + 1 < t.timepunches.Count) {
                        double timeIn = t.timepunches[i].datetime.Subtract(t.timepunches[i - 1].datetime).TotalHours;
                        if (timeIn > 5) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col - 1];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        double minutesOut = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalMinutes;
                        xlWorkSheet.Cells[row, col] = minutesOut;

                        if (minutesOut < 30) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }
                        col++;
                     }
                  }
                  if (t.mealsTaken == 1 && t.totalHrsActual.TotalHours > 10)//one meal taken over 10 hours
                  {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 10];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                  }
                  row++;
               }


            }

         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WriteEmployeeTimeCards(Dictionary<string, List<Timecard>> timeCards)
      {
         object misValue = System.Reflection.Missing.Value;
         // string newPath = @"C:\Users\kmccracken\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         string newPath = @"C:\Users\CYAN1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);

         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;

         xlWorkSheet.Cells[row, 1] = "EmpID";
         xlWorkSheet.Cells[row, 2] = "DateTime";
         xlWorkSheet.Cells[row, 3] = "Punch Type";
         xlWorkSheet.Cells[row, 4] = "length";

         xlWorkSheet.Cells[row, 6] = "No/short lnch";
         xlWorkSheet.Cells[row, 7] = "Late lnch";
         xlWorkSheet.Cells[row, 8] = "Late brk";
         xlWorkSheet.Cells[row, 9] = "Short brk";

         xlWorkSheet.Cells[row, 11] = "List hrs";
         xlWorkSheet.Cells[row, 12] = "Act hrs";

         foreach (KeyValuePair<string, List<Timecard>> item in timeCards) {
            row++;
            foreach (Timecard t in item.Value) {
               xlWorkSheet.Cells[row, 11] = t.regHrsListed + t.otListed + t.dtListed;
               xlWorkSheet.Cells[row, 12] = t.totalHrsActual.TotalHours;

               foreach (Timepunch punch in t.timepunches) {
                  xlWorkSheet.Cells[row, 1] = t.name;

                  if (punch is TimeIn) {
                     xlWorkSheet.Cells[row, 2] = punch.datetime.ToShortTimeString();
                     //xlWorkSheet.Cells[row, 3] = punch.payCode.code.ToString();
                  }

                  //xlWorkSheet.Cells[row, 4] = punch.clockType.length.TotalMinutes;
                  row++;
               }

               row++;
            }

            if (row == 1000)
               break;
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WriteTimecardsFlat(Dictionary<string, List<Timecard>> timeCards)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\OneDrive\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                            // string newPath = @"C:\Users\kebin\Documents\Timecards\Marin v Community Convalescent\Input\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                            // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;

         foreach (KeyValuePair<string, List<Timecard>> personCard in timeCards) {
            //if (row > 1600)
            //   break;

            row++;

            foreach (Timecard t in personCard.Value) {
               // if (t.invalid == true || row == 2134)
               //    continue;
               xlWorkSheet.Cells[row, 1] = t.identifier;
               xlWorkSheet.Cells[row, 2] = t.shiftDate.Value.ToShortDateString();

               if (t.HasViolation() == true) {//Meal violation indicator
                  xlWorkSheet.Cells[row, 3] = "Meal Viol";
                  Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 3];
                  cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
               }

               if (t.timepunches.Count % 2 != 0)
                  continue;

               xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

               int col = 5;
               for (int i = 0; i < t.timepunches.Count; i += 2) {

                  xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                  col++;
                  xlWorkSheet.Cells[row, col] = t.timepunches[i + 1].datetime;

                  double timeIn = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalHours;
                  if (timeIn > 5 && i == 0) {
                     Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                     cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
                  }
                  col++;


                  if (t.timepunches.Count > i + 2) {
                     double minutesOut = t.timepunches[i + 2].datetime.Subtract(t.timepunches[i + 1].datetime).TotalMinutes;
                     xlWorkSheet.Cells[row, col] = minutesOut;

                     if (minutesOut < 30) {
                        Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                        cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                     }
                  }

                  col++;
                  //
               }
               row++;
            }


         }

         xlWorkBook.Close(true, misValue, misValue);
         // xlWorkBook.SaveAs(newPath, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false, XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);

         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WritePayDetails(Dictionary<string, List<PayStub>> stubs)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = Path.Combine(currentDir, "CardGraph.xlsx");
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(3);

         int row = 1;

         xlWorkSheet.Cells[row, 1] = "EID";
         xlWorkSheet.Cells[row, 2] = "Reg Rate";
         xlWorkSheet.Cells[row, 3] = "Reg Hrs";
         xlWorkSheet.Cells[row, 4] = "OT Rate";
         xlWorkSheet.Cells[row, 5] = "OT Hrs";
         xlWorkSheet.Cells[row, 6] = "DT Rate";
         xlWorkSheet.Cells[row, 7] = "DT Hrs";
         xlWorkSheet.Cells[row, 8] = "Meal Hrs";
         xlWorkSheet.Cells[row, 9] = "Meal Pay";

         foreach (KeyValuePair<string, List<PayStub>> entry in stubs) {
            row++;
            xlWorkSheet.Cells[row, 1] = entry.Key;

            List<Payment> payments = new List<Payment>();
            Payment reg = new Payment();
            Payment ot = new Payment();
            Payment dt = new Payment();
            Payment meal = new Payment();

            payments.Add(reg);
            payments.Add(ot);
            payments.Add(dt);
            payments.Add(meal);

            foreach (PayStub s in entry.Value) {
               if (s.regHrs > 0) { reg.hrs += (Decimal)s.regHrs; reg.pay += (Decimal)s.regPay; }
               if (s.otHrs > 0) { ot.hrs += (Decimal)s.otHrs; ot.pay += (Decimal)s.otPay; }
               if (s.doubleOtHrs > 0) { dt.hrs += (Decimal)s.doubleOtHrs; dt.pay += (Decimal)s.doubleOtPay; }
               if (s.penaltyHrs > 0) { meal.hrs += (Decimal)s.penaltyHrs; meal.pay += (Decimal)s.penaltyPay; }
            }

            for (int pos = 0; pos < 3; pos++) {
               if (payments[pos].hrs > 0) {
                  xlWorkSheet.Cells[row, 2 + pos * 2] = payments[pos].pay / payments[pos].hrs;
                  xlWorkSheet.Cells[row, 3 + pos * 2] = payments[pos].hrs;
               }
            }
            xlWorkSheet.Cells[row, 8] = payments[3].hrs;//meal hrs
            xlWorkSheet.Cells[row, 9] = payments[3].pay;//meal pay
         }

         xlWorkBook.Close(true, Path.Combine(currentDir, "PayDetails.xlsx") , misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public static string GetColumnNameFromNumber(int columnNumber)
      {
         string columnName = "";

         while (columnNumber > 0) {
            int remainder = columnNumber % 26;
            if (remainder == 0) {
               columnNumber--;
               remainder = 26;
            }
            columnName = (char)(65 + remainder - 1) + columnName;
            columnNumber = (columnNumber - 1) / 26;
         }

         return columnName;
      }

      public void WritePagaViolations(Dictionary<string, List<Timesheet>> empSheets, List<Period> periods)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\CYAN1\OneDrive\Desktop\PAGA.xlsx";
         // string newPath = @"C:\Users\kebin\Documents\Timecards\Quezada-Sanchez v Halkat Inc\schedule.xlsx";
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);


         Dictionary<DateTime, int> columnPeriods = new Dictionary<DateTime, int>();
         int totalPeriods = 0;
         int row = 1;

         for (int col = 0; col < periods.Count; col++) {//start dates
            xlWorkSheet.Cells[row, col + 2] = periods[col].start.ToShortDateString();
            columnPeriods[periods[col].start] = 0;
            columnPeriods[periods[col].start] = col;

            Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col + 2];
            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
         }
         //columnPeriods[new DateTime(2018, 5, 28)] = 78;

         row = 2;
         //end dates
         for (int col = 0; col < periods.Count; col++) {
            Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col + 2];
            cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            xlWorkSheet.Cells[row, col + 2] = periods[col].end.ToShortDateString();
         }

         row += 2;

         for (int pos = 0; pos < empSheets.Count; pos++) {
            int violation = 100;

            var employee = empSheets.ElementAt(pos).Value;
            employee = employee.OrderBy(x => x.periodBegin).ToList();


            foreach (Timesheet s in employee) {
               if (s.periodBegin < periods[0].start) //Move to initial period
                  continue;

               totalPeriods++;

               xlWorkSheet.Cells[row, 1] = s.identifier; //write employee ID
               Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 1];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGoldenrodYellow);

               if (columnPeriods[s.periodBegin.Value] == 0) {
                  //int pause = 0;
               }
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value] + 2];

               if (s.HasFirstMealViolation() || s.HasSecondMealViolation()) {
                  xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value] + 2] = violation;
                  if (violation == 100)
                     violation = 200;

                  //if (s.PaidPenalties() == true) {
                  //   // if (s.HasSecondMealViolation()) {
                  //   xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value] + 2] = violation;
                  //}

                  cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
               } else {
                  xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value] + 2] = 0;
                  cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
               }
            }
            row++;
         }

         int totalCols = periods.Count + 1;
         string letterCol = GetColumnNameFromNumber(totalCols);

         for (row = 4; row < empSheets.Count + 4; row++) {
            Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, totalCols + 1];

            //Violation Amounts
            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.CornflowerBlue);
            cell.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
            cell.Borders[XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
            cell.Borders[XlBordersIndex.xlEdgeLeft].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic; // Black by default
            cell.NumberFormat = "$#,##0.00";
            cell.Formula = "=SUM(B" + row + ":" + letterCol + row + ")";

            //Non-penalties
            cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, totalCols + 2];
            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            cell.Formula = "=COUNTIF(B" + row + ":" + letterCol + row + ", \"=0\")";

            //Penalties
            cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, totalCols + 3];
            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            cell.Formula = "=COUNTIF(B" + row + ":" + letterCol + row + ", \"=200\") +" + " COUNTIF(B" + row + ":" + letterCol + row + ", \"=100\")";

            //Violation Rates
            int nonpenaltyCol = totalCols + 2;
            int penaltyCol = nonpenaltyCol + 1;
            string nonpenaltyColName = GetColumnNameFromNumber(nonpenaltyCol);
            string penaltyColName = GetColumnNameFromNumber(penaltyCol);
            cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, totalCols + 4];
            cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            cell.NumberFormat = "0.00%";
            cell.Formula = string.Format("=IF({0}{2} = 0, 0, {0}{2}/({0}{2}+{1}{2}))", penaltyColName, nonpenaltyColName, row);
         }

         #region lables
         Microsoft.Office.Interop.Excel.Range cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[3, totalCols + 1];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.Font.Bold = true;
         cellB.Formula = "Violation Amount";

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[3, totalCols + 2];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.Font.Bold = true;
         cellB.Formula = "Non-Penalty Mo.";

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[3, totalCols + 3];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.Font.Bold = true;
         cellB.Formula = "Penalty Mo.";

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[3, totalCols + 4];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.Font.Bold = true;
         cellB.Formula = "Violation %";
         #endregion

         #region Sums
         //Violation Total
         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, totalCols + 1];
         string colName = GetColumnNameFromNumber(totalCols + 1);
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSeaGreen);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.NumberFormat = "$#,##0.00";
         cellB.Formula = string.Format("=SUM({0}$4:{0}${1}", colName, empSheets.Count + 4);
         cellB.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeLeft].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic; // Black by default
         cellB.Borders[XlBordersIndex.xlEdgeTop].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         //Non-penalty total
         string colNameNonPen = GetColumnNameFromNumber(totalCols + 2);
         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, totalCols + 2];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSeaGreen);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.NumberFormat = "#,##0";
         cellB.Formula = string.Format("=SUM({0}$4:{0}${1}", colNameNonPen, empSheets.Count + 4);
         cellB.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeTop].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         //Penalty total
         string colNamePen = GetColumnNameFromNumber(totalCols + 3);
         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, totalCols + 3];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSeaGreen);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.NumberFormat = "#,##0";
         cellB.Formula = string.Format("=SUM({0}$4:{0}${1}", colNamePen, empSheets.Count + 4);
         cellB.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeTop].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, totalCols + 4];
         colName = GetColumnNameFromNumber(totalCols + 4);
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSeaGreen);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.NumberFormat = "0.00%";
         cellB.Formula = string.Format("={0}2/({1}2+{0}2)", colNamePen, colNameNonPen);
         cellB.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeRight].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;
         cellB.Borders[XlBordersIndex.xlEdgeTop].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[2, totalCols + 5];
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.NumberFormat = "#,##0";
         cellB.Formula = string.Format("={0}2 + {1}2", colNamePen, colNameNonPen);
         cellB.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeRight].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         cellB = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[1, totalCols + 5];
         cellB.Formula = "Total PAGA Periods";
         cellB.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightSkyBlue);
         cellB.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
         cellB.Borders[XlBordersIndex.xlEdgeRight].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeTop].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeRight].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeTop].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeRight].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;
         cellB.Borders[XlBordersIndex.xlEdgeTop].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;
         cellB.Borders[XlBordersIndex.xlEdgeLeft].LineStyle = Microsoft.Office.Interop.Excel.XlLineStyle.xlContinuous;
         cellB.Borders[XlBordersIndex.xlEdgeLeft].Weight = Microsoft.Office.Interop.Excel.XlBorderWeight.xlThin;
         cellB.Borders[XlBordersIndex.xlEdgeLeft].ColorIndex = Microsoft.Office.Interop.Excel.XlColorIndex.xlColorIndexAutomatic;

         Microsoft.Office.Interop.Excel.Range columnRange = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Columns[totalCols + 5];
         columnRange.ColumnWidth = 18;
         #endregion

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WritePagaViolations(Dictionary<string, List<Timesheet>> empTimesheets, List<DateTime> periodStart)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\pcappello\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                       // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
         Dictionary<DateTime, int> columnPeriods = new Dictionary<DateTime, int>();

         int row = 1;
         for (int col = 2; col < periodStart.Count; col++) {//start dates
            xlWorkSheet.Cells[row, col] = periodStart[col - 2].ToShortDateString();
            columnPeriods[periodStart[col - 2]] = 0;
            columnPeriods[periodStart[col - 2]] = col;
         }

         row = 2;
         for (int col = 2; col < periodStart.Count; col++) //end dates
            xlWorkSheet.Cells[row, col] = periodStart[col - 2].ToShortDateString();
         row++;

         // foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
         for (int pos = 0; pos < empTimesheets.Count; pos++) {
            int violation = 100;

            var employee = empTimesheets.ElementAt(pos).Value;
            employee = employee.OrderBy(x => x.periodBegin).ToList();

            foreach (Timesheet s in employee) {
               if (s.periodBegin < periodStart[0] || s.periodBegin > new DateTime(2018, 6, 25))
                  continue;

               xlWorkSheet.Cells[row, 1] = s.identifier;

               Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value]];

               if (s.HasViolation()) {
                  xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value]] = violation;
                  if (violation == 100)
                     violation = 200;
                  cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
               } else {
                  xlWorkSheet.Cells[row, columnPeriods[s.periodBegin.Value]] = 0;
                  cell.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
               }
            }
            row++;
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WritePost15EmployeeTimeSheets(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\knguyen\Desktop\schedule.xls"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                    // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            // if (row > 1000)
            //   break;
            foreach (Timesheet s in employee.Value) {
               xlWorkSheet.Cells[row, 1] = "EmpID";
               xlWorkSheet.Cells[row, 2] = "DateTime";
               xlWorkSheet.Cells[row, 3] = "Late/short lunch";
               xlWorkSheet.Cells[row, 4] = "Hrs wrkd";

               row++;

               foreach (Timecard t in s.timeCards) {
                  xlWorkSheet.Cells[row, 1] = t.identifier;
                  xlWorkSheet.Cells[row, 2] = t.shiftDate;
                  //xlWorkSheet.Cells[row, 3] = (t.mealAfter5hrs || t.mealUnder30 || t.mealsTaken == 0) ? "1" : "0";
                  xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

                  int col = 5;
                  for (int i = 0; i < t.timepunches.Count; i++) {
                     xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                     col++;

                     if (i % 2 == 1 && i + 1 < t.timepunches.Count) //
                     {
                        double timeIn = t.timepunches[i].datetime.Subtract(t.timepunches[i - 1].datetime).TotalHours;
                        if (timeIn > 5) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col - 1];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        double minutesOut = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalMinutes;
                        xlWorkSheet.Cells[row, col] = minutesOut;

                        if (minutesOut < 30) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        col++;
                     }

                  }
                  row++;
               }
               row++;

               xlWorkSheet.Cells[row - 1, 4] = "Tot act hrs";
               xlWorkSheet.Cells[row, 4] = s.actualTotalHours;// - s.actualOT.TotalHours; //Actual hours are hrs worked minus OT
               Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 4];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               //DISPLAY TIME SHEET SUMMARY
               xlWorkSheet.Cells[row - 1, 5] = "Tot list hrs";
               xlWorkSheet.Cells[row, 5] = s.listedTotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 5];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               xlWorkSheet.Cells[row - 1, 6] = "List OT";
               xlWorkSheet.Cells[row, 6] = s.listOT.TotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 6];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

               xlWorkSheet.Cells[row - 1, 7] = "Act OT";
               xlWorkSheet.Cells[row, 7] = s.actualOT.TotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 7];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
               row++;
            }
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }

      public void WriteEmployeeTimeSheets(Dictionary<string, List<Timesheet>> empTimesheets)
      {
         object misValue = System.Reflection.Missing.Value;
         string newPath = @"C:\Users\knguyen\Desktop\schedule.xls"; //System.IO.Path.Combine(activeDir, newFileName);
                                                                    // string newPath = @"C:\Users\gcr1\Desktop\schedule.xlsx"; //System.IO.Path.Combine(activeDir, newFileName);
         Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
         Microsoft.Office.Interop.Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(newPath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
         Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

         int row = 1;

         foreach (KeyValuePair<string, List<Timesheet>> employee in empTimesheets) {
            // if (row > 3000)
            //    break;
            foreach (Timesheet s in employee.Value) {
               xlWorkSheet.Cells[row, 1] = "EmpID";
               xlWorkSheet.Cells[row, 2] = "DateTime";
               xlWorkSheet.Cells[row, 3] = "Late/short lunch";
               xlWorkSheet.Cells[row, 4] = "Hrs wrkd";

               row++;

               foreach (Timecard t in s.timeCards) {
                  xlWorkSheet.Cells[row, 1] = t.identifier;
                  xlWorkSheet.Cells[row, 2] = t.shiftDate;
                  //xlWorkSheet.Cells[row, 3] = (t.mealAfter5hrs || t.mealUnder30 || t.mealsTaken == 0) ? "1" : "0";
                  xlWorkSheet.Cells[row, 4] = t.totalHrsActual.TotalHours;

                  int col = 5;
                  for (int i = 0; i < t.timepunches.Count; i++) {
                     xlWorkSheet.Cells[row, col] = t.timepunches[i].datetime;
                     col++;

                     if (i % 2 == 1 && i + 1 < t.timepunches.Count) //
                     {
                        double timeIn = t.timepunches[i].datetime.Subtract(t.timepunches[i - 1].datetime).TotalHours;
                        if (timeIn > 5) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col - 1];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        }

                        double minutesOut = t.timepunches[i + 1].datetime.Subtract(t.timepunches[i].datetime).TotalMinutes;
                        xlWorkSheet.Cells[row, col] = minutesOut;

                        if (minutesOut < 30) {
                           Microsoft.Office.Interop.Excel.Range cellColor = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, col];
                           cellColor.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        }
                        col++;
                     }

                  }
                  row++;
               }
               row++;

               xlWorkSheet.Cells[row - 1, 4] = "Tot act hrs";
               xlWorkSheet.Cells[row, 4] = s.actualTotalHours;// - s.actualOT.TotalHours; //Actual hours are hrs worked minus OT
               Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 4];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               //DISPLAY TIME SHEET SUMMARY
               xlWorkSheet.Cells[row - 1, 5] = "Tot list hrs";
               xlWorkSheet.Cells[row, 5] = s.listedTotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 5];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Blue);

               xlWorkSheet.Cells[row - 1, 6] = "List OT";
               xlWorkSheet.Cells[row, 6] = s.listOT.TotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 6];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);

               xlWorkSheet.Cells[row - 1, 7] = "Act OT";
               xlWorkSheet.Cells[row, 7] = s.actualOT.TotalHours;
               cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row, 7];
               cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
               row++;
            }
         }

         xlWorkBook.Close(true, misValue, misValue);
         xlApp.Quit();

         releaseObject(xlWorkSheet);
         releaseObject(xlWorkBook);
         releaseObject(xlApp);
      }


      //public TimeSpan CalculatePenalties(Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet, Timecard card, int row)
      //{
      //    TimeSpan timeWithoutBreak = new TimeSpan(0);
      //    TimeSpan timeWithoutMeal = new TimeSpan(0);
      //    TimeSpan totalHrs = new TimeSpan(0);

      //    int breaksTaken = 0;
      //    int mealsTaken = 0;
      //    int breaksUnder10 = 0;
      //    int mealUnder30 = 0;
      //    int mealAfter5hrs = 0;
      //    int mealAfter6hrs = 0;
      //    int missedBreak = 0;
      //    int noBreakAfter3hrs30Min = 0;
      //    int missClassified = 0;

      //    for (int i = 0; i < card.timepunches.Count - 1; i++)
      //    {
      //        var punchIn = card.timepunches[i];
      //        var punchOut = card.timepunches[i + 1];

      //        if (timeWithoutBreak.Hours > 3.5)
      //            noBreakAfter3hrs30Min++;

      //        if (timeWithoutMeal.Hours >= 5)
      //            mealAfter5hrs++;

      //        if (timeWithoutMeal.Hours >= 6)
      //            mealAfter6hrs++;

      //        if (timeWithoutBreak.Hours >= 3.5)
      //            missedBreak++;

      //        TimeSpan timeBetweenPunches = punchOut.datetime.Subtract(punchIn.datetime);

      //        //   if (punchIn.type == Constants.TYPE.INDAY)
      //        {
      //            totalHrs += timeBetweenPunches;
      //            timeWithoutBreak += timeBetweenPunches;
      //            timeWithoutMeal += timeBetweenPunches;
      //        }
      //   else if (punchIn.type == Constants.TYPE.OUTBREAK)
      //        {
      //            if (punchOut.type == Constants.TYPE.INBREAK || punchOut.type == Constants.TYPE.INLUNCH)
      //            {
      //                //TODO: Counting only up to 10 Minutes
      //                if (timeBetweenPunches.TotalMinutes > 10)
      //                    totalHrs += new TimeSpan(0, 10, 0);
      //                else
      //                    totalHrs += timeBetweenPunches; //total breaktime

      //                if (timeBetweenPunches.TotalMinutes < 10 && timeWithoutBreak.TotalHours >= 3.5) //if the break is less than 10 minutes and at least 3.5 hours has passed
      //                    breaksUnder10++;

      //                timeWithoutBreak = new TimeSpan(0);
      //                breaksTaken++;

      //                if (punchOut.type == Constants.TYPE.INLUNCH)
      //                    missClassified++;
      //            }
      //            else
      //                throw new Exception("Did not clock in from break");
      //        }
      //        else if (punchIn.type == Constants.TYPE.INBREAK)
      //        {
      //            totalHrs += timeBetweenPunches;
      //            timeWithoutBreak += timeBetweenPunches;
      //            timeWithoutMeal += timeBetweenPunches;
      //        }
      //        else if (punchIn.type == Constants.TYPE.OUTLUNCH)
      //        {
      //            if (punchOut.type == Constants.TYPE.INLUNCH || punchOut.type == Constants.TYPE.INBREAK)
      //            {
      //                if (timeBetweenPunches.TotalMinutes < 30)
      //                {
      //                    mealUnder30++;
      //                }

      //                timeWithoutMeal = new TimeSpan(0);
      //                timeWithoutBreak = new TimeSpan(0);
      //                mealsTaken++;

      //                if (punchOut.type == Constants.TYPE.INBREAK)
      //                    missClassified++;

      //            }
      //            else
      //                throw new Exception("Did not clock in for lunch");
      //        }
      //        else if (punchIn.type == Constants.TYPE.INLUNCH)
      //        {
      //            totalHrs += timeBetweenPunches;
      //            timeWithoutBreak += timeBetweenPunches;
      //            timeWithoutMeal += timeBetweenPunches;
      //        }
      //    }


      //    int offset = 6;
      //    xlWorkSheet.Cells[row, 3 + offset] = "Total hours";
      //    xlWorkSheet.Cells[row + 1, 3 + offset] = totalHrs.TotalHours;


      //    xlWorkSheet.Cells[row, 4 + offset] = "Meal after 5hrs";
      //    if (mealAfter5hrs > 0)
      //    {
      //        xlWorkSheet.Cells[row + 1, 4 + offset] = mealAfter5hrs;//Taken after 5 hours over work
      //        Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row + 1, 4 + offset];
      //        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
      //    }

      //    xlWorkSheet.Cells[row, 5 + offset] = "Meal After 6hrs";
      //    if (mealsTaken == 0 && mealAfter6hrs == 0 && totalHrs.Hours >= 6)
      //    {
      //        xlWorkSheet.Cells[row + 1, 5 + offset] = mealAfter6hrs;//Taken after 6 hours over work
      //    }

      //    xlWorkSheet.Cells[row, 6 + offset] = "No meals taken";
      //    if (mealsTaken == 0 && mealAfter5hrs == 0 && totalHrs.Hours >= 5)
      //    {
      //        xlWorkSheet.Cells[row + 1, 6 + offset] = 1; //no meals taken
      //        Microsoft.Office.Interop.Excel.Range cell = (Microsoft.Office.Interop.Excel.Range)xlWorkSheet.Cells[row + 1, 6 + offset];
      //        cell.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Pink);
      //    }

      //    xlWorkSheet.Cells[row, 7 + offset] = "Shift lngr 10 hrs w/only 1 meal"; //20
      //    if (totalHrs.TotalHours >= 10 && mealsTaken == 1)
      //    {
      //        xlWorkSheet.Cells[row + 1, 7 + offset] = 1; //no meals taken
      //    }

      //    xlWorkSheet.Cells[row, 8 + offset] = "No break (after at least 3.5 hrs)";
      //    if (missedBreak > 0) //no break besides meal, worked at least 3.5 Hours
      //        xlWorkSheet.Cells[row + 1, 8 + offset] = missedBreak;

      //    //xlWorkSheet.Cells[row, 22] = "No break or meal all shift (worked at least 3.5 hrs)";
      //    //xlWorkSheet.Cells[row, 23] = "No break/meal after 5 hrs";
      //    xlWorkSheet.Cells[row, 9 + offset] = "Break under 10 min after 3.5Hrs hours consec";
      //    if (breaksUnder10 != 0)
      //        xlWorkSheet.Cells[row + 1, 9 + offset] = breaksUnder10;

      //    xlWorkSheet.Cells[row, 10 + offset] = "Shift lngr 12 hrs";
      //    if (totalHrs.TotalHours >= 12)
      //        xlWorkSheet.Cells[row + 1, 10 + offset] = 1;

      //    xlWorkSheet.Cells[row, 11 + offset] = "Shift lngr 10 hrs";
      //    if (totalHrs.TotalHours >= 10)
      //        xlWorkSheet.Cells[row + 1, 11 + offset] = 1;

      //    xlWorkSheet.Cells[row, 12 + offset] = "Shift lngr 6 hrs";
      //    if (totalHrs.TotalHours >= 6)
      //        xlWorkSheet.Cells[row + 1, 12 + offset] = 1;

      //    xlWorkSheet.Cells[row, 13 + offset] = "Shift 5hrs+";
      //    if (totalHrs.TotalHours >= 5)
      //        xlWorkSheet.Cells[row + 1, 13 + offset] = 1;

      //    xlWorkSheet.Cells[row, 14 + offset] = "Shift 3.5hrs+";
      //    if (totalHrs.TotalHours >= 3.5)
      //        xlWorkSheet.Cells[row + 1, 14 + offset] = 1;

      //    xlWorkSheet.Cells[row, 15 + offset] = "Missclassified";
      //    xlWorkSheet.Cells[row + 1, 15 + offset] = 1;

      //    return totalHrs;
      //}

      private static void releaseObject(object obj)
      {
         try {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
            obj = null;
         } catch (Exception ex) {
            obj = null;
            throw new Exception("Unable to release the Object " + ex.ToString());
         } finally {
            GC.Collect();
         }
      }
   }
}
