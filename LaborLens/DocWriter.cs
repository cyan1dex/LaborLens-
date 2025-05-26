using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;

namespace LaborLens {
   class DocWriter {

      public void WriteDocument(Analysis analysis)
      {
         Object oMissing = System.Reflection.Missing.Value;
         //OBJECTS OF FALSE AND TRUE
         Object oTrue = true;
         Object oFalse = false;
         //CREATING OBJECTS OF WORD AND DOCUMENT

         Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
         Microsoft.Office.Interop.Word.Document oWordDoc = new Microsoft.Office.Interop.Word.Document();
         //MAKING THE APPLICATION VISIBLE
         oWord.Visible = false;
         //ADDING A NEW DOCUMENT TO THE APPLICATION
         oWordDoc = oWord.Documents.Add("C:\\Users\\CYAN1\\OneDrive\\Desktop\\Law Cases\\Analysis - Case.docx");

         oWordDoc.Variables["mealViolations"].Value = analysis.mealViolations.ToString("###,###");
         oWordDoc.Variables["totalShifts"].Value = analysis.totMealViolShifts.ToString("##,###,###");// //analysis.totalShifts.ToString();
         oWordDoc.Variables["startDate"].Value = Timecard.earliest.ToShortDateString();
         oWordDoc.Variables["endDate"].Value = Timecard.latest.ToShortDateString();
         oWordDoc.Variables["violRate"].Value = (100 * ((analysis.mealViolations / (double)analysis.totMealViolShifts))).ToString("#0.00");

         oWordDoc.Variables["synMeals"].Value = analysis.meal30.ToString("##,###,###");
         oWordDoc.Variables["mealsTaken"].Value = analysis.mealsTaken.ToString("##,###,###");
         oWordDoc.Variables["synShifts"].Value = analysis.shift8.ToString("##,###,###");

         oWordDoc.Variables["totEmpsTime"].Value = analysis.totalEmployeesTimedata.ToString();
         oWordDoc.Variables["totEmpsPay"].Value = analysis.totalEmployeesPaydata.ToString();

         oWordDoc.Variables["mealPremHrs"].Value = analysis.hrsPaidMealViolations.ToString("###,###");
         oWordDoc.Variables["mealPremPay"].Value = analysis.paidMealViolationsAmt.ToString("##,###,###.##");
         oWordDoc.Variables["mealPremFirstDt"].Value = analysis.minMealViolPayDate.ToShortDateString();

         oWordDoc.Variables["pagaViols"].Value = "?";
         oWordDoc.Variables["pagaShifts"].Value = "?";
         oWordDoc.Variables["pageViolRate"].Value = "?";

         oWordDoc.Variables["avgShiftsWeek"].Value = (analysis.totalShifts / (double)analysis.totalWorkweeks).ToString("######.0");
         oWordDoc.Variables["avgShiftLngth"].Value = analysis.avgShiftlength.ToString("######.0");
         oWordDoc.Variables["totalWeeks"].Value = analysis.totalWorkweeks.ToString("###,###");
         oWordDoc.Variables["avgPay"].Value = analysis.regRate.ToString("##.##");

         oWordDoc.Fields.Update();
         oWordDoc.SaveAs2("C:\\Users\\CYAN1\\OneDrive\\Desktop\\Law Cases\\updated.docx");
         oWordDoc.Close();
         oWord.Quit();

      }
   }
}
