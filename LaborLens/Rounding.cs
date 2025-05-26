using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
    class Rounding {

      public void CalculateRounding(Dictionary<string, List<Timesheet>> empSheets)
      {
         Decimal roundingCompare = 0;
         int roundingForEmp = 0;
         int roundingForComp = 0;
         int tlWeeks = 0;
         int t = 0;

         foreach (KeyValuePair<string, List<Timesheet>> entry in empSheets) {
            foreach (Timesheet s in entry.Value) {
               if (s.timeCards.Count() <= 14) {
                  s.ProcessRounding();

                  roundingForComp += s.roundedShiftsForCompany;
                  roundingForEmp += s.roundedShiftsForEmployee;
                  roundingCompare += s.roudingBalance;
                  tlWeeks += s.roundedWOrksWeeks;
               }
            }
         }

      }
   }
}
