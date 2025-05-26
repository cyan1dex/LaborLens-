using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
     class AutoDeductions {

      public  int AutoDeductHours(Dictionary<string, List<Timesheet>> timeSheets)
      {
         int autodeduct = 0;

         int totFavorEmp = 0;
         double emloyeeHrsSaved = 0;

         int totFavorCmpny = 0;
         double companyHrsSaved = 0;

         foreach (KeyValuePair<string, List<Timesheet>> entry in timeSheets) {

            foreach (Timesheet s in entry.Value) {
               int ad = 0;
               double act = s.actualTotalHours;// s.actualHours.TotalHours;// - s.actualOT.TotalHours; //Actual hours are hrs worked minus OT
               var tot = s.listedTotalHours;
               var check = s.stub.doubleOtHrs + s.stub.otHrs + s.stub.regHrs; //Check hours can be different from the timecard hours
               var otList = s.stub.otHrs;
               var otACt = s.actualOT.TotalHours;
               var otCheck = s.stub.doubleOtHrs;

               double periodAD = 0;

               foreach (Timecard c in s.timeCards) {
                  ad += c.possibleAutoDeduct == true ? 1 : 0;

                  if (c.possibleAutoDeduct && Math.Abs(c.totalHrsActual.TotalHours - (c.regHrsListed + c.otListed + .5)) < .25)
                     autodeduct++;
               }



               //year auto-deduct implemented
               //if (ad > 0 && s.periodBegin > new DateTime(2020, 3, 15)) {
               //   autodeduct += ad;
               //   periodAD += ad;
               //}

               if (periodAD > 0) {
                  act = act - periodAD / 2.0;
               }


               if (Math.Abs(act - check) > .04 && Math.Abs(act - check) < 3) {
                  if (check > act) {
                     totFavorEmp++;
                     emloyeeHrsSaved += check - act;
                  } else {
                     totFavorCmpny++;
                     companyHrsSaved += act - check;
                  }
               }
            }

         }

         return autodeduct;
      }
   }
}
