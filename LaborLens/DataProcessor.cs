using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
  public class DataProcessor {

      public List<Shift> GetAnalyzedShifts(Dictionary<string, List<Timecard>> empSheets)
      {
         List<Shift> shifts = new List<Shift>();

         foreach (KeyValuePair<string, List<Timecard>> entry in empSheets) {
            foreach (Timecard c in entry.Value) {

               if (!c.invalid) {
                  Shift s = new Shift();
                  s.AnalyzeShift(c);
                  shifts.Add(s);
               }
            }
         }
         return shifts;
      }

      public class EmployeeDateRanges {
         public DateTime? ShiftStart { get; private set; }
         public DateTime? ShiftEnd { get; private set; }
         public DateTime? PayStart { get; private set; }
         public DateTime? PayEnd { get; private set; }

         public EmployeeDateRanges(DateTime? shiftStart, DateTime? shiftEnd,
             DateTime? payStart, DateTime? payEnd)
         {
            ShiftStart = shiftStart;
            ShiftEnd = shiftEnd;
            PayStart = payStart;
            PayEnd = payEnd;
         }
      }

      public Dictionary<string, EmployeeDateRanges> ProcessDateRanges(Dictionary<string, List<Timecard>> shifts, Dictionary<string, List<PayStub>> paystubs)
      {
         var results = new Dictionary<string, EmployeeDateRanges>();
         var employeeIds = new HashSet<string>(shifts.Keys.Union(paystubs.Keys));

         foreach (var empId in employeeIds) {
            List<Timecard> employeeShifts;
            List<PayStub> employeePaystubs;
            shifts.TryGetValue(empId, out employeeShifts);
            paystubs.TryGetValue(empId, out employeePaystubs);

            DateTime? shiftStart = null;
            DateTime? shiftEnd = null;
            DateTime? payStart = null;
            DateTime? payEnd = null;

            if (employeeShifts != null && employeeShifts.Any()) {
               shiftStart = employeeShifts.Min(s => s.shiftDate);
               shiftEnd = employeeShifts.Max(s => s.shiftDate);
            }

            if (employeePaystubs != null && employeePaystubs.Any()) {
               payStart = employeePaystubs.Min(p => p.periodBegin);
               payEnd = employeePaystubs.Max(p => p.periodEnd);
            }

            results[empId] = new EmployeeDateRanges(shiftStart, shiftEnd, payStart, payEnd);
         }

         return results;
      }
   }
}
