using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {

   class Payment {
      public Decimal rate;
      public Decimal hrs;
      public Decimal pay;
   }

   public class PayStub : IEquatable<PayStub> {
      public string identifier;
      public string name;
      public string ssn;
      public DateTime? periodBegin;
      public DateTime? periodEnd;
      public DateTime? checkDate;
      public bool periodsMissing;

      public bool invalid;
      public double totalHrs;
      public double regPayWrong;

      public double regRate;
      public double regHrs;
      public double regPay;

      public double otRate;
      public double otHrs;
      public double otPay;

      public double doubleOtHrs;
      public double doubleOtPay;
      public double doubltOtRate;

      public double penaltyRate;
      public double penaltyHrs;
      public double penaltyPay;

      public double swingRate;
      public double swingHrs;
      public double swingPay;

      public double swingDblRate;
      public double swingDblHrs;
      public double swingDblPay;

      public double swingOtRate;
      public double swingOtHrs;
      public double swingOtPay;

      public bool onDutyLunch;

      public double bonus;
      public double unpaidBonusOT;
      public double commissions;
      public double performance;

      public static DateTime earliest = DateTime.MaxValue;
      public static DateTime latest = DateTime.MinValue;
      public List<string> validationIssues = new List<string>();

      public void AnalyzeUnpaidBonus()
      {
         if (bonus > 0 && regHrs > 0 && otHrs > 0) {
            unpaidBonusOT = ((bonus / regHrs) * 1.5) * otHrs;
         }
      }

      public void Merge(PayStub two)
      {
         this.penaltyHrs += two.penaltyHrs;
         this.penaltyPay += two.penaltyPay;

         this.regHrs += two.regHrs;
         this.regPay += two.regPay;
         this.regRate = two.regRate;

         this.otHrs += two.otHrs;
         this.otPay += two.otPay;
         this.otRate = two.otRate;

         this.doubleOtHrs += two.doubleOtHrs;
         this.doubleOtPay += two.doubleOtPay;
         this.doubltOtRate = two.doubltOtRate;

         this.bonus += two.bonus;

         // doubltOtRate = doubleOtHrs != 0 ? doubleOtPay / doubleOtHrs : 0;
         // otRate = otHrs != 0 ? otPay / otHrs : 0;
         //regRate = regPay / regHrs;
      }

      public bool Equals(PayStub obj)
      {
         if (obj == null) return false;

         PayStub objAsPart = obj as PayStub;

         if (this.identifier == obj.identifier && this.periodBegin.Value == obj.periodBegin.Value && this.periodEnd.Value == this.periodEnd.Value)
            return true;
         return false;
      }
   }
}
