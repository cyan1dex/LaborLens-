using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens {
   public class Timepunch {

      public DateTime datetime;
      public double hrsListed;
      public bool isRestBreak;
      public double regHrsListed;
      public double otHrsListed;
      public double dblOtListed;
   }

   public class TimeIn : Timepunch {

   }

   public class Timeout : Timepunch {

   }
}
