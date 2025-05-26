using System.Data;
using Microsoft.Data.SqlClient;

namespace LaborLens.SQL {
   public abstract class ParameterTypeAdapter {
      public abstract SqlParameter SetType(ref SqlParameter p);

      public ParameterTypeAdapter() { }
   }

   public class StringAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.String;
         p.SqlDbType = SqlDbType.VarChar;
         return p;
      }
   }

   public class NvarAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.String;
         p.SqlDbType = SqlDbType.NVarChar;
         return p;
      }
   }

   public class IntegerAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Int32;
         p.SqlDbType = SqlDbType.Int;
         return p;
      }
   }

   public class DecimalAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Decimal;
         p.SqlDbType = SqlDbType.Decimal;
         return p;
      }
   }

   public class BooleanAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Boolean;
         p.SqlDbType = SqlDbType.Bit;
         return p;
      }
   }

   public class SingleAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Single;
         p.SqlDbType = SqlDbType.Float;
         return p;
      }
   }

   public class DoubleAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Double;
         p.SqlDbType = SqlDbType.Float;
         return p;
      }
   }

   public class LongAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Int64;
         p.SqlDbType = SqlDbType.BigInt;
         return p;
      }
   }

   public class DateTimeAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.DateTime;
         p.SqlDbType = SqlDbType.DateTime;
         return p;
      }
   }

   public class StructuredAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Object;
         p.SqlDbType = SqlDbType.Structured;
         return p;
      }
   }

   public class VarBinaryAdapter : ParameterTypeAdapter {
      public override SqlParameter SetType(ref SqlParameter p)
      {
         p.DbType = DbType.Byte;
         p.SqlDbType = SqlDbType.VarBinary;
         return p;
      }
   }
}
