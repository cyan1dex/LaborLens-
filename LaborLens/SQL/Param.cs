using System.Data;
using System.Data.Sql;
using Microsoft.Data.SqlClient;

namespace LaborLens.SQL {
   struct Param {
      public ParameterDirection direction;
      public string name;
      public int size;
      public object value;
      public ParameterTypeAdapter typeAdapter;
      public SqlDbType Sqltype;
      public DbType type;

      public static Param createParam(string nm, int len, ParameterTypeAdapter tp, object val, ParameterDirection dir)
      {
         var ret = new Param();
         ret.name = nm;
         ret.size = len;
         ret.typeAdapter = tp;
         ret.value = val;
         ret.direction = dir;

         return ret;
      }

      public static Param createParam(string nm, ParameterTypeAdapter tp, object val, ParameterDirection dir)
      {
         var ret = new Param();
         ret.name = nm;
         ret.typeAdapter = tp;
         ret.value = val;
         ret.direction = dir;

         return ret;
      }

      public static Param createresourceParam(string nm, DbType tp, object val, ParameterDirection dir)
      {
         var ret = new Param();
         ret.name = nm;
         ret.type = tp;
         ret.value = val;
         ret.direction = dir;

         return ret;
      }
      public static Param createresourceParam(string nm, SqlDbType tp, object val, ParameterDirection dir)
      {
         var ret = new Param();
         ret.name = nm;
         ret.Sqltype = tp;
         ret.value = val;
         ret.direction = dir;

         return ret;
      }

      public object[] toArray()
      {
         var ret = new object[5];

         ret[0] = name;
         ret[1] = size;
         ret[2] = typeAdapter;
         ret[3] = value;
         ret[4] = direction;

         return ret;
      }

      //JBienvenu 19202 2012-12-31 new
      public SqlParameter ToSqlParameter(SqlCommand cmd)
      {
         {
            SqlParameter p = cmd.Parameters.AddWithValue(this.name, this.value);
            p = this.typeAdapter.SetType(ref p);
            return p;
         }
      }
   }
}
