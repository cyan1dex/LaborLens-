using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaborLens.SQL {
   #region Enums

   /// <summary>
   /// Simple DB directional enumeration.
   /// </summary>
   public enum ParameterDirection {
      In,
      Out
   }

   /// <summary>
   /// Enumeration of the types of parameters.  Done due to the fact that 
   /// the classes/types representing parameters are DB specific and we want 
   /// it DB agnostic.  Add to this as needed.
   /// </summary>
   public enum ParameterType {
      DBString,
      DBNvar,
      DBInteger,
      DBDecimal,
      DBBoolean,
      DBSingle,
      DBDouble,
      DBLong,
      DBDateTime,
      DBVarBinary,
      Structured
   }

   #endregion
   class StoredProcedureParameter {
      #region Member Vars

      protected ParameterDirection _direction = ParameterDirection.In;
      protected int _len = int.MinValue;
      protected string _name = String.Empty;
      protected ParameterType _type = ParameterType.DBString;
      protected Object _value = String.Empty; //may need to mess with this for clob values

      #endregion

      #region constructors

      public StoredProcedureParameter()
          : this(String.Empty, ParameterType.DBString, int.MinValue, ParameterDirection.In, String.Empty)
      {
      }

      /// <summary>
      /// Full constructor all other constructors should end up calling this
      /// </summary>
      /// <param name="name">Name of the parameter</param>
      /// <param name="type">The type of the parameter</param>
      /// <param name="len">Length of any strings</param>
      /// <param name="dir">Input or output parameter</param>
      /// <param name="val">the value, can be null</param>
      public StoredProcedureParameter(string name, ParameterType type, int len, ParameterDirection dir, Object val)
      {
         _name = name;
         _type = type;
         _len = len;
         _direction = dir;
         _value = val;
      }

      /// <summary>
      /// Non-string constructor.  Doesnt require a length.  
      /// </summary>
      /// <param name="name">Name of the parameter</param>
      /// <param name="type">The type of the parameter</param>
      /// <param name="val">The value, can be null</param>
      public StoredProcedureParameter(string name, ParameterType type, Object val)
      {
         _name = name;
         _type = type;
         _value = val;
      }

      #endregion

      #region Properties

      public String Name
      {
         get { return _name; }
         set { _name = value; }
      }

      public ParameterType DBValueType
      {
         get { return _type; }
         set { _type = value; }
      }

      public int Length
      {
         get { return _len; }
         set { _len = value; }
      }

      public ParameterDirection Direction
      {
         get { return _direction; }
         set { _direction = value; }
      }

      public Object Value
      {
         get { return _value; }
         set { _value = value; }
      }
   }
   #endregion
}
