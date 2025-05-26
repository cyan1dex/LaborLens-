using System;
using System.Collections.Generic;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Configuration;


namespace LaborLens.SQL {
   class StoredProcedure {

      #region Private Variables

      private readonly List<StoredProcedureParameter> _params;
      private int _commandTimeOut = 240;
      private string _dbschema = string.Empty;
      //  protected ErrorLogRepository _errorLogDAL;
      private string _returnName = "RET";
      private string _storedProcedureName = String.Empty;

      #endregion

      #region Constructor

      /// <summary>
      /// Default constructor.  Does nothing
      /// </summary>
      public StoredProcedure()
          : this(String.Empty, string.Empty, new List<StoredProcedureParameter>(), String.Empty)
      {
         //_errorLogDAL = new ErrorLogRepository();
      }

      /// <summary>
      /// Full constructor.  All shortened should call
      /// </summary>
      /// <param name="name">Name of the stored procedure</param>
      /// <param name="pars">List of the parameters</param>
      /// <param name="ret">name of the return variable in oracle</param>
      public StoredProcedure(string name, string schemaName, List<StoredProcedureParameter> pars, string ret)
      {
         _storedProcedureName = name;
         _dbschema = schemaName;
         _params = pars;
         _returnName = ret;
      }

      /// <summary>
      /// Simple call.  use this in most cases.  Uses default return variable name
      /// </summary>
      /// <param name="name">Name of the stored procedure</param>
      /// <param name="pars">List of the parameters</param>
      public StoredProcedure(string name, string schemaName, List<StoredProcedureParameter> pars)
          : this(name, schemaName, pars, String.Empty)
      {
      }

      #endregion

      #region Public Properties

      public int CommandTimeOut
      {
         set { _commandTimeOut = value; }
      }

      /// <summary>
      /// The name of the stored procedure
      /// </summary>
      public string StoredProcedureName
      {
         set { _storedProcedureName = value; }
         get { return _storedProcedureName; }
      }

      /// <summary>
      /// The name of the stored procedure
      /// </summary>
      public string SchemaName
      {
         set { _dbschema = value; }
         get { return _dbschema; }
      }

      /// <summary>
      /// The parameters of the stored procedure
      /// Note:  MAV - There are ways that you can make this non-accessable but I don't recall this
      /// at the moment
      /// </summary>
      public List<StoredProcedureParameter> Parameters
      {
         get { return _params; }
      }

      /// <summary>
      /// The name of the return value for oracle
      /// </summary>
      public String ReturnName
      {
         get { return _returnName; }
         set { _returnName = value; }
      }

      #endregion

      #region public methods

      /// <summary>
      /// Create the dataset.  This should be only visible to the DAL layer.  For MS SQL DBs
      /// Note:   THe core functionality for this procedure has been cloned to executeNonQuery() to handle routines that 
      ///         do not expect any type of return value.
      /// </summary>
      /// <returns></returns>
      private DataTable MSSQL_executeDataSet()
      {
         var ret = new DataTable();

         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            //JBienvenu 19202 2012-12-28 switch: Implemented ParameterTypeAdapter system.
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createParam(parm.Name, new StructuredAdapter(), parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createParam(parm.Name, new DateTimeAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createParam(parm.Name, new NvarAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBVarBinary:
                  parms.Add(Param.createParam(parm.Name, new VarBinaryAdapter(), parm.Value, ParameterDirection.In));
                  break;
               default:    //handles any case that is not specificed above with default of appaneding to the string
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }
            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               //JBienvenu 19202 2012-12-28 foreach: Changed to use parms (which were not used at all before). Implemented ParameterTypeAdapter system.
               parms.ForEach((parm) => parm.ToSqlParameter(sqlCommand));
               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               using (SqlDataReader sqlDataReader = sqlCommand.ExecuteReader()) {
                  if (sqlDataReader.VisibleFieldCount > 0) {
                     foreach (DataRow dRow in sqlDataReader.GetSchemaTable().Rows) {
                        if (!(dRow["ColumnName"].ToString().Trim().Equals(""))) {
                           ret.Columns.Add(dRow["ColumnName"].ToString());
                           ret.Columns[dRow["ColumnName"].ToString()].ReadOnly = false;
                           ret.Columns[dRow["ColumnName"].ToString()].AllowDBNull = true;
                           ret.Columns[dRow["ColumnName"].ToString()].DataType =
                                   Type.GetType(dRow["DataType"].ToString());
                        }
                     }
                     ret.Load(sqlDataReader);
                  }
               }
            }
         }
         return ret;
      }

      public DataTable GetDataTableFromDataReader(IDataReader dataReader)
      {
         DataTable schemaTable = dataReader.GetSchemaTable();
         DataTable resultTable = new DataTable();
         foreach (DataRow dataRow in schemaTable.Rows) {
            DataColumn dataColumn = new DataColumn();
            dataColumn.ColumnName = dataRow["ColumnName"].ToString();
            dataColumn.DataType = Type.GetType(dataRow["DataType"].ToString());
            dataColumn.ReadOnly = (bool)dataRow["IsReadOnly"];
            dataColumn.AutoIncrement = (bool)dataRow["IsAutoIncrement"];
            dataColumn.Unique = (bool)dataRow["IsUnique"];

            resultTable.Columns.Add(dataColumn);
         };

         while (dataReader.Read()) {
            DataRow dataRow = resultTable.NewRow();
            for (int i = 0; i < resultTable.Columns.Count; i++) {
               dataRow[i] = dataReader[i];
            }
            resultTable.Rows.Add(dataRow);
         }
         return resultTable;
      }


      /// <summary>
      /// This routine is for the execution of code that will not return any values.
      /// Note:  That this is a copy past from MSSSQL_ExecuteQuery with just the returns changed.  
      /// </summary>
      public void ExecuteNonQuery()
      {
         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            //JBienvenu 19202 2012-12-28 switch: Implemented ParameterTypeAdapter system.
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createParam(parm.Name, new StructuredAdapter(), parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createParam(parm.Name, new DateTimeAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createParam(parm.Name, new NvarAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBVarBinary:
                  parms.Add(Param.createParam(parm.Name, new VarBinaryAdapter(), parm.Value, ParameterDirection.In));
                  break;
               //handles any case that is not specified above with default of appending to the string
               default:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }

            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               //JBienvenu 19202 2012-12-28 foreach: Changed to use parms (which were not used at all before). Implemented ParameterTypeAdapter system.
               parms.ForEach((parm) => parm.ToSqlParameter(sqlCommand));
               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               sqlCommand.ExecuteNonQuery();
            }
         }
      }

      public void ExecuteDataSetWithIDOutputParam(out int ID)
      {
         int result;
         string strResult;
         ExecuteDataSetWithIDOutputParam(out strResult);
         if (int.TryParse(strResult, out result)) {
            ID = result;
         } else {
            ID = -1;
         }
      }

      public void ExecuteDataSetWithIDOutputParam(out string ID)
      {
         string sID = null;

         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            //JBienvenu 19202 2012-12-28 switch: Implemented ParameterTypeAdapter system.
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createParam(parm.Name, new StructuredAdapter(), parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createParam(parm.Name, new DateTimeAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createParam(parm.Name, new NvarAdapter(), parm.Value, ParameterDirection.In));
                  break;
               //handles any case that is not specificed above with default of appaneding to the string
               default:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }
            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               //JBienvenu 19202 2012-12-28 foreach: Changed to use parms (which were not used at all before). Implemented ParameterTypeAdapter system.
               parms.ForEach((parm) => parm.ToSqlParameter(sqlCommand));

               SqlParameter myparam3 = sqlCommand.Parameters.Add("@ID", SqlDbType.VarChar, 20);
               myparam3.Direction = System.Data.ParameterDirection.Output;

               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               SqlDataReader reader = sqlCommand.ExecuteReader();
               //ret = GetDataTableFromDataReader(reader);
               reader.Close();
               sID = (string)sqlCommand.Parameters["@ID"].Value;
               reader.Close();
            }
         }

         ID = sID;
      }

      /// <summary>
      /// Create the dataset.  This should be only visible to the DAL layer.  For MS SQL DBs
      /// </summary>
      /// <returns></returns>
      public void ExecuteDataSetWithOutputParam(out DataTable dtRet, out String sID)
      {

         DataTable ret = null;
         String ID = string.Empty;

         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            //JBienvenu 19202 2012-12-28 switch: Implemented ParameterTypeAdapter system.
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createParam(parm.Name, new StructuredAdapter(), parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createParam(parm.Name, new DateTimeAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createParam(parm.Name, new NvarAdapter(), parm.Value, ParameterDirection.In));
                  break;
               //handles any case that is not specificed above with default of appaneding to the string
               default:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }
            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               //JBienvenu 19202 2012-12-28 foreach: Changed to use parms (which were not used at all before). Implemented ParameterTypeAdapter system.
               parms.ForEach((parm) => parm.ToSqlParameter(sqlCommand));

               SqlParameter myparam3 = sqlCommand.Parameters.Add("@referenceTableType", SqlDbType.VarChar, 10);
               myparam3.Direction = System.Data.ParameterDirection.Output;

               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               SqlDataReader reader = sqlCommand.ExecuteReader();
               ret = GetDataTableFromDataReader(reader);
               reader.Close();
               ID = (string)sqlCommand.Parameters["@referenceTableType"].Value;
               reader.Close();
            }
         }
         dtRet = ret; sID = ID;
      }

      /// <summary>
      /// Create the dataset.  This should be only visible to the DAL layer.  For MS SQL DBs
      /// </summary>
      /// <returns></returns>
      public DataSet ExecuteMultipleDataSet()
      {
         var ret = new DataSet();

         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            //JBienvenu 19202 2012-12-28 switch: Implemented ParameterTypeAdapter system.
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createParam(parm.Name, new StructuredAdapter(), parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  if (parm.Direction == ParameterDirection.In) {
                     parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  } else {
                     parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.Out));
                  }
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createParam(parm.Name, new DateTimeAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createParam(parm.Name, new NvarAdapter(), parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBInteger:
                  parms.Add(Param.createParam(parm.Name, new IntegerAdapter(), parm.Value, ParameterDirection.In));
                  break;
               //handles any case that is not specificed above with default of appaneding to the string
               default:
                  parms.Add(Param.createParam(parm.Name, new StringAdapter(), parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }
            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               //JBienvenu 19202 2012-12-28 foreach: Changed to use parms (which were not used at all before). Implemented ParameterTypeAdapter system.
               parms.ForEach((parm) => parm.ToSqlParameter(sqlCommand));
               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               var mySqlDataAdapter = new SqlDataAdapter();
               mySqlDataAdapter.SelectCommand = sqlCommand;
               mySqlDataAdapter.Fill(ret);
            }
         }
         return ret;
      }

      public string ReadConnectionString()
      {
         // ConnectionStringSettingsCollection connections = ConfigurationManager.ConnectionStrings;
         // ConnectionStringSettings constr = connections["ApplicationServices"];


         // return @"Data Source=DESKTOP-C5I1SPH;User ID=kevin;Password=time;Initial Catalog=master";
         // return @"Data Source=localhost\SQLEXPRESS01;User ID=pcapello;Password=agppci22;Initial Catalog=master";

         //TODO: update to use app.config
         // return @"Data Source=35.239.230.162;User ID=codici;Password=C0D1C1TODAY;Initial Catalog=master";
         return @"Data Source=CODICI;User ID=codici;Password=agppci22;Initial Catalog=master;TrustServerCertificate=true;";
      }


      public DataTable ExecuteDataSet()
      {
         DataTable ret = null;
         ret = MSSQL_executeDataSet();
         return ret;
      }

      public DataSet ExecuteResourceDataSet()
      {
         var ret = new DataSet();

         var parms = new List<Param>();
         foreach (StoredProcedureParameter parm in _params) {
            switch (parm.DBValueType) {
               case ParameterType.Structured:
                  parms.Add(Param.createresourceParam(parm.Name, SqlDbType.Structured, parm.Value,
                                                                          ParameterDirection.In));
                  break;
               case ParameterType.DBString:
                  if (parm.Direction == ParameterDirection.In) {
                     parms.Add(Param.createresourceParam(parm.Name, DbType.String, parm.Value, ParameterDirection.In));
                  } else {
                     parms.Add(Param.createresourceParam(parm.Name, DbType.String, parm.Value, ParameterDirection.Out));
                  }
                  break;
               case ParameterType.DBDateTime:
                  parms.Add(Param.createresourceParam(parm.Name, DbType.DateTime, parm.Value, ParameterDirection.In));
                  break;
               case ParameterType.DBNvar:
                  parms.Add(Param.createresourceParam(parm.Name, DbType.String, parm.Value, ParameterDirection.In));
                  break;
               //handles any case that is not specificed above with default of appaneding to the string
               default:
                  parms.Add(Param.createresourceParam(parm.Name, DbType.String, parm.Value, ParameterDirection.In));
                  break;
            }
         }
         using (var sqlConn = new SqlConnection(ReadConnectionString())) {
            try {
               sqlConn.Open();
            } catch {
               throw;
            }
            using (var sqlCommand = new SqlCommand(_storedProcedureName, sqlConn)) {
               foreach (StoredProcedureParameter parm in _params) {
                  sqlCommand.Parameters.AddWithValue(parm.Name, parm.Value);
               }
               sqlCommand.CommandTimeout = _commandTimeOut;
               sqlCommand.CommandType = CommandType.StoredProcedure;
               var mySqlDataAdapter = new SqlDataAdapter();
               mySqlDataAdapter.SelectCommand = sqlCommand;
               mySqlDataAdapter.Fill(ret);
            }
         }
         return ret;
      }

      #endregion
   }
}
