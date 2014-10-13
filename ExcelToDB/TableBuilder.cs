using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    public class TableBuilder
    {
        public TableBuilder()
        {
            KeyColumn = -1;
        }
        //if not -1, use as key
        public int KeyColumn { get; set; }
        public string TableName { get; set; }
        protected virtual string getColName(CSVFile theData, string columnName,int columnIndex)
        {
            return columnName.Replace(".","");
        }
        protected virtual string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            return "char(50)";
        }
        //default will be file name
        

        public virtual List<string> getColumns(CSVFile theData)
        {
            return getDefaultColumns(theData);
        }

        protected virtual string getTableName(CSVFile theData)
        {
            return theData.FileName.Remove(theData.FileName.LastIndexOf(".")).Substring(theData.FileName.LastIndexOf("\\")+1).Replace(" ","_");
        }

        //just make a column for each column in the csv, with no name
        protected List<string> getDefaultColumns(CSVFile theData)
        {
            List<string> ans = new List<string>();
            //assuming data is a simple table
            if (theData.Count > 0)
                for (int i = 0; i < theData[0].Count; i++)
                {
                    ans.Add("Column" + i.ToString());
                }
            return ans;
        }
        //create the table in the DB
        public void makeOrUpdateTable(CSVFile theData, DbConnection db)
        {
            if (TableName==null || TableName == string.Empty)
                TableName = getTableName(theData);
            DataTable structureTable = db.GetSchema("Tables");
            bool doesExist = false;
            for (int i = 0; i < structureTable.Rows.Count; i++)
            {
                DataRow dr = structureTable.Rows[i];
                if (dr["TABLE_NAME"].ToString() == TableName)
                {
                    doesExist = true;
                    break;
                }
            }
            if (!doesExist)
            {
                string createTableString = "CREATE TABLE " + TableName;
                if (KeyColumn<0) createTableString += " (ID AUTOINCREMENT)";
                else
                {
                    string keyColumnName= getColumns(theData)[KeyColumn];
                    createTableString += " (" + keyColumnName + " " + getColType(theData, keyColumnName, KeyColumn) + ")";
                }

                DbCommand createCmd = db.CreateCommand();
                createCmd.CommandText = createTableString;
                createCmd.ExecuteNonQuery();

                makeColumns(theData, db);
            }

            fillTable(theData, db);


        }

        protected void makeColumns(CSVFile theData, DbConnection db)
        {
            List<string> columns = getColumns(theData);
            if (columns!=null)
            {
                for (int i = 0; i < columns.Count; i++)
                {
                    if (i != KeyColumn) addColumn(theData, db, columns[i], "char(50)", i);
                    else makeColumnKey(theData, db, columns[i], "char(50)", i);
                }
            }
        }
        protected virtual void makeColumnKey(CSVFile theData, DbConnection db, string columnName, string columnType, int columnIndex)
        {
            DbCommand createCmd = db.CreateCommand();
            createCmd.CommandText = "ALTER TABLE " + TableName + " ADD PRIMARY KEY (" + columnName + ")";
            createCmd.ExecuteNonQuery();

        }
        protected virtual void addColumn(CSVFile theData, DbConnection db, string columnName, string columnType, int columnIndex)
        {
            DataTable columnTable = db.GetSchema("Columns");
            bool doesColumnExists = false;
            for (int i = 0; i < columnTable.Rows.Count; i++)
            {
                DataRow dr = columnTable.Rows[i];
                if (dr["TABLE_NAME"].ToString() == TableName)
                {
                    if (dr["COLUMN_NAME"].ToString() == columnName)
                    {
                        doesColumnExists = true;
                        break;
                    }
                }
                else continue;
            }
            string nameToUse = getColName(theData, columnName, columnIndex);
            string typeToUse = getColType(theData, columnName, columnIndex);

            if (!doesColumnExists)
            {
                string addColumnString = "ALTER TABLE " + TableName + " ADD `" + nameToUse +  "` " + typeToUse;
                DbCommand addCmd = db.CreateCommand();
                DbParameter par = addCmd.CreateParameter();
                addCmd.CommandText = addColumnString;
                addCmd.ExecuteNonQuery();
            }
        }
        protected virtual void fillTable(CSVFile theData,DbConnection db)
        {           
            fillTable(theData, db, 0);
        }
        protected void fillTable(CSVFile theData, DbConnection db, int startIndex)
        {
            fillTable(theData, db, startIndex, false);
        }
        //when last row is summary, you want to ignore it
        protected void fillTable(CSVFile theData, DbConnection db, int startIndex, bool ignoreLastRow)
        {
            List<string> columns = getColumns(theData);
            for (int i = startIndex; i < (ignoreLastRow ? theData.Count -1: theData.Count); i++)
            {
                fillRow(db, i, columns, theData[i]);
            }
        }
          
        protected void fillRow(DbConnection db, int rowToWorkOn, List<string> columns, List<string> values)
        {
            bool shouldUpdate = false;
            if (KeyColumn >= 0) //then first check if exists
            {
                DbCommand checkIfRowExists = db.CreateCommand();
                checkIfRowExists.CommandText = "SELECT COUNT(*) FROM " + TableName + " WHERE " + columns[KeyColumn] + "=@Value";
                DbParameter idParam = checkIfRowExists.CreateParameter();
                idParam.ParameterName = "@Value";
                idParam.Value = values[KeyColumn];
                checkIfRowExists.Parameters.Add(idParam);
                shouldUpdate = ((int)checkIfRowExists.ExecuteScalar()) > 0;
            }
            DbCommand updateOrInsertCmd = db.CreateCommand();
            string valuesString = "";
            string columnsString = "";
            string cmdSql;
            if (!shouldUpdate) cmdSql = "insert into " + TableName;
            else
            {
                cmdSql = "update " + TableName;
                columnsString = " set ";
            }

            for (int j = 0; j < Math.Min(values.Count, columns.Count); j++)
            {
                if (!shouldUpdate)
                {
                    valuesString += ((j > 0) ? "," : "") + "@Value" + j.ToString();
                    columnsString += ((j > 0) ? "," : "") + "[" + columns[j] + "]";
                }
                else
                {
                    if (j == KeyColumn)
                        continue;
                    columnsString += ((j > 0 && (KeyColumn==0 && j>1)) ? "," : "") + "[" + columns[j] + "]=@Value" + j.ToString();

                }
                DbParameter par = updateOrInsertCmd.CreateParameter();
                par.ParameterName = "@Value" + j.ToString();
                if (values[j] == "") par.Value = DBNull.Value;
                else par.Value = values[j];
                updateOrInsertCmd.Parameters.Add(par);
            }
            if (!shouldUpdate)
                cmdSql += " (" + columnsString + ") values (" + valuesString + ")";
            else
            {
                cmdSql += columnsString + " where [" + columns[KeyColumn] + "]=@KeyValue";
                DbParameter keyParam = updateOrInsertCmd.CreateParameter();
                keyParam.ParameterName = "@KeyValue";
                keyParam.Value = values[KeyColumn];
                updateOrInsertCmd.Parameters.Add(keyParam);
            }
            updateOrInsertCmd.CommandText = cmdSql;
            updateOrInsertCmd.ExecuteNonQuery();
        }
    }
}
