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

        protected virtual string getColName(CSVFile theData, string columnName,int columnIndex)
        {
            return columnName;
        }

        protected virtual string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            return "Text";
        }
        //default will be file name
        public string TableName { get; set; }

        public virtual List<string> getColumns(CSVFile theData)
        {
            return getDefaultColumns(theData);
        }

        protected virtual string getTableName(CSVFile theData)
        {
            return theData.FileName.Remove(theData.FileName.LastIndexOf(".")).Substring(theData.FileName.LastIndexOf("\\")+1);
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
                string createTableString = "CREATE TABLE " + TableName + "(ID AUTOINCREMENT)";
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
                    addColumn(theData, db, columns[i], "", i);
                }
            }
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
                string addColumnString = "ALTER TABLE " + TableName + " ADD " + nameToUse + " " + typeToUse;
                DbCommand addCmd = db.CreateCommand();
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
            List<string> columns = getColumns(theData);
            for (int i = startIndex; i < theData.Count; i++)
            {
                DbCommand insertCmd = db.CreateCommand();
                string valuesString = "";
                string columnsString = "";
                string insertSql = "insert into " + TableName;
                for (int j = 0; j < theData[i].Count; j++)
                {
                    DbParameter par =  insertCmd.CreateParameter();
                    par.ParameterName = "@Value" + j.ToString();
                    par.Value = theData[i][j];
                    valuesString += ((j > 0) ? "," : "") + "@Value" + j.ToString();
                    columnsString += ((j > 0) ? "," : "") + columns[j];
                    insertCmd.Parameters.Add(par);
                }
                insertSql += " (" + columnsString + ") values (" + valuesString + ")";
                insertCmd.CommandText = insertSql;
                insertCmd.ExecuteNonQuery();
                
            }
        }
    }
}
