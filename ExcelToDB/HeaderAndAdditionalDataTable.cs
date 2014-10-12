using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    public class HeaderAndAdditionalDataTable: TableBuilder
    {
        public HeaderAndAdditionalDataTable(int firstDataRow)
        {
            this.FirstDataRow = firstDataRow;
        }
        public int FirstDataRow { get; set; }
        //default column name
        protected override string getColName(CSVFile theData, string columnName, int columnIndex)
        {
            return (columnName == string.Empty) ? "Column" + columnIndex : columnName;
        }
        
        protected override string getTableName(CSVFile theData)
        {
            return (theData.Count>0) ? theData[0][0] : base.getTableName(theData);

        }
        //check the column, and find if it is a number or a string
        protected override string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            double somethingToOut;
            int numberOfFieldsInHeader = theData[FirstDataRow].Count;
            if (columnIndex >= numberOfFieldsInHeader)
            {
                return "Text";
            }
            else
            {
                for (int i = FirstDataRow; i < theData.Count; i++)
                {
                    if (theData[i][columnIndex] == string.Empty || !double.TryParse(theData[i][columnIndex], out somethingToOut))
                    {
                        return "Text";
                    }
                }
                return "Double";
            }
            
        }
        public override List<string> getColumns(CSVFile theData)
        {
            List<string> headerColumns = new List<string>();
            headerColumns.AddRange(theData[FirstDataRow]);
            for (int i = 1; i < FirstDataRow; i++)
            {
                if (theData[i][0] != string.Empty) headerColumns.Add(theData[i][0]);
            }
            return headerColumns;
        }
        protected override void fillTable(CSVFile theData, DbConnection db)
        {
            List<string> constantValues = new List<string>();
            List<string> columns = getColumns(theData);
            int numOfDataRowsInHeader = theData[FirstDataRow].Count;
            for (int i = 1; i < FirstDataRow; i++)
            {
                if (theData[i][0] != string.Empty)
                    constantValues.Add(theData[i][1]);
            }
            for (int i = FirstDataRow+1; i < theData.Count; i++)
            {

                DbCommand insertCmd = db.CreateCommand();
                string valuesString = "";
                string columnsString = "";
                string insertSql = "insert into " + TableName;
                int j;
                for (j=0; j < theData[i].Count; j++)
                {
                    DbParameter par = insertCmd.CreateParameter();
                    par.ParameterName = "@Value" + j.ToString();
                    par.Value = theData[i][j];
                    valuesString += ((j > 0) ? "," : "") + "@Value" + j.ToString();
                    columnsString += ((j > 0) ? "," : "") + columns[j];
                    insertCmd.Parameters.Add(par);
                }
                for (; j < columns.Count; j++)
                {
                    if (theData[i][0] != string.Empty)
                    {
                        DbParameter par = insertCmd.CreateParameter();
                        par.ParameterName = "@Value" + j.ToString();
                        par.Value = constantValues[j - numOfDataRowsInHeader];
                        valuesString += ((j > 0) ? "," : "") + "@Value" + j.ToString();
                        columnsString += ((j > 0) ? "," : "") + columns[j];
                        insertCmd.Parameters.Add(par);
                    }
                }

                insertSql += " (" + columnsString + ") values (" + valuesString + ")";
                insertCmd.CommandText = insertSql;
                insertCmd.ExecuteNonQuery();
            }
        }
    }
}
