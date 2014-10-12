using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    public class BasicTableBuilder: TableBuilder
    {
        //ctor
        public BasicTableBuilder()
        {
            //GetColType = new getName(getDefaultType);
            TableName = "";
        }

        //default column name
        protected override string getColName(CSVFile theData, string columnName, int columnIndex)
        {
            return (columnName == string.Empty) ? "Column" + columnIndex : columnName;
        }
        
        //check the column, and find if it is a number or a string
        protected override string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            double somethingToOut;
            for (int i = 0; i < theData.Count; i++)
            {
                if (theData[i][columnIndex] == string.Empty || !double.TryParse(theData[i][columnIndex], out somethingToOut))
                {
                    return "Text";
                }
            }
            return "Double";
        }
    }
}
