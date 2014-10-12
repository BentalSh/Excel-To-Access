using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    public class SimpleHeaderTable: BasicTableBuilder
    {
        //default column name
        protected override string getColName(CSVFile theData, string columnName, int columnIndex)
        {
            return (columnName == string.Empty) ? "Column" + columnIndex : columnName;
        }

        //check the column, and find if it is a number or a string
        protected override string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            double somethingToOut;
            for (int i = 1; i < theData.Count; i++)
            {
                if (theData[i][columnIndex] == string.Empty || !double.TryParse(theData[i][columnIndex], out somethingToOut))
                {
                    return "Text";
                }
            }
            return "Double";
        }
        public override List<string> getColumns(CSVFile theData)
        {
            return theData[0];
        }
        protected override void fillTable(CSVFile theData, DbConnection db)
        {
            fillTable(theData, db, 1);
        }
    }
    
}
