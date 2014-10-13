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
            return (columnName == string.Empty) ? "Column" + columnIndex : columnName.Replace(" ","_");
        }
        public bool IgnoreLastRow { get; set; }

        //check the column, and find if it is a number or a string
        protected override string getColType(CSVFile theData, string columnName, int columnIndex)
        {
            double somethingToOut;
            for (int i = 1; i < theData.Count; i++)
            {
                if (theData[i][columnIndex] == string.Empty || !double.TryParse(theData[i][columnIndex], out somethingToOut))
                {
                    return "char(50)";
                }
            }
            return "Double";
        }
        public override List<string> getColumns(CSVFile theData)
        {
            List<string> ans = new List<string>();
            for (int i = 0; i < theData[0].Count; i++)
            {
                ans.Add(getColName(theData, theData[0][i], i));
            }
            return ans;
        }
        protected override void fillTable(CSVFile theData, DbConnection db)
        {
            fillTable(theData, db, 1, IgnoreLastRow);
        }
    }
    
}
