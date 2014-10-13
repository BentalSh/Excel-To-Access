using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    public class CSVFile
    {
        readonly List<List<string>> values = new List<List<string>>();
        public string FileName { get; private set; }
        public void init(string fileName)
        {
            values.Clear();
            StreamReader reader = new StreamReader(fileName,Encoding.Default);
            while (!reader.EndOfStream)
            {
                string theLine = reader.ReadLine();
                List<string> line = new List<string>();
                string[] possibleValues = theLine.Split(new char[] {','});
                bool foundQuotes = false;
                string tempValue = "";
                for (int i = 0; i < possibleValues.Length; i++)
                {
                    string theValue = possibleValues[i];
                    if (!foundQuotes && theValue.Length>0 && theValue.Trim()[0] == '"')
                    {
                        tempValue = theValue.Substring(theValue.IndexOf('"') + 1);
                        if (theValue.Length>1 && theValue.Trim()[theValue.Trim().Length - 1] == '"')
                        {
                            line.Add(tempValue.Substring(0,tempValue.LastIndexOf('"')));
                        }
                        else foundQuotes = true;
                    }
                    else if (foundQuotes)
                    {
                        if (theValue.Length>0 && theValue.Trim()[theValue.Trim().Length - 1] == '"')
                        {
                            foundQuotes = false;
                            line.Add(tempValue + theValue.Substring(0,theValue.LastIndexOf('"')));
                            tempValue = "";
                        }
                        else
                        {
                            tempValue += theValue;
                        }
                        continue;
                    }
                    else line.Add(theValue);
                }
                values.Add(line);
            }
            this.FileName = fileName;
            reader.Close();
        }

        public List<string> this[int index]
        {
            get
            {
                return values[index];
            }
            set
            {
                values[index].Clear();
                values[index].AddRange(value);
            }
        }
        public int Count
        {
            get
            {
                return values.Count;
            }
        }
    }
}
