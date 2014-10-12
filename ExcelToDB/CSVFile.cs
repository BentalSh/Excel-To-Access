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
            StreamReader reader = new StreamReader(fileName);
            while (!reader.EndOfStream)
            {
                string theLine = reader.ReadLine();
                List<string> line = new List<string>();
                line.AddRange(theLine.Split(new char[] {','}));
                values.Add(line);
            }
            this.FileName = fileName;
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
