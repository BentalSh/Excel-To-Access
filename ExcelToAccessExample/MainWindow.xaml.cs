using ExcelToDB;
using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ExcelToAccessExample
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            List<Type> tableBuilders = new List<Type>();
            tableBuilders.Add(typeof(BasicTableBuilder));
            tableBuilders.Add(typeof(SimpleHeaderTable));
            tableBuilders.Add(typeof(HeaderAndAdditionalDataTable));
            chooseTableBuilder.ItemsSource = tableBuilders;
            chooseTableBuilder.SelectedIndex = 0;
            
        }

        private void makeDB_Click(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(fileName.Text))
            {
                MessageBox.Show("File doesn't exist");
                return;
            }
            if (!File.Exists(accessFile.Text))
            {
                MessageBox.Show("Access file doesn't exist");
                return;
            }
            TableBuilder tb;
            if (chooseTableBuilder.SelectedItem as Type == typeof(HeaderAndAdditionalDataTable))
            {
                int theFirstDataRow = 0;
                if (!int.TryParse(firstDataRow.Text, out theFirstDataRow))
                {
                    MessageBox.Show("First data row must be set to an integer number");
                    return;
                }
                tb = new HeaderAndAdditionalDataTable(theFirstDataRow);
            }
            else if (chooseTableBuilder.SelectedItem as Type == typeof(SimpleHeaderTable))
            {
                tb = new SimpleHeaderTable();
            }
            else
            {
                tb = new BasicTableBuilder();
            }
            CSVFile theData = new CSVFile();
            string theFileName = fileName.Text;
            OleDbConnection conn = new OleDbConnection(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + accessFile.Text + ";Persist Security Info=True");
            conn.Open();
            if (theFileName.Substring(theFileName.LastIndexOf(".")+1).ToLower()!="csv")
            { 
                Type officeType = Type.GetTypeFromProgID("Excel.Application");
                if (officeType == null)
                {
                }
                else
                {
                    Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
                    app.DisplayAlerts = false;
                    Microsoft.Office.Interop.Excel.Workbook excelWorkbook = app.Workbooks.Open(fileName.Text);
                    string theActualFileName = theFileName.Remove(theFileName.LastIndexOf(".")).Substring(theFileName.LastIndexOf("\\") + 1);
                    theFileName = System.IO.Directory.GetCurrentDirectory() + "\\" + theActualFileName + ".csv";
                    excelWorkbook.SaveAs(theFileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV);
                    excelWorkbook.Close();
                    app.Quit();
                }
            }
            theData.init(theFileName);
            tb.makeOrUpdateTable(theData, conn);
            conn.Close();
        }

        private void chooseTableBuilder_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            firstDataRow.IsEnabled = (chooseTableBuilder.SelectedItem as Type == typeof(HeaderAndAdditionalDataTable));
        }

        private void chooseFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            // Set filter for file extension and default file extension 
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx|Excel Files (2007) (*.xls)|*.xls|Comma separated value files (*.csv)|*.csv";
            Nullable<bool> result = dlg.ShowDialog();
            if (result.HasValue && result.Value)
            {
                fileName.Text = dlg.FileName;
            }
        }

        private void chooseAccessFileButton_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            // Set filter for file extension and default file extension 
            dlg.Filter = "Access file (*.accdb)|*.accdb";
            Nullable<bool> result = dlg.ShowDialog();
            if (result.HasValue && result.Value)
            {
                accessFile.Text = dlg.FileName;
            }
        }
    }
}
