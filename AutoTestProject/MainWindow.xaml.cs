
using OfficeOpenXml;
using System;
using System.Collections.Generic;
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

namespace AutoTestProject
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            loadData();
        }
        public void loadData()
        {
            IList<object> users = new List<object>();
            /*string path = "TestCase.xlsx";
            FileInfo fileInfo = new FileInfo(path);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            ExcelPackage package = new ExcelPackage(fileInfo);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault();

            // get number of rows and columns in the sheet
            int rows = worksheet.Dimension.Rows; // 20
            // loop through the worksheet rows and columns
            for (int i = 2; i <= rows; i++)
            {
                String ten = worksheet.Cells[i, 1].Value.ToString();
                String dvt = worksheet.Cells[i, 2].Value.ToString();
                String loinhuan = worksheet.Cells[i, 3].Value.ToString();
                users.Add(new String[] { ten, dvt, loinhuan });
                //Console.WriteLine($"{ten}  {dvt}  {loinhuan}");

            }*/
            FileInfo existingFile = new FileInfo("TestCase.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                        String ten = worksheet.Cells[row, 1].Value?.ToString().Trim();
                        String dvt = worksheet.Cells[row, 2].Value?.ToString().Trim();
                        String loinhuan = worksheet.Cells[row, 3].Value?.ToString().Trim();
                        users.Add(new String[] {  ten, dvt, loinhuan});
                }
            }
            /*foreach (String[] obj in users)
            {
                //Console.WriteLine($"{obj[0]} ");
                //Console.WriteLine($"{obj[0]}  {obj[1]}  {obj[2]}");
                string startupPath = System.IO.Directory.GetCurrentDirectory();
                Console.WriteLine($" \n \n {startupPath}");
            }*/
            Console.WriteLine($" \n \n {System.IO.Path.Combine(@Directory.GetCurrentDirectory(), "\\TestCase.xlsx")}");
            //Console.WriteLine($" \n {@"D:\", "TestCase.xlsx"}");
        }
    }
}
