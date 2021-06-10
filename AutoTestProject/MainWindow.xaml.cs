
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
            List<String> list = new List<string>();
            for (int i = 0; i < 10; i++)
                list.Add(i.ToString());
            writeData(list);
        }
        public void loadData()
        {
            IList<object> users = new List<object>();
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
            
        }
        public void writeData(List<String> list)
        {

            FileInfo existingFile = new FileInfo("TestCase.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;
                worksheet.Cells[1, 4].Value = "Result";
                for (int i=2; i < list.Count+2; i++)
                {
                    worksheet.Cells[i, 4].Value = i-2;
                }
                try
                {
                    package.Save();
                }
                catch (Exception) { }
            }
            

        }
    }
}
