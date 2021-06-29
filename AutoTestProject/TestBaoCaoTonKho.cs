using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;

namespace AutoTestProject
{
    public class TestBaoCaoTonKho
    {
        public Dictionary<int, string> result = new Dictionary<int, string>();
        public List<String> result1 = new List<string>();
        private static IEnumerable<object> GetLists()
        {
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            ExcelFilePath path = new ExcelFilePath();
            FileInfo existingFile = new FileInfo(path.showPath());
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "BCTonKho");
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    String thang = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String nam= worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    yield return new String[] { thang,nam, ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/BCTonKho/Index");

        }
        [Test, TestCaseSource("GetLists")]

        public void Process(String thang,String nam, String ketQuaMongMuon, String viTri)
        {
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);
            IWebElement Thang = driver.FindElement(By.XPath("/html/body/section/section/section/div/form[1]/div/fieldset/div[1]/div/div[1]/div/input"));
            Thang.Clear();
            Thang.SendKeys(thang);
            System.Threading.Thread.Sleep(500);

            IWebElement Nam = driver.FindElement(By.XPath("/html/body/section/section/section/div/form[1]/div/fieldset/div[1]/div/div[2]/div/input"));
            Nam.Clear();
            Nam.SendKeys(nam);
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/div/form[2]/div/div[3]/button"));
            btThem.Submit();
            System.Threading.Thread.Sleep(1000);
            try
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                dialog.Click();
                ketQuaThucTe = "FAIL";
            }
            catch (Exception)
            {
                ketQuaThucTe = "PASS";
            }
            if (ketQuaThucTe.Equals(ketQuaMongMuon)) result[vitri] = "PASS";
            else result[vitri] = "FAIL";
            Assert.AreEqual(ketQuaMongMuon.Trim(), ketQuaThucTe.Trim());

        }
        /// </summary>
        [OneTimeTearDown]
        public void Close()
        {
            //save result in excel file
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            ExcelFilePath path = new ExcelFilePath();
            FileInfo existingFile = new FileInfo(path.showPath());
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "BCTonKho");
                int rowCount = worksheet.Dimension.End.Row;
                worksheet.Cells[1, 2].Value = "Result";
                foreach (var obj in result)
                    worksheet.Cells[obj.Key, 4].Value = obj.Value;
                package.Save();
            }
            driver.Close();
        }
    }
}
