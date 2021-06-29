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
    public class TestThamSo
    {
        public Dictionary<int, string> result = new Dictionary<int, string>();
        public List<String> actualResult = new List<string>();
        private static IEnumerable<object> GetLists()
        {
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            ExcelFilePath path = new ExcelFilePath();
            FileInfo existingFile = new FileInfo(path.showPath());
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "ThamSo");
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    String giatri = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    yield return new String[] { giatri, ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            //driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/ThamSo/Index");

        }
        [Test, TestCaseSource("GetLists")]

        public void Process(String giatri, String ketQuaMongMuon, String viTri)
        {
            if (giatri == null  || ketQuaMongMuon == null) return;
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);

            IWebElement giaTri = driver.FindElement(By.XPath("/html/body/section/section/section/div/div[1]/div/div/div/input"));
            giaTri.Clear();
            giaTri.SendKeys(giatri);
            System.Threading.Thread.Sleep(1000);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/div/div[2]/div/button"));
            btThem.Submit();
            System.Threading.Thread.Sleep(1000);
            try
            {
                IWebElement dialogTile = driver.FindElement(By.XPath("/html/body/div/div/div[2]"));
                if(dialogTile.Text == "Lỗi !")
                    ketQuaThucTe = "FAIL";
                else ketQuaThucTe = "PASS";
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                dialog.Click();
            }
            catch (Exception){}
            int s;
            if(!Int32.TryParse(giatri, out s))
                ketQuaThucTe = "FAIL";
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "ThamSo");
                int rowCount = worksheet.Dimension.End.Row;
                foreach (var obj in result)
                    worksheet.Cells[obj.Key, 3].Value = obj.Value;
                package.Save();
            }
            driver.Close();
        }
    }
}