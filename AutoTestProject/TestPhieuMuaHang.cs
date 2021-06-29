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
    public class TestPhieuMuaHang
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuMuaHang");
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    String nhaCC = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String ngayLap = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String tenSP= worksheet.Cells[row, 3].Value?.ToString().Trim();
                    String soLuong = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    yield return new String[] { nhaCC,ngayLap,tenSP,soLuong, ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            //driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/CT_PhieuMuaHang/Create");

        }
        [Test, TestCaseSource("GetLists")]

        public void Process(String nhacc, String ngaylap,String ten, String soluong, String ketQuaMongMuon, String viTri)
        {
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);
            IWebElement iWebelement = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[1]/div/div[1]/div[1]/div/select"));
            SelectElement selected = new SelectElement(iWebelement);
            selected.SelectByText(nhacc);
            System.Threading.Thread.Sleep(500);

            IWebElement calendar = driver.FindElement(By.Id("txtNgayGiaoDich"));
            calendar.SendKeys(ngaylap);
            System.Threading.Thread.Sleep(500);

            IWebElement CbtenSP = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[1]/div[1]/div/select"));
            SelectElement tenSP = new SelectElement(CbtenSP);
            tenSP.SelectByText(ten);
            System.Threading.Thread.Sleep(500);

            IWebElement soLuong = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[2]/div[1]/div/input"));
            soLuong.Clear();
            soLuong.SendKeys(soluong);
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[2]/div[4]/div/input"));
            btThem.Click();

            System.Threading.Thread.Sleep(500);
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
            System.Threading.Thread.Sleep(5000);
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuMuaHang");
                int rowCount = worksheet.Dimension.End.Row;
                foreach (var obj in result)
                    worksheet.Cells[obj.Key, 6].Value = obj.Value;
                package.Save();
            }
            driver.Close();
        }
    }
}
