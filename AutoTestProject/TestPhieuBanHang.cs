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
    public class TestPhieuBanHang
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuBanHang");
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    String tenKH = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String ngayLap = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String tenSP = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    String soLuong = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    yield return new String[] { tenKH, ngayLap, tenSP, soLuong, ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            //driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/CT_PhieuBanHang/Create");

        }
        [Test, TestCaseSource("GetLists")]

        public void Process(String tenKH, String ngaylap, String tensp, String soluong, String ketQuaMongMuon, String viTri)
        {
            if (ngaylap == null) return;
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);
            IWebElement tenKh = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[1]/div/div/div[1]/div/input"));
            tenKh.Clear();
            if (tenKH != null)
                tenKh.SendKeys(tenKH);
            System.Threading.Thread.Sleep(500);

            IWebElement calendar = driver.FindElement(By.Id("txtNgayGiaoDich"));
            calendar.SendKeys(ngaylap);
            System.Threading.Thread.Sleep(500);

            IWebElement CbtenSP = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[1]/div[1]/div/select"));
            SelectElement tenSP = new SelectElement(CbtenSP);
            tenSP.SelectByText(tensp);
            System.Threading.Thread.Sleep(500);
            
            IWebElement soLuong = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[2]/div[1]/div/input"));
            soLuong.Clear();
            if (soluong != null)
                soLuong.SendKeys(soluong);
            tenKh.SendKeys("");
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[2]/div[4]/div/input"));
            btThem.Click();

            System.Threading.Thread.Sleep(500);
            try
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                System.Threading.Thread.Sleep(1000);
                dialog.Click();
                ketQuaThucTe = "FAIL";
                IWebElement btXoa = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[3]/table/tbody/tr/td[9]/input"));
                btXoa.Click();
            }
            catch (Exception)
            {
                ketQuaThucTe = "PASS";
            }
            
            if (ketQuaThucTe=="PASS")
            {
                IWebElement btXacNhan = driver.FindElement(By.XPath("/html/body/section/section/section/form/div/div/input"));
                btXacNhan.Click();
                System.Threading.Thread.Sleep(500);
                try
                {
                    IWebElement dialogTile = driver.FindElement(By.XPath("/html/body/div/div/div[2]"));
                    if (dialogTile.Text == "Lỗi !")
                        ketQuaThucTe = "FAIL";
                    else ketQuaThucTe = "PASS";
                    IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                    System.Threading.Thread.Sleep(1000);
                    dialog.Click();
                }
                catch (Exception)
                {
                }
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuBanHang");
                int rowCount = worksheet.Dimension.End.Row;
                foreach (var obj in result)
                    worksheet.Cells[obj.Key, 6].Value = obj.Value;
                package.Save();
            }
            
            driver.Close();

        }
    }
}
