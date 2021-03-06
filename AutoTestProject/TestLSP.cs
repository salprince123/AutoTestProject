using NUnit.Framework;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoTestProject
{
    public class TestLSP
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
                ExcelWorksheet worksheet= package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "LoaiSP");
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    String ten = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String dvt = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String loinhuan = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    yield return new String[] { ten, dvt, loinhuan , ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/LoaiSP/Index");
            
        }
        [Test, TestCaseSource("GetLists")]

        public void codeProcess(String ten, String dvt, String loinhuan, String ketQuaMongMuon, String viTri)
        {
            if (ten == null || dvt == null || loinhuan == null) return;
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);
            IWebElement tenLoaiSP = driver.FindElement(By.Name("TenLoaiSP"));
            tenLoaiSP.Clear();
            tenLoaiSP.SendKeys(ten);
            System.Threading.Thread.Sleep(500);

            IWebElement iWebelement = driver.FindElement(By.Id("MaDVT"));
            SelectElement selected = new SelectElement(iWebelement);
            selected.SelectByText(dvt);
            System.Threading.Thread.Sleep(500);

            IWebElement loiNhuan = driver.FindElement(By.Name("PhanTramLoiNhuan"));
            loiNhuan.Clear();
            loiNhuan.SendKeys(loinhuan);
            tenLoaiSP.SendKeys("");
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/main/div/div[1]/div/div[2]/form/input[2]"));
            btThem.Submit();
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "LoaiSP");
                int rowCount = worksheet.Dimension.End.Row;
                foreach (var obj in result)
                    worksheet.Cells[obj.Key, 5].Value = obj.Value;
                package.Save();
            }
            driver.Close();
        }
    }
}
