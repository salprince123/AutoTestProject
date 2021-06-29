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
    public class TestPhieuDichVu
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuDichVu");
                int rowCount = worksheet.Dimension.End.Row;
                for (int row = 2; row <= rowCount; row++)
                {
                    String tenKH = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String sdt = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String ngayLap = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    String tenloaiDV = worksheet.Cells[row, 4].Value?.ToString().Trim();
                    String donGia = worksheet.Cells[row, 5].Value?.ToString().Trim();
                    String soLuong = worksheet.Cells[row, 6].Value?.ToString().Trim();
                    String ngayGiao = worksheet.Cells[row, 7].Value?.ToString().Trim();
                    String tinhTrang = worksheet.Cells[row,8].Value?.ToString().Trim();
                    String traTruoc = worksheet.Cells[row, 9].Value?.ToString().Trim();
                    String ketQuaMongMuon = worksheet.Cells[row, 10].Value?.ToString().Trim();
                    yield return new String[] { tenKH, sdt,ngayLap, tenloaiDV,donGia, soLuong,ngayGiao,tinhTrang,traTruoc, ketQuaMongMuon, row.ToString() };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            //driver.Manage().Window.Minimize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/CT_PhieuDichVu/Create");

        }
        [Test, TestCaseSource("GetLists")]

        public void Process(String tenKH, String sdt, String ngayLap, String tenloaiDV, String donGia, String soLuong, 
                                                String ngayGiao, String tinhTrang, String traTruoc, String ketQuaMongMuon, String viTri)
        {
            String ketQuaThucTe = "FAIL";
            int vitri = int.Parse(viTri);
           
            IWebElement tenKh = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[1]/div/div[1]/div[1]/div/input"));
            tenKh.Clear();
            if (tenKH != null)
                tenKh.SendKeys(tenKH);
            System.Threading.Thread.Sleep(500);

            IWebElement SDT = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[1]/div/div[1]/div[2]/div/input"));
            SDT.Clear();
            SDT.SendKeys(sdt);
            System.Threading.Thread.Sleep(500);

            IWebElement NgayLap = driver.FindElement(By.Id("txtNgayLap"));
            NgayLap.SendKeys(ngayLap);
            System.Threading.Thread.Sleep(500);

            IWebElement cbTenLoaiDV = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[1]/div[1]/div/select"));
            SelectElement TenLoaiDV = new SelectElement(cbTenLoaiDV);
            TenLoaiDV.SelectByText(tenloaiDV);
            System.Threading.Thread.Sleep(500);


            IWebElement DonGia = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[1]/div[2]/div/input"));
            DonGia.Clear();
            DonGia.SendKeys(donGia);
            System.Threading.Thread.Sleep(500);

            IWebElement SoLuong = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[2]/div[1]/div/input"));
            SoLuong.Clear();
            SoLuong.SendKeys(soLuong);
            System.Threading.Thread.Sleep(500);

            IWebElement NgayGiao = driver.FindElement(By.Id("txtNgayGiao"));
            NgayGiao.SendKeys(ngayGiao);
            System.Threading.Thread.Sleep(500);

            IWebElement cbTinhTrang = driver.FindElement(By.XPath("/ html / body / section / section / section / form / fieldset[2] / div / div[2] / div[3] / div / select"));
            
            //System.Threading.Thread.Sleep(500);

            IWebElement TraTruoc = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[3]/div[1]/div/input"));
            TraTruoc.Clear();
            TraTruoc.SendKeys(traTruoc);
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[2]/div/div[3]/div[5]/div/input"));
            btThem.Click();            
            System.Threading.Thread.Sleep(500);
            try
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                System.Threading.Thread.Sleep(500);
                dialog.Click();
                ketQuaThucTe = "FAIL";
                IWebElement btXoa = driver.FindElement(By.XPath("/html/body/section/section/section/form/fieldset[3]/table/tbody/tr/td[12]/input"));
                btXoa.Click();
            }
            catch (Exception)
            {
                ketQuaThucTe = "PASS";
            }
            //System.Threading.Thread.Sleep(5000);
            
            if (ketQuaThucTe == "PASS")
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
                ExcelWorksheet worksheet = package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "PhieuDichVu");
                int rowCount = worksheet.Dimension.End.Row;
                foreach (var obj in result)
                {
                    worksheet.Cells[obj.Key, 11].Value = obj.Value;
                    if(obj.Value=="FAIL")
                        worksheet.Cells[obj.Key, 11].Style.Font.Bold=true;
                    else worksheet.Cells[obj.Key, 11].Style.Font.Bold = false;
                }
                    
                package.Save();
            }

            driver.Close();

        }
    }
}
