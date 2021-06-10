﻿using NUnit.Framework;
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
        /*public class LSP
        {
            public string Ten { get; set; }
            public string DVT { get; set; }
            public string PTLN { get; set; }
            public LSP(String ten, String dvt, String ptln)
            {
                Ten = ten; DVT = dvt; PTLN = ptln;
            }
        }
        private static  LSP[] sourceTestCase =
        {
            new LSP("c1", "Bộ", "1"), 
           new LSP("c1", "Bộ", "1"),
           new LSP("c2", "Bộ", "1"),
            new LSP("c3", "Bộ", "1"),
           new LSP("c3", "Bộ", "1")
        };

        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/LoaiSP/Index");            
        }
        [Test]
        [TestCaseSource("sourceTestCase")]
        [TestCase("a1","Bộ","1")]
        [TestCase("a1", "Bộ", "2")]
        [TestCase("a2", "Bộ", "3")]
        [TestCase("a5", "Bộ", "4")]

        public void codeProcess(LSP lsp)
        {
            IWebElement tenLoaiSP = driver.FindElement(By.Name("TenLoaiSP"));
            //tenLoaiSP.SendKeys("aaaa1");
            tenLoaiSP.Clear();
            tenLoaiSP.SendKeys(lsp.Ten);
            System.Threading.Thread.Sleep(500);

            IWebElement iWebelement = driver.FindElement(By.Id("MaDVT"));
            SelectElement selected = new SelectElement(iWebelement);
            //selected.SelectByText("Bộ");
            selected.SelectByText(lsp.DVT);
            System.Threading.Thread.Sleep(500);

            IWebElement loiNhuan = driver.FindElement(By.Name("PhanTramLoiNhuan"));
            //loiNhuan.SendKeys("1");
            loiNhuan.Clear();
            loiNhuan.SendKeys(lsp.PTLN);
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/main/div/div[1]/div/div[2]/form/input[2]"));
            btThem.Submit();
            
            IWebElement vali = driver.FindElement(By.XPath("/html/body/section/section/section/main/div/div[1]/div/div[2]/form/div[1]/div[2]/span"));
            if(vali.Text.Length ==0)
            {
                Assert.Pass();
            }    
            else
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                System.Threading.Thread.Sleep(1000);
                dialog.Click();
                Assert.Fail();
                
            }
        }*/


        /// VERSION 2
        /* private static object[] sourceTestCase =
         {
             new object[] {"c1", "Bộ", "1" },
             new object[] {"c2", "Bộ", "1" },
             new object[] {"c3", "Bộ", "1" },
             new object[] {"c4", "Bộ", "1" }
         };*/
        public List<String> result = new List<string>();
        private static IEnumerable<object> GetLists()
        {
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            FileInfo existingFile = new FileInfo("D:\\TestCase.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet= package.Workbook.Worksheets.FirstOrDefault(x => x.Name == "Sheet1");
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    String ten = worksheet.Cells[row, 1].Value?.ToString().Trim();
                    String dvt = worksheet.Cells[row, 2].Value?.ToString().Trim();
                    String loinhuan = worksheet.Cells[row, 3].Value?.ToString().Trim();
                    yield return new String[] { ten, dvt, loinhuan };
                }
            }
        }
        public IWebDriver driver { get; set; }
        [OneTimeSetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/LoaiSP/Index");
            
        }
        [Test, TestCaseSource("GetLists")]

        public void codeProcess(String ten, String dvt, String loinhuan)
        {
            if (ten == null || dvt == null || loinhuan == null) return;
            IWebElement tenLoaiSP = driver.FindElement(By.Name("TenLoaiSP"));
            //tenLoaiSP.SendKeys("aaaa1");
            tenLoaiSP.Clear();
            tenLoaiSP.SendKeys(ten);
            System.Threading.Thread.Sleep(500);

            IWebElement iWebelement = driver.FindElement(By.Id("MaDVT"));
            SelectElement selected = new SelectElement(iWebelement);
            //selected.SelectByText("Bộ");
            selected.SelectByText(dvt);
            System.Threading.Thread.Sleep(500);

            IWebElement loiNhuan = driver.FindElement(By.Name("PhanTramLoiNhuan"));
            //loiNhuan.SendKeys("1");
            loiNhuan.Clear();
            loiNhuan.SendKeys(loinhuan);
            System.Threading.Thread.Sleep(500);

            IWebElement btThem = driver.FindElement(By.XPath("/html/body/section/section/section/main/div/div[1]/div/div[2]/form/input[2]"));
            btThem.Submit();

            IWebElement vali = driver.FindElement(By.XPath("/html/body/section/section/section/main/div/div[1]/div/div[2]/form/div[1]/div[2]/span"));
            if (vali.Text.Length == 0)
            {
                result.Add("Pass");
                Assert.Pass();
            }
            else
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                System.Threading.Thread.Sleep(1000);
                dialog.Click();
                result.Add("Fail");
                Assert.Fail();
            }
        }
    /// </summary>
    [OneTimeTearDown]
        public void Close()
        {
            //save result in excel file
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            FileInfo existingFile = new FileInfo("D:\\TestCase.xlsx");
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                //get the first worksheet in the workbook
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.End.Row;
                worksheet.Cells[1, 4].Value = "Result";
                for (int i = 2; i < result.Count + 2; i++)
                {
                    worksheet.Cells[i, 4].Value = result[i - 2];
                }
                try
                {
                    package.Save();
                }
                catch (Exception) { }
            }
            driver.Close();
        }
    }
}
