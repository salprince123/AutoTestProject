using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
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
        private static object[] sourceTestCase =
        {
            new object[] {"c1", "Bộ", "1" },
            new object[] {"c2", "Bộ", "1" },
            new object[] {"c3", "Bộ", "1" },
            new object[] {"c4", "Bộ", "1" }
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

        public void codeProcess(String ten, String dvt, String loinhuan)
        {
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
                Assert.Pass();
            }
            else
            {
                IWebElement dialog = driver.FindElement(By.XPath("/html/body/div/div/div[4]/div/button"));
                System.Threading.Thread.Sleep(1000);
                dialog.Click();
                Assert.Fail();

            }
        }
    /// </summary>
    [OneTimeTearDown]
        public void Close()
        {
            driver.Close();
        }
    }
}
