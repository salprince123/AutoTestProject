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
        public string Ten { get; set; }
        public string DVT { get; set; }
        public string PTLN { get; set; }

        public IWebDriver driver { get; set; }
        [SetUp]
        public void SetUp()
        {
            driver = new ChromeDriver("D:\\Driver");
            driver.Manage().Window.Maximize();
            driver.Navigate().GoToUrl("https://localhost:44324/Manager/LoaiSP/Index");
        }
        [Test]
        public void Test1()
        {
            codeProcess();
            //Assert.Pass();
        }
        public void codeProcess()
        {
            IWebElement tenLoaiSP = driver.FindElement(By.Name("TenLoaiSP"));
            tenLoaiSP.SendKeys("aaaa1");
            System.Threading.Thread.Sleep(1000);
            IWebElement iWebelement = driver.FindElement(By.Id("MaDVT"));
            SelectElement selected = new SelectElement(iWebelement);
            selected.SelectByText("Bộ");
            System.Threading.Thread.Sleep(1000);
            IWebElement loiNhuan = driver.FindElement(By.Name("PhanTramLoiNhuan"));
            loiNhuan.SendKeys("1");
            System.Threading.Thread.Sleep(1000);
            IWebElement btThem = driver.FindElement(By.XPath("(//*[@value='Thêm'])"));
            btThem.Submit();
            System.Threading.Thread.Sleep(10000);
            
        }
    }
}
