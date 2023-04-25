using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Sprint1._71
{
    public class Tests
    {
        IWebDriver driver;
        

        [SetUp]
        public void Setup()
        {
            driver = new ChromeDriver();
            driver.Navigate().GoToUrl("https://testpostframesolver.azurewebsites.net/");
            driver.Manage().Window.Maximize();
            IWebElement userName = driver.FindElement(By.Name("Username"));
            IWebElement password = driver.FindElement(By.Name("Password"));
            userName.SendKeys("hshah@keymark.com");
            password.SendKeys("Sss1234!");
            IWebElement remenberMe = driver.FindElement(By.Id("RememberMe"));
            remenberMe.Click();
            IWebElement loginBtn = driver.FindElement(By.XPath("//*[@id='loginForm']/form/div[4]/div/input"));
            loginBtn.Click();
            Wait(5000);
        }

        [Test]
        [Obsolete]
        public void Pfs2841()
        {
            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
            IWebElement startfromScratch = driver.FindElement(By.XPath("/html/body/div[2]/ul/li[1]/a"));
            startfromScratch.Click();
            Wait(5000);
            WaitForElement(driver, ElementSelectorType.XPath, "/html/body/div/div[1]/div/div[5]/div[4]/div/div[5]/div[4]");
            Wait(5000);
            WebDriverWait wait = WebDriverWait(driver, 2, 2);
            WaitForElement(driver, ElementSelectorType.XPath, "//*[@id='middleToolbar']/table");
            Wait(8000);
           
            //Select Job Review Tab
            //IWebElement table = driver.FindElement(By.XPath("//*[@id='middleToolbar']/table"));
            //var row = table.FindElements(By.TagName("tr"));



            IWebElement jobReview = driver.FindElement(By.XPath("//*[@id='tb_middleToolBar_item_showMaterials']/table/tbody/tr/td/table/tbody/tr/td[2]"));
            jobReview.Click();
            Wait(5000);
            ////Select Trim tab 
            //IWebElement tablemenu = driver.FindElement(By.XPath("//*[@id='tabsup']/table"));
            //var tableRow = tablemenu.FindElements(By.TagName("tr"));
            IWebElement tabTrim = driver.FindElement(By.Id("tabs_tabsMaterial_tab_1"));
            tabTrim.Click();
            Wait(5000);

            IWebElement table = driver.FindElement(By.Id("grid_MaterialsGrid_body"));
            IReadOnlyCollection<IWebElement> rows = table.FindElements(By.XPath("//*[@id='grid_MaterialsGrid_records']/table/tbody/tr"));
           

            foreach (IWebElement row in rows)
            {
                string file = @"C:\New folder\SmartBuild.txt";

                var id = row.GetAttribute("id");
                
                if (id == ""  || id == "grid_MaterialsGrid_rec_top")
                    continue;

                if (id == "grid_MaterialsGrid_rec_bottom")
                    break;

                IWebElement table_Td1 = driver.FindElement(By.Id($"{id}"));
                IReadOnlyCollection<IWebElement> rows_Td1 = table_Td1.FindElements(By.XPath($"//*[@id='{id}']/td[1]/div"));
                var id1 = row.GetAttribute("rows_Td1");
                string stext = rows_Td1.FirstOrDefault().Text;

                File.AppendAllText(file, stext);

                IReadOnlyCollection<IWebElement> rows_Td2 = table_Td1.FindElements(By.XPath($"//*[@id='{id}']/td[7]/div"));
                var id2 = row.GetAttribute("rows_Td2");
                string s2text = rows_Td2.FirstOrDefault().Text;
                File.AppendAllText(file, s2text);

            }
        }

        

        #region Private Method(s)
        private static bool WaitForElement(IWebDriver driver, ElementSelectorType type, string itemPath)
        {
            Wait(3000);
            WebDriverWait wait = WebDriverWait(driver, 1, 4);
            wait.IgnoreExceptionTypes(typeof(NotFoundException), typeof(NoSuchElementException));
            switch (type)
            {
                case ElementSelectorType.ClassName:
                    wait.Until(e =>
                    {
                        var elementList = e.FindElements(By.ClassName(itemPath));
                        return elementList.Count == 1;
                    });
                    break;
                case ElementSelectorType.XPath:
                    wait.Until(e =>
                    {
                        var elementList = e.FindElements(By.XPath(itemPath));
                        return elementList.Count == 1;
                    });
                    break;
                case ElementSelectorType.Id:
                    break;
                default:
                    break;
            }


            return true;
        }

        private static WebDriverWait WebDriverWait(IWebDriver driver, int timeout, int pollingInterval)
        {
            return new(driver, TimeSpan.FromMinutes(timeout))
            {
                PollingInterval = TimeSpan.FromSeconds(pollingInterval),
            };
        }

        private static void Wait(int time) => Thread.Sleep(time);

        private enum ElementSelectorType
        {
            ClassName = 0,
            XPath,
            Id
        }

        #endregion
        //[TearDown]
        public void CloseEvent()
        {
            driver.Close();
        }
    }
}