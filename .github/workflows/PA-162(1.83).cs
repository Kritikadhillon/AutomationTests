using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Events;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartbuildAutomation
{
    class VerifyCloning : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        [Test, Order(1)]
        public void Clone()
        {
            test = extent.CreateTest("Verify Cloning for the Color");

            SetUpWizard();
            DeleteOldEntries();

            // Click on First row of Color table
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_body']/div[1]/table/tbody/tr[3]"))).Click();
            test.Log(Status.Info, "Click on First row of Color table");

            // Click on Clone Icon 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_toolbar']/table/tbody/tr/td[9]"))).Click();
            test.Log(Status.Info, "Click on Clone Icon");

            // Enter Color Name 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[2]/div/input[1]"))).SendKeys("Red Roof Color");
            test.Log(Status.Info, "Enter Color Name ");

            // Enter Color Code
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[4]/div/input[1]"))).SendKeys("Test Clone Color");
            test.Log(Status.Info, "Enter Color Code");

            // Select HEX code 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[3]/div/input[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div[1]/div/table/tbody/tr[8]/td[1]"))).Click();
            test.Log(Status.Info, "Select HEX code");

            // Click on Save button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/div/button[1]"))).Click();
            test.Log(Status.Info, "Click on Save button");

            // Find the Latest entry
            var table = _driver.FindElements(By.XPath("//*[@id='grid_grid_records']/table"))[0];
            EventFiringWebDriver eventFiringWebDriver = new(_driver);
            eventFiringWebDriver.ExecuteScript("document.querySelector('#grid_grid_records').scrollTop=document.getElementById('grid_grid_records').scrollHeight");
            IList<IWebElement> tableRow = table.FindElements(By.TagName("tr"));
            var rows = tableRow.Where(x => x.GetAttribute("recid") != null).ToList();
            IWebElement latestEntry = rows.LastOrDefault();
            latestEntry.Click();
            string getData = latestEntry.Text;

            if (getData.Contains("Black Roof Color"))
            {
                test.Log(Status.Info, " Verify the new Cloned color is applied to the Colors list");
            }
            else
            {
                test.Log(Status.Info, " Verify the new Cloned color is not applied to the Colors list");
            }

            // Click on Sheathing Tab 
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol[1]/li[5]")));
            builder().Click(element).Perform();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.Id("grid_grid_records")));
            test.Log(Status.Info, "Click on Sheathing Tab");

            // Click on First row of Sheathing table
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_body']/div[1]/table/tbody/tr[3]"))).Click();
            test.Log(Status.Info, "Click on First row of Sheathing table");

            // Table Scroll Down
            eventFiringWebDriver.ExecuteScript("document.querySelector('#grid_partcosts_records').scrollTop = document.getElementById('grid_grid_records').scrollHeight");

            // Verify new entries automatically added in Color Table
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("grid_partcosts_records")));
            string sheathing = element.Text;

            if (sheathing.Contains("Test Clone Color"))
            {
                test.Log(Status.Info, "Verify the New Cloned Color has been applied automatically to the list of colors");
            }
            else
            {
                test.Log(Status.Info, "Verify the New Cloned Color has been not applied automatically to the list of colors");
            }

            SaveButton();

            // Click on Start From Scratch
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Start from Scratch"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(3);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-page page-0'])[1]/div[1]")));
            test.Log(Status.Info, "Click on Start From Scratch");

            // Click on Color Tab
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-0'])[1]/div[1]/div[7]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Click on Color Tab");

            // Click on Roof Color Dropdown
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-0'])[1]/div[1]/div[8]/div[1]/div[1]/div[2]"))).Click();
            test.Log(Status.Info, "Click on Roof Color Dropdown");

            IWebElement tableData = _driver.FindElement(By.XPath("//div[@class='menu']/table/tbody"));
            IList<IWebElement> Table = tableData.FindElements(By.TagName("tr"));
            string a = "//div[@class='menu']/table/tbody/tr[";
            string b = "]/td[1]/div[1]/span[1]";
            var rowCount = Table.Count();

            for (int i = 25; i <= rowCount; i++)
            {
                string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                if (c.Contains("Red Roof Color"))
                {
                    element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b)));
                    builder().Click(element).Perform();
                    GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
                    Wait(2);
                    GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
                    GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));

                    if (c.Contains("Red Roof Color"))
                    {
                        test.Log(Status.Info, " Verify the new cloned color is present in the drop down list.");
                    }
                    else
                    {
                        test.Log(Status.Info, " Verify the new cloned color is not present in the drop down list.");
                    }
                    break;
                }
            }

            _driver.Navigate().GoToUrl("https://testpostframesolver.azurewebsites.net/");
            _driver.SwitchTo().Alert().Accept();
            Delete();
        }

        public void Delete()
        {
            test = extent.CreateTest("Delete");

            SetUpWizard();

            ColorTab();

            // Search Sku
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-toolbar-search']/table/tbody/tr/td[2]/input[1]")));
            element.SendKeys("Red Roof Color" + Keys.Enter);
            test.Log(Status.Info, "Search Sku in the Color Table");

            // Click first row of Color table
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_records']/table/tbody/tr[3]"))).Click();
            test.Log(Status.Info, "Click first row of Color table");

            // Click on Delete Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_delete"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-msg-body'])[1]")));
            test.Log(Status.Info, "Click on Delete Button");

            // Click on Yes Button
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Click on Yes Button");

            SaveButton();
        }

        private void ColorTab()
        {
            // Click on Color Tab 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol[1]/li[4]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.Id("grid_grid_records")));
            test.Log(Status.Info, "Click on Color Tab");
        }

        private void SaveButton()
        {
            // Save Changes in SetUp Wizard
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("btnFinish"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='w2ui-page page-0']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/div/button[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Click on Save Button");
        }

        public void SetUpWizard()
        {
            // Navigate to the setup wizard page
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]/ul/li[2]/a"))).Click();
            test.Log(Status.Info, "Click on Setting button ");
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            test.Log(Status.Info, "Click on Setup Wizard button ");
        }

        private void DeleteOldEntries()
        {
            ColorTab();
            try
            {
                IWebElement tableDelete = _driver.FindElement(By.XPath("//div[@id='grid_grid_records']/table/tbody"));
                IList<IWebElement> elements = tableDelete.FindElements(By.TagName("tr"));
                string a = "//div[@id='grid_grid_records']/table/tbody/tr[";
                string b = "]/td[1]/div[1]";
                var rowCountDelete = elements.Count() - 2;

                for (int i = 30; i <= rowCountDelete; i++)
                {
                    string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                    if (c.Contains("Red Roof Color"))
                    {
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_records']/table/tbody/tr[" + i + "]"))).Click();

                        // Click on Delete Button
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_delete"))).Click();
                        GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-msg-body'])[1]")));

                        // Click on Yes Button
                        element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
                        builder().MoveToElement(element).Click(element).Perform();

                        SaveButton();
                        SetUpWizard();
                        ColorTab();
                        break;
                    }
                }
            }

            catch (NoSuchElementException)
            {
                Console.WriteLine("Delete old Data");
            }
        }



        [OneTimeTearDown]
        public void ExtentClose()
        {
            SendEmail("Test Report");
            extent.Flush();
        }
    }
}