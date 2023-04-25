using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SmartbuildAutomation
{
    class BundleOption : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        [Test,Order(1)]
        public void PriceCheck()
        {
            test = extent.CreateTest("Options from Packages price is incorrect");

            PackageButton();
            ElementPresent();
            CreateOptionMethod();

            // get the price of catalog items(For Test 1)
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]/div/div[3]/div/div[11]/div/input")));
            string priceForTest1 = element.GetAttribute("value");
            Console.WriteLine("The price of catalog items for Test 1 is" + " " + priceForTest1);

            SaveButton();
            CreateBundleMethod();

            // Get the price of catalog items(For Test 2)
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]/div/div[3]/div/div[11]/div/input")));
            string priceForTest2 = element.GetAttribute("value");
            Console.WriteLine("The price of catalog items for Test 2 is" + " " + priceForTest2);

            SaveButton();
            SaveInPackage();

            // Click on Start from Scratch
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("Start from Scratch"))).Click();
            PageLoader();
            test.Log(Status.Info, "Click on Start from Scratch Button");

            // Click on Package Tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tabs_myForm_tabs_tab_tab6"))).Click();
            test.Log(Status.Info, "Click on Package Tab");
            Wait(2);

            // Click on the Group 1 Checkbox 
            IWebElement table = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_records']/table/tbody")));
            IList<IWebElement> list = table.FindElements(By.TagName("tr"));
            string a = "//div[@id='grid_packageList_records']/table/tbody/tr[";
            string b = "]/td[2]/div[1]";
            var rowCount = list.Count();

            for (int i = 3; i < rowCount; i++)
            {
                string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                if (c.Contains("Group 1"))
                {
                    GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_records']/table/tbody/tr["+i+"]/td[1]/div[1]"))).Click();
                    break;
                }
            }
            test.Log(Status.Info, "Click on the Group 1 Cehckbox ");

            // Select the “Make option” Option
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/table/tbody/tr[2]"))).Click();
            PageLoader();
            test.Log(Status.Info, "Select the “Make option” Option");

            JobReview();

            // Get the price of catalog items(For Test 1(Make Option))
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='grid_SummaryGrid_records']/table/tbody/tr[22]/td[3]/div")));
            string optionForTest1 = element.Text;

            // Get the price of catalog items(For Test 2(Make Option)
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='grid_SummaryGrid_records']/table/tbody/tr[23]/td[3]/div")));
            string optionForTest2 = element.Text;

            if ((priceForTest1 == optionForTest1) && (priceForTest2 == optionForTest2))
            {
                test.Log(Status.Info, "Verify that if the “Make Option” is selected into a Bundle then the price shows up correctly.");
                Console.WriteLine("Verify that if the “Make Option” is selected into a Bundle then the price shows up correctly." + priceForTest1 + "==" + optionForTest1 + " " + priceForTest2 + "==" + optionForTest2);
            }
            else
            {
                test.Log(Status.Info, "Verify that if the “Make Option” is selected into a Bundle then the price shows up incorrectly.");
                Console.WriteLine("Verify that if the “Make Option” is selected into a Bundle then the price shows up incorrectly." + priceForTest1 + "==" + optionForTest1 + " " + priceForTest2 + "==" + optionForTest2);
            }

            // Click on 3D View button
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_middleToolBar_item_show3dView")));
            builder().Click(element).Perform();

            for (int i = 3; i < rowCount; i++)
            {
                string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                if (c.Contains("Group 1"))
                {
                    GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_records']/table/tbody/tr[" + i + "]/td[1]/div[1]"))).Click();
                    break;
                }
            }
            test.Log(Status.Info, "Click on the Group 1 Cehckbox ");

            // Select the “Make Bundle” Option
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/table/tbody/tr[1]"))).Click();
            PageLoader();
            test.Log(Status.Info, "Select the “Make Bundle” Option");

            JobReview();

            // Get the price of catalog items(For Test 1(Make Option))
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='grid_SummaryGrid_records']/table/tbody/tr[21]/td[3]/div")));
            string bundleForTest1 = element.Text;

            // Get the price of catalog items(For Test 2(Make Option)
            element = GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@id='grid_SummaryGrid_records']/table/tbody/tr[23]/td[3]/div")));
            string bundleForTest2 = element.Text;

            if ((priceForTest1 == bundleForTest1) && (bundleForTest2 == optionForTest2))
            {
                test.Log(Status.Info, "Verify that if the “Make Bundle” is selected into a Bundle then the price shows up correctly.");
                Console.WriteLine("Verify that if the “Make Bundle” is selected into a Bundle then the price shows up correctly." + priceForTest1 + "==" + bundleForTest1 + " " + priceForTest2 + "==" + bundleForTest2);
            }
            else
            {
                test.Log(Status.Info, "Verify that if the “Make Bundle” is selected into a Bundle then the price shows up incorrectly.");
                Console.WriteLine("Verify that if the “Make Bundle” is selected into a Bundle then the price shows up incorrectly." + priceForTest1 + "==" + bundleForTest1 + " " + priceForTest2 + "==" + bundleForTest2);
            }
        }

        [Test,Order(2)]
        public void Delete()
        {
            PackageButton();
            DeleteNewData();
            SaveInPackage();
        }

        private void ElementPresent()
        {
            // Delete the Latest Entries of Packages table
            IWebElement table = _driver.FindElement(By.XPath("//div[@id='grid_packageList_records']/table/tbody"));
            IList<IWebElement> tableTr = table.FindElements(By.TagName("tr"));
            string x = "//div[@id='grid_packageList_records']/table/tbody/tr[";
            string y = "]/td[3]/div[1]";
            var rowCountForJob = tableTr.Count();

            try
            {
                for (int k = 1; k <= 2;)
                {
                    for (int i = 3; i < rowCountForJob; i++)
                    {
                        string c = x + i + y;
                        element = _driver.FindElement(By.XPath(c));
                        string d = element.Text;
                        if (d.Contains("Group 1"))
                        {
                            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_records']/table/tbody/tr[" + i + "]/td[3]")));
                            builder().DoubleClick(element).Perform();
                            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_packageList_toolbar_item_w2ui-delete"))).Click();
                            Wait(1);
                            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
                            builder().MoveToElement(element).Click(element).Perform();
                            break;
                        }
                    }
                    k++;
                }
                SaveInPackage();
                PackageButton();
            }
            catch (NoSuchElementException)
            {
                Console.WriteLine("Already delete the Group 1 File");
            }
        }

        private void JobReview()
        {
            // Click on the Job Review button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_middleToolBar_item_showMaterials"))).Click();
            test.Log(Status.Info, "Click on the Job Review button");
        }

        private void DeleteNewData()
        {
            // Delete the Latest Entries of Packages table
            IWebElement table = _driver.FindElement(By.XPath("//div[@id='grid_packageList_records']/table/tbody"));
            IList<IWebElement> tableTr = table.FindElements(By.TagName("tr"));
            string x = "//div[@id='grid_packageList_records']/table/tbody/tr[";
            string y = "]/td[3]/div[1]";
            var rowCountForJob = tableTr.Count();

            for (int k = 1; k <= 2;)
            {
                for (int i = 3; i <= rowCountForJob; i++)
                {
                    string c = x + i + y;
                    element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(c)));
                    string d = element.Text;
                    if (d.Contains("Group 1"))
                    {
                        element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_records']/table/tbody/tr[" + i + "]/td[3]")));
                        builder().DoubleClick(element).Perform();
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_packageList_toolbar_item_w2ui-delete"))).Click();
                        Wait(1);
                        element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
                        builder().MoveToElement(element).Click(element).Perform();
                        break;
                    }
                }
                k++;
            }
        }

        private void SaveInPackage()
        {
            // Click on Save Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_baseLayout_top_toolbar_item_item6"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            test.Log(Status.Info, "Click on Save Button");
        }

        private void PackageButton()
        {
            // Click on Setting Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//ul[@class='nav navbar-nav navbar-right rightmenus']/li[2]/a[1]"))).Click();
            test.Log(Status.Info, "Click on Setting Button ");

            // Click on Packages Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Packages"))).Click();
            Wait(3);
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@id='grid_packageList_records']/table/tbody/tr[3]/td/div[1])[1]")));
            test.Log(Status.Info, "Click on Packages Button");
        }

        private void PageLoader()
        {
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(3);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));
            Wait(1);
        }

        private void CreateBundleMethod()
        {
            AddButton();
            BlankOption();
            GroupName();

            // Add test 1 in the package name field
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[2]/div/input")));
            builder().Click(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys("Test 2").Perform();
            test.Log(Status.Info, "Add test 1 in the package name field");

            SelectOption();
            AddCatalog();

            try
            {
                // Select Catalog Items 
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div/div/div[3]/div/div[4]/div/div[2]"))).Click();
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table/tbody/tr[6]"))).Click();
            }
            catch
            {
                // Select Catalog Items 
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0 rearrange-package-fields']/div/div[4]/div/div[2]"))).Click();
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table/tbody/tr[6]"))).Click();
            }
        }

        private void SelectOption()
        {
            // Select option in the package type dropdown
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[3]/div/div[2]"))).Click();
            Wait(1);
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/table/tbody/tr[3]"))).Click();
            test.Log(Status.Info, "Select option in the package type dropdown");
        }

        private void CreateOptionMethod()
        {
            AddButton();
            BlankOption();
            GroupName();

            // Add test 1 in the package name field
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[2]/div/input")));
            builder().Click(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys("Test 1").Perform();
            test.Log(Status.Info, "Add test 1 in the package name field");

            SelectOption();
            AddCatalog();

            try
            {
                // Select Catalog Items 
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div/div/div[3]/div/div[4]/div/div[2]"))).Click();
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table/tbody/tr[3]"))).Click();
            }
            catch
            {
                // Select Catalog Items 
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0 rearrange-package-fields']/div/div[4]/div/div[2]"))).Click();
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table/tbody/tr[3]"))).Click();
            }
        }

        private void SaveButton()
        {
            // Click on Save button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div/div/div[3]/div/div[1]/label"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/div/button[1]"))).Click();
            Wait(2);
        }

        private void GroupName()
        {
            // Add group 1 in the group name field
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[1]/div/input"))).SendKeys("Group 1" + Keys.Enter);
            Wait(1);
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[1]/label"))).Click();
            test.Log(Status.Info, "Add group 1 in the group name field");
        }

        private void BlankOption()
        {
            // Click on Blank Option
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/table/tbody/tr[1]"))).Click();
            test.Log(Status.Info, "Click on  Blank Option");
            Wait(1);
        }

        private void AddButton()
        {
            // Click on ADD Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_packageList_toolbar']/table/tbody/tr/td[1]"))).Click();
            test.Log(Status.Info, "Click on ADD Button");
        }

        private void AddCatalog()
        {
            // Click on Add Catalog button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@name='MaterialsGrid']/div/div[2]/table/tbody/tr/td[2]"))).Click();
            test.Log(Status.Info, "Click on Add Catalog button");
        }

        [OneTimeTearDown]
        public void ExtentClose()
        {
            SendEmail("Test Report");
            extent.Flush();
        }
    }
}