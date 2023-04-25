using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System.Collections.Generic;
using System.Linq;

namespace SmartbuildAutomation
{
    class AssemblyCrashes : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        /// <summary>
        /// 1. Navigate to SetUp Wizard and click on Sheathing Assemblies and add the new material data
        /// 2. Navigate to Home page and click on Start From Scratch 
        /// 3. Click on the details tab.
        /// 4. Click on the Wall sheathing tab
        /// 5. Verify that the New Sheathing Assemblies data present in the Wall Material and Wainscot dropdown.
        /// 6. Navigate to Setup Wizard and delete the new added data from the Sheathing Assemblies table.
        /// </summary>

        [Test, Order(1)]
        public void SheathingAssemblies()
        {
            test = extent.CreateTest(" Primary Material for an Assembly crashes the system");

            SetUpWizard();

            // Click on The Sheathing Assemblies
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol/li[23]"))).Click();
            test.Log(Status.Info, " Click on The Sheathing Assemblies");
            string[] report = new string[] { "'", "\"" };
            string[] sku = new string[2] { "Brick_bmp 2'x5'", "2\"x5\"Wood Light" };
            string[] primaryMaterial = new string[] { "6", "8" };

            for (int f = 0; f < sku.Length;)
            {
                for (int i = 0; i < sku.Length;)
                {
                    for (int j = 0; j < primaryMaterial.Length; j++)
                    {
                        // Click on the add button
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_add"))).Click();
                        test.Log(Status.Info, "Click on the add button");

                        // Enter Sku with this symbol '  
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[2]/div/input[1]"))).SendKeys(sku[i]);
                        test.Log(Status.Info, "Enter Sku with this symbol" + "" + report[f]);

                        string x = "//div[@id='w2ui-overlay']/div/div[1]/table/tbody/tr[";
                        string y = "]";

                        // Select the Primary Material 
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-page page-0']/div/div[3]/div/div[2]"))).Click();
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(x + primaryMaterial[j] + y))).Click();
                        test.Log(Status.Info, "Select the Primary Material");
                        Wait(1);

                        // select all usages from the usage box
                        element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='UsagesItems']/div/div[3]/div[1]/table/tbody/tr[3]/td[1]")));
                        builder().Click(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).Perform();
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_UsagesItemsGrid_toolbar_item_add"))).Click();
                        test.Log(Status.Info, "select all usages from the usage box");

                        // Click on the Save button  
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/div[1]/button[1]"))).Click();
                        test.Log(Status.Info, "Click on the Save button ");
                        test.Log(Status.Info, " Verify the system doesn't crash after adding SKU with" + "" + report[f]);
                        i++;
                        f++;
                    }
                }
            }

            SaveTheChanges();

            // Click on Start from Scratch
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("Start from Scratch"))).Click();
            PageLoader();
            test.Log(Status.Info, "Click on Start from Scratch Button");

            // Click on the Details tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@id='form_myForm_tabs'])[1]/table/tbody/tr/td[3]"))).Click();
            test.Log(Status.Info, "Click on the Details tab");

            // Click on  the Wall Sheathing Tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-2'])[1]/div/div[35]"))).Click();
            test.Log(Status.Info, "Click on the Wall Sheathing tab");

            WallMaterial();
            AppliedSKU();
        }

        [Test,Order(2)]
        public void Delete()
        {
            SetUpWizard();
            DeleteEntries();
            SaveTheChanges();
        }


        /// <summary>
        /// Delete the new data from the Sheathing Assemblies table
        /// </summary>
        private void DeleteEntries()
        {
            // Click on The Sheathing Assemblies
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol/li[23]"))).Click();

            string[] sku = new string[2] { "2\"x5\"Wood Light", "Brick_bmp 2'x5'" };
            for (int n = 0; n < sku.Length; n++)
            {
                // Search Sku
                element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-toolbar-search']/table/tbody/tr/td[2]/input[1]")));
                element.SendKeys(sku[n] + Keys.Enter);

                // Click first row of table
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid']/div[1]/div[3]/div[1]/table/tbody/tr[3]"))).Click();

                // Click on Delete Button
                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_delete"))).Click();

                // Click on Yes Button
                element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
                builder().MoveToElement(element).Click(element).Perform();
            }
        }

        /// <summary>
        ///  Verify that the new Sheathing Assemblies data present in the Wall Sheathing 
        /// </summary>
        private void AppliedSKU()
        {
            IWebElement tableData = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table")));
            IList<IWebElement> Table = tableData.FindElements(By.TagName("tr"));
            string a = "//div[@id='w2ui-overlay']/div/div/table/tbody/tr[";
            string b = "]/td[1]";
            var rowCount = Table.Count();

            string[] wainscotdata = new string[2] { "Brick_bmp 2'x5'+", "2\"x5\"Wood Light+" };
            string[] symbol = new string[2] { "'", "\"" };
            for (int t = 1; t <= 2;)
            {
                for (int l = 0; l < symbol.Length;)
                {
                    for (int e = 0; e < wainscotdata.Length;)
                    {
                        Wainscot();
                        Wait(1);
                        for (int i = 1; i <= rowCount; i++)
                        {
                            string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                            if (c.Contains(wainscotdata[e]))
                            {
                                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Click();
                                PageLoader();

                                if (c.Contains(wainscotdata[e]))
                                {
                                    test.Log(Status.Info, "Verify that the selected Sheathing Assemblies SKU with" + " " + symbol[l] + " " + "is applied to the Wainscot Dropdown");
                                }
                                else
                                {
                                    test.Log(Status.Info, "Verify that the selected Sheathing Assemblies SKU with " + " " + symbol[l] + " " + " is not applied to the Wainscot Dropdown");
                                }
                                break;
                            }
                        }
                        test.Log(Status.Info, " Verify that the system doesn't crash after applying SKU with" + " " + symbol[l] + " " + "the material in the  Wainscot Dropdown");
                        e++;
                        l++;
                        t++;
                    }
                }
            }

            for (int t = 1; t <= 2;)
            {
                for (int l = 0; l < symbol.Length;)
                {
                    for (int e = 0; e < wainscotdata.Length;)
                    {
                        WallMaterial();
                        Wait(1);
                        for (int i = 1; i <= rowCount; i++)
                        {
                            string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                            if (c.Contains(wainscotdata[e]))
                            {
                                GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Click();
                                PageLoader();

                                if (c.Contains(wainscotdata[e]))
                                {
                                    test.Log(Status.Info, "Verify that the selected Sheathing Assemblies SKU with" + " " + symbol[l] + " " + "is applied to the Wall Material Dropdown");
                                }
                                else
                                {
                                    test.Log(Status.Info, "Verify that the selected Sheathing Assemblies SKU with " + " " + symbol[l] + " " + " is not applied to the Wall Material Dropdown");
                                }
                                break;
                            }
                        }
                        test.Log(Status.Info, " Verify that the system doesn't crash after applying SKU with" + " " + symbol[l] + " " + "the material in the  Wall Material Dropdown");
                        e++;
                        l++;
                        t++;
                    }
                }
            }
        }

        private void Wainscot()
        {
            // Click on the Wainscot Dropdown
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-2'])[1]/div/div[36]/div[6]/div/div[2]"))).Click();
            Wait(1);
            test.Log(Status.Info, "Click on the Wainscot Dropdown");
        }

        private void WallMaterial()
        {
            // Click on  the Wall Material Dropdown
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-2'])[1]/div/div[36]/div[1]/div/div[2]"))).Click();
            test.Log(Status.Info, "Click on  the Wall Material Dropdown");
        }

        /// <summary>
        /// Waiting for the load Model
        /// </summary>
        private void PageLoader()
        {
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));
            Wait(1);
        }

        /// <summary>
        /// Navigate to SetUp Wizard
        /// </summary>
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

        /// <summary>
        /// Save the Changes in the SetUp Wizard
        /// </summary>
        public void SaveTheChanges()
        {
            // Save Changes in SetUp Wizard
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='footer-pagination footer-wizard-btn-wrap']/div[2]/button[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='w2ui-page page-0']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/div/button[1]"))).Click();
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Save Changes in SetUp Wizard");
        }

        [OneTimeTearDown]
        public void ExtentClose()
        {
            SendEmail("Test Report");
            extent.Flush();
        }
    }
}