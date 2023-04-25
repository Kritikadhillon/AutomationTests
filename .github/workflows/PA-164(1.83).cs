using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using AventStack.ExtentReports;

namespace SmartbuildAutomation
{
    class RoofMart : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        [Test,Order(1)]
        public void Roof()
        {
            test = null;
            test = extent.CreateTest("Verify the Assembly Drawing");

            SetUpWizard();
            DeleteOldEntries();
            // Click on Add Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//td[@id='tb_grid_toolbar_item_add']/table/tbody/tr/td/table/tbody/tr/td[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='w2ui-box1']")));
            test.Log(Status.Info, "Click on Add Button");

            // Enter SKU 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[1]/div[1]/input[1]"))).SendKeys("Wood ` Trading` 0 ` 2X0-4");
            Console.WriteLine("Enter the SKU in the: 'Wood ` Trading` 0 ` 2X0-4'");
            test.Log(Status.Info, "Add Sku with this symbol `");

            // Enter Description
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[2]/div[1]/input[1]"))).SendKeys(" Wood ` material ` Trading` 0 ` 2X0-4");
            Console.WriteLine("Enter the Description in the: ' Wood ` material ` Trading` 0 ` 2X0-4'");
            test.Log(Status.Info, "Add Description with this Symbol `");

            // Enter Width
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[3]/div[1]/input[1]"))).SendKeys("2");
            test.Log(Status.Info, "Enter Width");

            // Enter Depth
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[4]/div[1]/input[1]"))).SendKeys("6");
            test.Log(Status.Info, "Enter Depth");

            // select all usages from the usage box
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='UsagesItems']/div/div[3]/div[1]/table/tbody/tr[3]/td[1]")));
            builder().Click(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).Perform();
            test.Log(Status.Info, "select all usages from the usage box");

            // Click on Add button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_UsagesItemsGrid_toolbar_item_add"))).Click();
            test.Log(Status.Info, "Click on Add button");

            // Click on Save Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-buttons']/button[1]"))).Click();
            test.Log(Status.Info, "Click on Save Button");

            SaveTheChanges();

            // Click on Start From Scratch
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Start from Scratch"))).Click();
            PageLoader();

            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-page page-0'])[1]/div[1]")));
            test.Log(Status.Info, "Click on Start From Scratch");

            // Click on Details Tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@id='form_myForm_tabs'])[1]/table/tbody/tr/td[3]/div[1]"))).Click();
            test.Log(Status.Info, "Click on Details Tab");

            // Click on the wall girt framing tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-2'])[1]/div/div[13]"))).Click();
            test.Log(Status.Info, "Click on the wall girt framing tab");

            // Click on the girt material dropdown
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-2'])[1]/div/div[14]/div[3]/div[1]/div[2]"))).Click();
            test.Log(Status.Info, "Click on the girt material drop down");
            Wait(1);

            IWebElement tableData = _driver.FindElement(By.XPath("//div[@class='menu']/table/tbody"));
            IList<IWebElement> Table = tableData.FindElements(By.TagName("tr"));
            string a = "//div[@class='menu']/table/tbody/tr[";
            string b = "]";
            var rowCount = Table.Count();

            for (int i = 1; i <= rowCount; i++)
            {
                string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                if (c.Contains("Wood ` material ` Trading` 0 ` 2X0-4"))
                {
                    element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b)));
                    builder().Click(element).Perform();
                    PageLoader();
                    test.Log(Status.Info, "Select the newly created material");
                    break;
                }
            }

            // Click on Sync Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("serverRefresh"))).Click();
            PageLoader();
            test.Log(Status.Info, "Click on Sync Button ");

            // Click on Job Tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//td[@id='tabs_myForm_tabs_tab_tab5']/div[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-page page-5'])[1]/div[1]")));
            test.Log(Status.Info, "Click on Job Tab");

            // Enter Job Name
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-5'])[1]/div[1]/div[4]/div/input[1]"))).SendKeys("Assembly Drawing Data" + Keys.Enter);
            PageLoader();
            test.Log(Status.Info, "Enter Job Name");

            // Click on Save button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_editToolbar_item_save"))).Click();
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(2);
            test.Log(Status.Info, "Click on Save button");

            // Click on Output button
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.Id("totalPrice")));
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.Id("tb_editToolbar_item_outputs")));
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_editToolbar_item_outputs")));
            builder().MoveToElement(element).Click(element).Perform();
            Wait(2);
            test.Log(Status.Info, "Click on Output button");

            // Click on Assembly Drawing Checkbox

            string x = "(//div[@class='w2ui-page page-0'])[3]/div/div[2]/div[";
            string y = "]/div/span[1]";

            for (int i = 3; i <= 20; i++)
            {
                string c = x + i + y;
                element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(c)));
                string d = element.Text;
                if (d.Contains("Assembly Drawings"))
                {
                    GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-page page-0'])[3]/div/div[2]/div[" + i + "]/div/input[1]"))).Click();
                    break;
                }
            }
            test.Log(Status.Info, "Click on Assembly Drawing Checkbox");

            // Click on Download Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='w2ui-msg-buttons'])/div/button[3]"))).Click();
            test.Log(Status.Info, "Click on Download Button");

            // Click on Home Button
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='primaryToolbar']/table/tbody/tr/td[1]/table/tbody/tr/td[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Click on Home Button");

            ReadPDF();

        }

        [Test, Order(2)]
        public void Delete()
        {
            test = extent.CreateTest("Delete Data");
            SetUpWizard();
            FramingTab();

            // Search Sku
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-toolbar-search']/table/tbody/tr/td[2]/input[1]")));
            element.SendKeys("Wood ` Trading` 0 ` 2X0-4" + Keys.Enter);
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

            SaveTheChanges();

            // Click on Job Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("Jobs"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            test.Log(Status.Info, "Click on Job Button");

            IWebElement table = _driver.FindElement(By.XPath("//div[@id='grid_grid_records']/table/tbody"));
            IList<IWebElement> tableTr = table.FindElements(By.TagName("tr"));
            string x = "//div[@id='grid_grid_records']/table/tbody/tr[";
            string y = "]/td[2]/div";
            var rowCountForJob = tableTr.Count();

            for (int i = 3; i <= rowCountForJob; i++)
            {
                string c = x + i + y;
                element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(c)));
                string d = element.Text;
                if (d.Contains("Assembly Drawing Data"))
                {
                    element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(" //div[@id='grid_grid_records']/table/tbody/tr[" + i + "]/td[1]/div/button[4]")));
                    builder().DoubleClick(element).Perform();
                    break;
                }
            }

            // Click on Delete button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div/table/tbody/tr[3]"))).Click();
            test.Log(Status.Info, "Click on Delete button");

            // Click on Yes Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            test.Log(Status.Info, "Click on Yes Button");

        }
        public void SaveTheChanges()
        {
            // Save Changes in SetUp Wizard
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='footer-pagination footer-wizard-btn-wrap']/div[2]/button[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@class='w2ui-page page-0']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/div/button[1]"))).Click();
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]")));
            builder().MoveToElement(element).Click(element).Perform();
            test.Log(Status.Info, "Save Changes in SetUp Wizard");
        }
        private void SetUpWizard()
        {
            // Navigate to the setup wizard page
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]/ul/li[2]/a"))).Click();
            test.Log(Status.Info, "Click on the Setting drop-down");
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            test.Log(Status.Info, "Click on the setup wizard option");
        }
        public void PageLoader()
        {
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(3);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));
        }
        #region Read PDF File Data
        public string ReadPDF()
        {
            string result1 = "Wood , material , Trading, 0 , 2X0-4";
            string result2 = "Wood , Trading, 0 , 2X0-4";

            DateTime dateTime = DateTime.Now;
            string s;
            s = dateTime.ToString("MM-dd-yyyy");
            var downloadFolder = GetDownloadsFolder();
            var file = $@"{downloadFolder}\";
            string a = "Assembly Drawing Data_";
            var partialName = a + s;
            var hdDirectoryInWhichToSearch = new DirectoryInfo(@$"{file}");
            var filesInDir = hdDirectoryInWhichToSearch.GetFiles("*" + partialName + "*.*");
            var latestFile = filesInDir.OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
            string path = @$"{latestFile}";
            test.Log(Status.Info, "Verify PDF File is download");

            StringBuilder text = new StringBuilder();
            using (PdfReader reader = new PdfReader(path))
            {
                for (int i = 1; i <= reader.NumberOfPages; i++)
                {
                    text.Append(PdfTextExtractor.GetTextFromPage(reader, i));

                }
                string data = "Material List";
                string element = (text.ToString());
                Console.WriteLine(text.ToString().Contains(data));
                string[] results = new string[2] { result1, result2 };
                for (int i = 0; i < results.Length; i++)
                    if (element.Contains(results[i]))
                    {
                        Console.WriteLine("Verify the function replaces those icons ( ` ) with those icons ( , ) in the assembly drawing output." + " " + results[i] + " " + element.Contains(results[i]));
                        test.Log(Status.Info, "the function replaces those icons ( ` ) with those icons ( , ) in the assembly drawing output." + " " + results[i] + " " + element.Contains(results[i]));
                    }
                    else
                    {
                        Console.WriteLine("Verify the function is not replaces those icons ( ` ) with those icons ( , ) in the assembly drawing output." + " " + results[i] + " " + element.Contains(results[i]));
                        test.Log(Status.Info, "the function is not replaces those icons ( ` ) with those icons ( , ) in the assembly drawing output." + " " + results[i] + " " + element.Contains(results[i]));
                    }
                Wait(1);
                return text.ToString();

            }
        }
        #endregion

        private void FramingTab()
        {
            // Click on Framing Tab
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol[1]/li[2]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("//div[@id='grid_grid_toolbar']")));
            test.Log(Status.Info, "Click on Framing Tab");
        }

        private void DeleteOldEntries()
        {
            FramingTab();

            try
            {
                IWebElement tableDelete = _driver.FindElement(By.XPath("//div[@id='grid_grid_records']/table/tbody"));
                IList<IWebElement> elements = tableDelete.FindElements(By.TagName("tr"));
                string a = "//div[@id='grid_grid_records']/table/tbody/tr[";
                string b = "]/td[1]/div[1]";
                var rowCountDelete = elements.Count()-2;

                for (int i = 90; i <= rowCountDelete; i++)
                {
                    string c = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath(a + i + b))).Text;
                    if (c.Contains("Wood ` Trading` 0 ` 2X0-4"))
                    {
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='grid_grid_records']/table/tbody/tr[" + i + "]"))).Click();

                        // Click on Delete Button
                        GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_delete"))).Click();
                        GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='w2ui-msg-body'])[1]")));

                        // Click on Yes Button
                        element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='w2ui-msg-buttons']/button[1]")));
                        builder().MoveToElement(element).Click(element).Perform();

                        SaveTheChanges();
                        SetUpWizard();
                        FramingTab();
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