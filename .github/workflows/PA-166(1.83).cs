using AventStack.ExtentReports;
using NUnit.Framework;
using OpenQA.Selenium;
using SeleniumExtras.WaitHelpers;
using System;

namespace SmartbuildAutomation
{
    class ChangingRooFPitch : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        /// <summary>
        /// 1. Navigate to Framing Rules Page and Enter the any value in the Roof Pitch
        /// 2. Click on the Save button. 
        /// 3. Go to the Framing Rules page again
        /// 4. Verified that changes in roof pitch are applied after clicking on the save button in framing rules. 
        /// 5. Navigate to the Home page and click on the start from scratch button and 
        /// 6. Verified that the roof pitch is also updated in the default job.
        /// </summary>

        [Test]
        public void RoofPitch()
        {
            test = extent.CreateTest("Changing Roof Pitch");

            FramingRules();
            EditButton();

            // Enter In the Roof Pitch Field
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@id='grid_formGrid_records'])[1]/table/tbody/tr[12]/td[2]")));
            Wait(1);
            builder().DoubleClick(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys("6").Perform();
            Console.WriteLine("Enter 6 in the Roof pitch of Framing Rules");
            test.Log(Status.Info, " Enter in the Roof Pitch Field");

            SaveButton();
            EditButton();

            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@id='grid_formGrid_records'])[1]/table/tbody/tr[12]/td[2]/div[1]")));
            string roofPitchFramingRules = element.Text;

            if (roofPitchFramingRules.Contains("6"))
            {
                test.Log(Status.Info, "Verified that changes in roof pitch are applied after clicking on the save button in framing rules");
                Console.WriteLine("Verified that changes in roof pitch are applied after clicking on the save button in framing rules");
            }
            else
            {
                test.Log(Status.Info, "Verified that changes in roof pitch are not applied after clicking on the save button in framing rules");
                Console.WriteLine("Verified that changes in roof pitch are not applied after clicking on the save button in framing rules");
            }

            // Click on Home Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Home"))).Click();
            _driver.SwitchTo().Alert().Accept();
            test.Log(Status.Info, "Click on Home Button");

            // Click on Start from Scratch
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("Start from Scratch"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("drawingArea")));
            Wait(3);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.ClassName("w2ui-spinner")));
            Wait(2);
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.Id("middleToolbar")));
            Wait(1);
            test.Log(Status.Info, "Click on Start from Scratch Button");

            // Click on Building Size Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='layout_formLayout_panel_main']/div[4]/div/div[3]/div/div[1]"))).Click();
            test.Log(Status.Info, "Click on Building Size Button");

            // Get the Roof Pitch value in the Default Job
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='layout_formLayout_panel_main']/div[4]/div/div[3]/div/div[2]//div[9]/div/input")));
            string defaultRoofPitch = element.GetAttribute("value");
            Console.WriteLine("Default Job Roof Pitch value is: "+ defaultRoofPitch);
            test.Log(Status.Info, "Get the Roof Pitch value in the Default Job");

            if (roofPitchFramingRules.Contains("6"))
            {
                test.Log(Status.Info, "Verified that the roof pitch is also updated in the default job");
                Console.WriteLine("Verified that the roof pitch is also updated in the default job");
            }
            else
            {
                test.Log(Status.Info, "Verified that the roof pitch is not updated in the default job");
                Console.WriteLine("Verified that the roof pitch is not updated in the default job");
            }

            // Click on Home button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='tb_editToolbar_item_cancel']/table/tbody/tr/td/table/tbody/tr/td[2]"))).Click();
            test.Log(Status.Info, "Click on Home");

            FramingRules();
            EditButton();

            // Enter In the Roof Pitch Field
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@id='grid_formGrid_records'])[1]/table/tbody/tr[12]/td[2]/div[1]")));
            Wait(1);
            builder().DoubleClick(element).KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).SendKeys("4").Perform();
            SaveButton();
        }

        private void FramingRules()
        {
            // Click on Setting Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//ul[@class='nav navbar-nav navbar-right rightmenus']/li[2]/a[1]"))).Click();
            test.Log(Status.Info, "Click on Setting Button ");

            // Click on Framing Rule Button 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.LinkText("Framing Rules"))).Click();
            test.Log(Status.Info, "Click on Framing Rule Button");
        }

        private void SaveButton()
        {
            // Click on Save button  
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_baseLayout_top_toolbar_item_item6"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@class='table-responsive']/table/tbody/tr/td[4])[1]/a")));
            test.Log(Status.Info, "Click on Framing Rule Button");
        }

        private void EditButton()
        {
            // Click on Edit button  
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("(//div[@class='table-responsive']/table/tbody/tr/td[4])[1]/a"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementIsVisible(By.XPath("(//div[@id='grid_formGrid_records'])[1]")));
            test.Log(Status.Info, "Click on Edit button");
            Wait(1);
        }

        [OneTimeTearDown]
        public void ExtentClose()
        {
            SendEmail("Test Report");
            extent.Flush();
        }
    }
}