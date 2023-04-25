using AventStack.ExtentReports;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.Events;
using SeleniumExtras.WaitHelpers;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SmartbuildAutomation
{
    #region Uploading Invalid HEX Code
    public class InvalidHEXCode : BaseClass
    {
        [OneTimeSetUp]
        public void ExtentStart()
        {
            AventStack.ExtentReports.Reporter.ExtentHtmlReporter htmlReport = new(reportpath);
            extent = new AventStack.ExtentReports.ExtentReports();
            extent.AttachReporter(htmlReport);
        }

        /// <summary>
        /// 1. Navigate to the SetUp Wizard and click on the Color tab
        /// 2. Create new color material for the job 
        /// 3. Click on the download button and also download the XLSX file 
        /// 4. Click on Upload Button and Upload the edited excel file with invalid HEX code
        /// 5. Click on Upload Button and Upload the edited excel file with valid HEX code
        /// 6. Verified that the XLSX file is successfully uploaded.
        /// 7. Click on the download button and also download the CSV file 
        /// 8. Click on Upload Button and Upload the edited CSV file with invalid HEX code
        /// 9. Click on Upload Button and Upload the edited CSV file with valid HEX code
        /// 10.Verified that the CSV file is successfully uploaded. 
        /// </summary>

        [Test]
        public void HEXColorCode()
        {
            test = extent.CreateTest("Invalid HEX Code").Info(" Login to AUTOTEST_PHTEST Distributor for Test Environment.");

            // Navigate to the setup wizard page
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//form[@id='logoutForm']/ul/li[2]/ul/li[2]/a"))).Click();
            test.Log(Status.Info, "Navigate to the setup wizard page");

            ColorData();

            // Click on Download Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_download"))).Click();
            test.Log(Status.Info, "Click on Download Button");

            // Click on XLSX button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='menu']/table/tbody/tr[2]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            test.Log(Status.Info, "Click on XLSX button and Download the XLSX File");

            ExcelData();
            CSVData();
        }

        #region Edit the Excel data 
        public void ExcelData()
        {
            var downloadFolder = GetDownloadsFolder();
            Wait(1);
            var file = $@"{downloadFolder}\";
            Wait(1);
            var partialName = "SetupWizard-Colors.xlsx";
            var hdDirectoryInWhichToSearch = new DirectoryInfo(@$"{file}");
            var filesInDir = hdDirectoryInWhichToSearch.GetFiles("*" + partialName + "*.*");
            var latestFile = filesInDir.OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
            string path = $"{latestFile}";
            test.Log(Status.Info, "Verify that the XLSX file is downloaded");
            Wait(1);

            // Open Excel File 
            XSSFWorkbook workbook = new XSSFWorkbook(File.Open(path, FileMode.Open));

            // Get the First Sheet of excel
            ISheet sheet = workbook.GetSheetAt(0);

            // Get the last row of excel sheet
            int lastRowNum = sheet.LastRowNum;
            string RowL = lastRowNum.ToString();
            string a = "Test12";

            // Edit the last row of 3 column
            XSSFRow dataRow = (XSSFRow)sheet.GetRow(lastRowNum);
            dataRow.Cells[2].SetCellValue(a);

            // Save the excel file 
            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }

            // Click on Upload Button and Upload the edited excel file with invalid HEX code 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_upload"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/input"))).SendKeys(path);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            Console.WriteLine("Click on Upload Button and Upload the edited excel file with invalid HEX code");
            test.Log(Status.Info, "Click on Upload Button and Upload the edited excel file with invalid HEX code ");

            // Get the error
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]")));
            string error = element.Text;
            Console.WriteLine($"{error}");
            test.Log(Status.Info, $"{error}");

            // Click on ok button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]"))).Click();
            test.Log(Status.Info, "Click on Ok button");

            XSSFWorkbook workbook1 = new XSSFWorkbook(File.Open(path, FileMode.Open));
            ISheet sheet1 = workbook1.GetSheetAt(0);

            int lastRowNum1 = sheet1.LastRowNum;
            string RowL1 = lastRowNum1.ToString();
            string b = "#456456";

            XSSFRow dataRow1 = (XSSFRow)sheet1.GetRow(lastRowNum1);
            dataRow1.Cells[2].SetCellValue(b);

            using (FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                workbook1.Write(stream);
            }

            // Click on Upload Button and Upload the edited Excel file with valid HEX code
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_upload"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/input"))).SendKeys(path);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            Console.WriteLine("Click on Upload Button and Upload the edited excel file with valid HEX code");
            test.Log(Status.Info, "Click on Upload Button and Upload the edited excel file with valid HEX code");

            // get the error
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]")));
            string popUp = element.Text;
            Console.WriteLine($"Message: {popUp}");

            // Click on ok button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]"))).Click();
            test.Log(Status.Info, "Click on Ok button");

            var tableData = _driver.FindElements(By.XPath("//*[@id='grid_grid_records']/table"))[0];
            EventFiringWebDriver eventFiringWebDriverData = new(_driver);
            eventFiringWebDriverData.ExecuteScript("document.querySelector('#grid_grid_records').scrollTop=document.getElementById('grid_grid_records').scrollHeight");
            IList<IWebElement> tableRowData = tableData.FindElements(By.TagName("tr"));
            var rowsData = tableRowData.Where(x => x.GetAttribute("recid") != null).ToList();
            IWebElement editedDataEntries = rowsData.LastOrDefault();
            editedDataEntries.Click();
            string getEditedData = editedDataEntries.Text;

            if (getEditedData.Contains(b))
            {
                test.Log(Status.Info, "Verify the updated excel file data is present in the Color table");
            }
            else
            {
                test.Log(Status.Info, " Verify the updated excel file data is not present in the Color table");
            }
        }
        #endregion
        #region Edit the CSV File data
        public void CSVData()
        {
            // Click on Download Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_download"))).Click();
            test.Log(Status.Info, "Click on Download Button");

            // Click on XLSX button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@class='menu']/table/tbody/tr[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            test.Log(Status.Info, "Click on CSV button and Download the CSV File");

            var downloadFolderForCSV = GetDownloadsFolder();
            Wait(1);
            var fileCSV = $@"{downloadFolderForCSV}\";
            Wait(1);
            var partialNameCSV = "SetupWizard-Colors.csv";
            var hdDirectoryInWhichToSearchCSV = new DirectoryInfo(@$"{fileCSV}");
            var filesInDirCSV = hdDirectoryInWhichToSearchCSV.GetFiles("*" + partialNameCSV + "*.*");
            var latestFileCSV = filesInDirCSV.OrderByDescending(x => x.LastWriteTime).FirstOrDefault();
            Wait(1);
            string pathOfCSV = $"{latestFileCSV}";
            test.Log(Status.Info, "Verify that the CSV file is downloaded");
            Wait(1);

            // Navigate to the directory where the CSV file is located
            _driver.Navigate().GoToUrl(pathOfCSV);
            _driver.SwitchTo().Alert().Dismiss();

            // Read the contents of the CSV file into a string array
            string[] lines = File.ReadAllLines(pathOfCSV);

            // Split the last line of the array into its individual columns
            string[] columns = lines.Last().Split(',');

            // Edit the values in the desired columns
            columns[2] = "Test125468";

            // Update the last line of the CSV file with the edited values
            lines[lines.Length - 1] = string.Join(",", columns);

            // Save the updated contents back to the CSV file
            File.WriteAllLines(pathOfCSV, lines);

            // Click on Upload Button and Upload the edited CSV file with invalid HEX code 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_upload"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/input"))).SendKeys(pathOfCSV);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            Console.WriteLine("Click on Upload Button and Upload the edited CSV file with invalid HEX code data");
            test.Log(Status.Info, "Click on Upload Button and Upload the edited CSV file with invalid HEX code data");

            // Verify the Error Message
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]")));
            string popUpMessage = element.Text;
            Console.WriteLine($"{popUpMessage}");
            test.Log(Status.Info, $"{popUpMessage}");

            // Click on ok button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]"))).Click();
            test.Log(Status.Info, "Click on Ok button");

            // Navigate to the directory where the CSV file is located
            _driver.Navigate().GoToUrl(pathOfCSV);
            _driver.SwitchTo().Alert().Dismiss();

            // Read the contents of the CSV file into a string array
            string[] lines1 = File.ReadAllLines(pathOfCSV);

            // Split the last line of the array into its individual columns
            string[] columns1 = lines1.Last().Split(',');

            // Edit the values in the desired columns
            columns1[2] = "#999999";

            // Update the last line of the CSV file with the edited values
            lines1[lines1.Length - 1] = string.Join(",", columns1);

            // Save the updated contents back to the CSV file
            File.WriteAllLines(pathOfCSV, lines1);
            Wait(2);

            // Click on Upload Button and Upload the edited CSV file with valid HEX code 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_upload"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div/div[1]/input"))).SendKeys(pathOfCSV);
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));
            Console.WriteLine("Click on Upload Button and Upload the edited CSV file with valid HEX code");
            test.Log(Status.Info, "Click on Upload Button and Upload the edited CSV file with valid HEX code");

            // Successful Popup Message
            element = GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[2]/div/div[1]")));
            string popUpMessages = element.Text;
            Console.WriteLine($"Message: {popUpMessages}");
            test.Log(Status.Info, $"Message: {popUpMessages}");

            // Click on ok button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-popup']/div[4]/button[1]"))).Click();
            test.Log(Status.Info, "Click on Ok button");

            var tableCSV = _driver.FindElements(By.XPath("//*[@id='grid_grid_records']/table"))[0];
            EventFiringWebDriver eventFiringWebDriverCSV = new(_driver);
            eventFiringWebDriverCSV.ExecuteScript("document.querySelector('#grid_grid_records').scrollTop=document.getElementById('grid_grid_records').scrollHeight");
            IList<IWebElement> tableRowCSV = tableCSV.FindElements(By.TagName("tr"));
            var rowsDataCSV = tableRowCSV.Where(x => x.GetAttribute("recid") != null).ToList();
            IWebElement editedDataEntriesCSV = rowsDataCSV.LastOrDefault();
            editedDataEntriesCSV.Click();
            string getEditedDataCSV = editedDataEntriesCSV.Text;

            if (getEditedDataCSV.Contains("#999999"))
            {
                test.Log(Status.Info, "Verify the updated CSV file data is present in the Color table");
            }
            else
            {
                test.Log(Status.Info, " Verify the updated CSV file data is not present in the Color table");
            }
        }
        #endregion
        #region Create New Color
        public void ColorData()
        {
            GetWebDriverWait().Until(ExpectedConditions.InvisibilityOfElementLocated(By.XPath("//div[@class='w2ui-lock-msg']")));
            GetWebDriverWait().Until(ExpectedConditions.ElementExists(By.XPath("//div[@class='tabs-wizrd']")));

            // Click on Color Tab 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='page']/div[1]/ol[1]/li[4]"))).Click();
            test.Log(Status.Info, "Click on Color Tab");

            // Click on Add Button
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Id("tb_grid_toolbar_item_add"))).Click();
            test.Log(Status.Info, "Click on Add Button");

            // Enter Color Name 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[2]/div/input[1]"))).SendKeys("Testing HEX Code");
            test.Log(Status.Info, "Enter Color Name ");

            // Enter Color Code
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[4]/div/input[1]"))).SendKeys("Test Color Code");
            test.Log(Status.Info, "Enter Color Code");

            // Select HEX code 
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='form']/div/div[3]/div/div[3]/div/input[1]"))).Click();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//div[@id='w2ui-overlay']/div[1]/div/table/tbody/tr[3]/td[1]"))).Click();
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

            if (getData.Contains("Testing HEX Code"))
            {
                test.Log(Status.Info, " Verify the new color is added to the Colors list");
            }
            else
            {
                test.Log(Status.Info, " Verify the new color is not added to the Colors list");
            }
        }
        #endregion

        [OneTimeTearDown]
        public void ExtentClose()
        {
            SendEmail("Test Report");
            extent.Flush();
        }
    }
}
#endregion

