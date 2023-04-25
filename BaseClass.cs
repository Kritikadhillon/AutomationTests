using AventStack.ExtentReports;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;

namespace SmartbuildAutomation
{
    public class BaseClass
    {
        public IWebDriver _driver;
        public static ExtentTest test;
        public static ExtentReports extent;
        public IWebElement element;
        private const int maxWaitTime = 3;

        public WebDriverWait GetWebDriverWait() => new WebDriverWait(_driver, TimeSpan.FromMinutes(maxWaitTime));
        public Actions builder() => new Actions(_driver);
        public static readonly string reportpath = Directory.GetParent(@"../../../").FullName
            + Path.DirectorySeparatorChar + "Result"
            + Path.DirectorySeparatorChar + "Result_" + DateTime.Now.ToString("ddmmyyyy HHmmss");

        public void takeScreenShot()
        {
            string msg = "";
            Screenshot file = ((ITakesScreenshot)_driver).GetScreenshot();
            string image = file.AsBase64EncodedString;

            test.Fail(msg, MediaEntityBuilder.CreateScreenCaptureFromBase64String(image).Build());

        }

        public object SeleniumExtras { get; private set; }
        public int delay;

        [SetUp]
        public void EnvironmentSetup()
        {
           // new DriverManager().SetUpDriver(new ChromeConfig());
            ChromeOptions options = new ChromeOptions();
            options.AddArguments("--headless");
            options.AddArguments("--window-size=1920,1080");
            _driver = new ChromeDriver(options)
            {

                Url = "https://testpostframesolver.azurewebsites.net/Account/Login?ReturnUrl=%2F"
            };

            Console.WriteLine("Url is working successfully.");
            _driver.Manage().Window.Maximize();
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Name("Username"))).SendKeys("kdhillon@keymark.com");
            Console.WriteLine("Username entered successfully.");
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.Name("Password"))).SendKeys("Sss4321!");
            Console.WriteLine("Password entered successfully.");
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.CssSelector("#RememberMe"))).Click();
            Console.WriteLine("Remember me- details saved successfully");
            GetWebDriverWait().Until(ExpectedConditions.ElementToBeClickable(By.XPath("//*[@id='loginForm']/form/div[4]/div/input"))).Click();
            Console.WriteLine("Login Successfully");
        }

        public void Wait(int delay)
        {
            // Causes the WebDriver to wait for at least a fixed delay
            var now = DateTime.Now;
            var wait = new WebDriverWait(_driver, TimeSpan.FromSeconds(delay));
            wait.Until(_driver => (DateTime.Now - now) - TimeSpan.FromSeconds(delay) > TimeSpan.Zero);
        }

        [TearDown]
        public void CloseBrowser()
        {
            if (test != null)
            {

                SendEmail("Test Report");
                if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Passed)
                {
                    test.Log(Status.Pass, "Test Pass");
                    _driver.Quit();
                }

                else if (TestContext.CurrentContext.Result.Outcome.Status == TestStatus.Failed)
                {
                    test.Log(Status.Fail, "Test Fail");
                    takeScreenShot();
                    _driver.Quit();

                }
                else
                {
                    test.Log(Status.Skip, "Test skip");
                }
            }

            else
            {
                _driver.Quit();
            }

        }

        public static void SendEmail(string reportName)
        {

            var path = reportpath;
            var fromAddress = new MailAddress("dev@keymark.com", "From Name");
            var toAddress = new MailAddress("kdhillon@keymark.com", "To Name");
            string fromPassword = "dev12345678";
            string subject = reportName;

            var smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(fromAddress.Address, fromPassword)
            };

            using var message = new MailMessage(fromAddress, toAddress)
            {
                Subject = subject,
                Body = "Please find the attachment of this Test"
            };
            message.Attachments.Add(new Attachment($"{Directory.GetParent(@"../../../").FullName}{Path.DirectorySeparatorChar}Result\\index.html"));
            smtp.Send(message);
        }
        public static string GetPath()
        {
            var path = Directory.GetCurrentDirectory();
            path = path.Replace("\\bin\\Debug\\net5.0", "");
            return Path.Combine(path, "Images");
        }
        public static string GetFileData()
        {
            var path = Directory.GetCurrentDirectory();
            path = path.Replace("\\bin\\Debug\\net5.0", "");
            return Path.Combine(path, "download");
        }
        public static string ScreenshotFile()
        {
            var path = Directory.GetCurrentDirectory();
            path = path.Replace("\\bin\\Debug\\net5.0", "");
            return Path.Combine(path, "ScreenShot");
        }


        public static string GetDownloadsFolder()
        {
            string downloads;
            SHGetKnownFolderPath(KnownFolder.Downloads, 0, IntPtr.Zero, out downloads);
            return downloads;
        }

        public static class KnownFolder
        {
            public static readonly Guid Downloads = new Guid("374DE290-123F-4565-9164-39C4925E467B");
        }

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        static extern int SHGetKnownFolderPath([MarshalAs(UnmanagedType.LPStruct)] Guid rfid, uint dwFlags, IntPtr hToken, out string pszPath);
    }
}
