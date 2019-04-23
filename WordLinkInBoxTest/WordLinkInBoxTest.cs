using System;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;

namespace WordLinkInBoxTest
{
    [TestClass]
    public class WordLinkInBoxTest
    {
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        private const string CalculatorAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";
        protected static WindowsDriver<WindowsElement> session;

        [TestInitialize]
        public void Initialize()
        {
            if (session == null)
            {
                // Create a new session to bring up an instance of the Word
                DesiredCapabilities appCapabilities = new DesiredCapabilities();
                appCapabilities.SetCapability("app", "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\WINWORD.EXE");
                appCapabilities.SetCapability("deviceName", "WindowsPC");

                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
                Assert.IsNotNull(session);

                //Word start with a process WindowsElement
                //When word loaded, that WindowsElement close and reopen a new WindowsElement.
                //Here we recreate a session with new WindowsElement.
                session.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1.5);
                session = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), appCapabilities);
            }
        }

        [TestMethod]
        [DataRow("选项", "领英功能1", "取消")]
        //[DataRow("Options", "Minus")]
        public void TestLinkInBox(string optionName, string linkinBoxName, string cancelButtonName)
        {
            session.FindElementByName(optionName).Click();
            try
            {
                session.FindElementByName(linkinBoxName);
            }
            catch (Exception)
            {
                throw new Exception("LinkIn box not found");
            }
            finally
            {
                //close option form
                session.FindElementByName(cancelButtonName).Click();
            }
        }

        [ClassCleanup]
        public static void TearDown()
        {
            // Close the application and delete the session
            if (session != null)
            {
                session.Quit();
                session = null;
            }
        }
    }
}
