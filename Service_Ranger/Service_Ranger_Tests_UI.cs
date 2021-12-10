using log4net;
using log4net.Appender;
using log4net.Core;
using log4net.Layout;
using log4net.Repository.Hierarchy;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Interactions;
using OpenQA.Selenium.Appium.Interfaces;
using OpenQA.Selenium.Appium.MultiTouch;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Interactions.Internal;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Service_Ranger
{
    [TestClass]
    public class Service_Ranger_Tests_UI
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private string _appId = @"C:\Program Files (x86)\Eaton\ServiceRanger4\Eaton.ServiceRanger.exe";
        private WindowsDriver<WindowsElement> driver;
        private Dictionary<string, Dictionary<string, List<AppiumWebElement>>> _main = new Dictionary<string, Dictionary<string, List<AppiumWebElement>>>();
        private Stopwatch _delta;

        [TestInitialize]
        public void SetUp()
        {
            _delta = Stopwatch.StartNew();
            
            SetUpLogger();

            //Start of Application
            _delta.Start();
            driver = CreateNewSession(_appId);
            var element = WaitUntil("Read the Latest Service Bulletins...", "TextBlock");//We are on main Screen
            Assert.IsNotNull(element);
            _delta.Stop();

            log.Info("Application has Succesfully Started");
            log.Info("Starting Time: " + (_delta.ElapsedMilliseconds / 1000) + " sec");
        }

        [TestCleanup]
        public void CleanUp()
        {
            if (driver != null)
            {
                driver.CloseApp();
                driver.Quit();
                driver = null;
            }
        }

        [TestMethod]
        [DoNotParallelize]
        [Ignore]
        public void Transition_Time_Between_Windows()
        {
            WindowsElement Button;

            Button = driver.FindElementByAccessibilityId("GoToButton");             //Main Menu Button
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");
           
            Button = driver.FindElementByXPath("/Window/Window/Button[10]");        //Comm Profiles Settings
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");

            Button = driver.FindElementByName("Add Profile");                       //Add new Profile Button
            Assert.IsTrue(Button.Text == "Add Profile");
            Button.Click();

            Button = driver.FindElementByAccessibilityId("Radio_ListBoxItem_1");    //Simulation Mode Selected
            Button.Click();
            log.Info("Radio Button Simulation " + " was Clicked");

            Button = driver.FindElementByAccessibilityId("ProfileName");            //Prfile Name Test Diego 2
            Button.Clear();
            Button.SendKeys("Test Diego");
            log.Info("Profile Name was set to " + " Test Diego Product");

            Button = driver.FindElementByAccessibilityId("SimProducts");            //Select Product
            Button.Click();
            Button.SendKeys(Keys.Down);
            Button.SendKeys(Keys.Down);
            Button.SendKeys(Keys.Enter);
            log.Info("Product Selected");

            Button = driver.FindElementByAccessibilityId("ComboBox_1");             //Select Adapter
            Button.Click();
            Button.SendKeys(Keys.Down);
            Button.SendKeys(Keys.Enter);
            log.Info("Adapter Selected");

            Button = driver.FindElementByAccessibilityId("SaveEditButton");         //Save the Profile
            Button.Click();
            log.Info("Profile Saved");

            Button = driver.FindElementByAccessibilityId("GoToButton");             //Mega Menu
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");

            Button = driver.FindElementByXPath("/Window/Window/Button[1]");         //Home Button from Mega Menu
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");

            Button = driver.FindElementByAccessibilityId("view:TitleBarButton_1"); //Connect Button
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");

        }
        [TestMethod]
        public void Transition_Time()
        {
            WindowsElement Button;

            Button = driver.FindElementByAccessibilityId("view:TitleBarButton_1"); //Connect Button
            Button.Click();
            log.Info("Button " + Button.Text + " was Clicked");
        }

        private static void SetUpLogger()
        {
            string _path = Directory.GetCurrentDirectory() + "\\" + "TestLogs" + "\\" + DateTime.Now.ToString("ddMMyy_hhmmss") + ".txt";

            Hierarchy hierarchy = (Hierarchy)LogManager.GetRepository();

            PatternLayout patternLayout = new PatternLayout();
            patternLayout.ConversionPattern = "%date [%thread] %-5level %logger - %message%newline";
            patternLayout.ActivateOptions();

            RollingFileAppender roller = new RollingFileAppender();
            roller.AppendToFile = false;
            roller.File = _path;
            roller.Layout = patternLayout;
            roller.MaxSizeRollBackups = 5;
            roller.MaximumFileSize = "1GB";
            roller.RollingStyle = RollingFileAppender.RollingMode.Size;
            roller.StaticLogFileName = true;
            roller.ActivateOptions();
            hierarchy.Root.AddAppender(roller);

            MemoryAppender memory = new MemoryAppender();
            memory.ActivateOptions();
            hierarchy.Root.AddAppender(memory);

            hierarchy.Root.Level = Level.Info;
            hierarchy.Configured = true;
        }

        private WindowsDriver<WindowsElement> CreateNewSession(string appId, string argument = null)
        {
            var capabilities = new AppiumOptions();
            capabilities.AddAdditionalCapability(MobileCapabilityType.App, appId);
            capabilities.AddAdditionalCapability(MobileCapabilityType.PlatformName, "Windows");
            capabilities.AddAdditionalCapability(MobileCapabilityType.DeviceName, "WindowsPC");

            var _appiumLocalService = new AppiumServiceBuilder().UsingAnyFreePort().Build();
            _appiumLocalService.Start();
            //var driver = new WindowsDriver<WindowsElement>(new Uri("http://127.0.0.1:49495/wd/hub"), capabilities);
            var driver = new WindowsDriver<WindowsElement>(_appiumLocalService, capabilities);
            return driver;
        }

        private AppiumWebElement WaitUntil(string txt, string type)
        {
            AppiumWebElement _element = null;
            bool success = RetryUntilSuccessOrTimeOut(() =>
            {
                try
                {
                    List<string> openWindows = new List<string>(driver.WindowHandles);
                    foreach (string openWindow in openWindows)
                    {
                        driver.SwitchTo().Window(openWindow);
                        IEnumerable<AppiumWebElement> elements = driver.FindElementsByClassName(type);
                        foreach (AppiumWebElement element in elements)
                        {
                            try
                            {
                                if (element.Text.ToLower() == txt.ToLower())
                                {
                                    _element = element;
                                    return true;
                                }
                            }
                            catch
                            {
                                continue;
                            }

                        }
                    }
                }
                catch (NoSuchElementException)
                {
                    //We should ignore this since we are waiting for the item;
                }
                catch (NotImplementedException)
                {
                    //
                }
                return false;
            }, TimeSpan.FromSeconds(60));

            return _element;

        }

        private bool RetryUntilSuccessOrTimeOut(Func<bool> task, TimeSpan timeSpan)
        {
            bool success = false;
            int elapsed = 0;
            while ((!success) && (elapsed < timeSpan.TotalMilliseconds))
            {
                Thread.Sleep(500);
                elapsed += 500;
                success = task();
            }
            return success;
        }

        private void GetElements()
        {
            List<string> types = new List<string> { "TextBlock", "Button" };
            Dictionary<string, List<AppiumWebElement>> _elements = new Dictionary<string, List<AppiumWebElement>>();
            List<AppiumWebElement> _elements_;

            List<string> openWindows = new List<string>(driver.WindowHandles);

            foreach (string openWindow in openWindows)
            {
                foreach (string type in types)
                {
                    IEnumerable<AppiumWebElement> elements = driver.FindElementsByClassName(type);
                    IEnumerable<AppiumWebElement> elements2 = driver.FindElementsByXPath("//*");
                    foreach (AppiumWebElement element in elements2)
                    {
                        try
                        {
                            if (element != null && element.Text != null && element.Id != null && element.TagName != null)
                            {
                                log.Info("**Added Element " + " " + type + "-->" + " with text --> " + element.Text);
                                log.Info("**Added Element " + " " + type + "-->" + " with Id --> " + element.Id);
                                log.Info("**Added Element " + " " + type + "-->" + " with Tag Name --> " + element.TagName);
                            }
                        }
                        catch
                        {
                            log.Info("Missed one element");
                        }

                    }
                    _elements_ = new List<AppiumWebElement>();
                    foreach (AppiumWebElement element in elements)
                    {
                        try
                        {
                            if (element != null && element.Text != null && element.Id != null && element.TagName!=null)
                            {
                                _elements_.Add(element);
                                log.Info("Added Element " + " "+  type + "-->"  + " with text --> " + element.Text);
                                log.Info("Added Element " + " " + type + "-->" + " with Id --> " + element.Id);
                                log.Info("Added Element " + " " + type + "-->" + " with Tag Name --> " + element.TagName);
                            }
                            
                        }
                        catch
                        {
                            //In case element is not completed
                        }
                    }
                    _elements.Add(type, _elements_);
                }
                _main.Add(openWindow, _elements);
            }

        }

        public void WaitForElement(string IDType, string elementName, int time)
        {
            var wait = new DefaultWait<WindowsDriver<WindowsElement>>(driver)
            {
                Timeout = TimeSpan.FromSeconds(time),
                PollingInterval = TimeSpan.FromSeconds(0.5)
            };

            wait.IgnoreExceptionTypes(typeof(InvalidOperationException));

            wait.Until(driver =>
            {
                int elementCount = 0;
                switch (IDType)
                {
                    case "id":
                        elementCount = driver.FindElementsByAccessibilityId(elementName).Count;
                        break;
                    case "xpath":
                        elementCount = driver.FindElementsByXPath(elementName).Count;
                        break;
                    case "name":
                        elementCount = driver.FindElementsByName(elementName).Count;
                        break;
                }
                return elementCount > 0;
            });
        }
    }
}
