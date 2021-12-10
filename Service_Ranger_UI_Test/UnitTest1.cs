using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UnitTesting.Logging;
using NLog.Config;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Enums;
using OpenQA.Selenium.Appium.Service;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Service_Ranger_UI_Test
{

    [TestClass]
    public class UnitTest1
    {
        private WindowsDriver<WindowsElement> driver;
        private Stopwatch _delta;
        private string _appId = @"C:\Program Files (x86)\Eaton\ServiceRanger4\Eaton.ServiceRanger.exe";
        private Dictionary<string, string> _applicationWindows = new Dictionary<string, string>();
        private static NLog.Logger _log;
        private static Logger log = new Logger();
        private Dictionary<string, Dictionary<string, AppiumWebElement>> _main = new Dictionary<string, Dictionary<string, AppiumWebElement>>();
        private Dictionary<string, AppiumWebElement> _appiumListElements = new Dictionary<string, AppiumWebElement>();

        [TestInitialize]
        public void SetUp()
        {
            string path = Directory.GetCurrentDirectory();
            Console.WriteLine();
            NLog.LogManager.Configuration = new XmlLoggingConfiguration(path + "/Nlog.config");
            _log =  NLog.LogManager.GetCurrentClassLogger();
            DateTime localDate = DateTime.Now;
            string _logpath = Path.Combine(path, localDate.ToString("yyyy_MM_dd_HH_mm_ss"));
            _delta = Stopwatch.StartNew();
            _delta.Start();
            driver = CreateNewSession(_appId);
            var element = WaitUntil("Read the Latest Service Bulletins...", "TextBlock");//We are on main Screen
            Assert.IsNotNull(element);
            _delta.Stop();
            _log.Info("time that it took");
        }
        [TestMethod]
        public void TestMethod1()
        {
            GetElements();
        }

        private WindowsDriver<WindowsElement> CreateNewSession(string appId, string argument = null)
        {
            var capabilities = new AppiumOptions();
            capabilities.AddAdditionalCapability(MobileCapabilityType.App, appId);
            capabilities.AddAdditionalCapability(MobileCapabilityType.PlatformName, "Windows");
            capabilities.AddAdditionalCapability(MobileCapabilityType.DeviceName, "WindowsPC");

            var _appiumLocalService = new AppiumServiceBuilder().UsingAnyFreePort().Build();
            _appiumLocalService.Start();
            var driver = new WindowsDriver<WindowsElement>(_appiumLocalService, capabilities);

            return driver;
         }

        private AppiumWebElement WaitUntil(string txt, string type)
        {
            AppiumWebElement _element = null;
            bool success = RetryUntilSuccessOrTimeOut( () =>
                           {
                               try
                               {
                                   List<string> openWindows = new List<string>(driver.WindowHandles);
                                   foreach(string openWindow in openWindows)
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
            List<string> types = new List<string> { "TextBlock" };
            Dictionary<string, AppiumWebElement> _elements = new Dictionary<string, AppiumWebElement>();

            List<string> openWindows = new List<string>(driver.WindowHandles);

            foreach(string openWindow in openWindows)
            {
                foreach(string type in types)
                {
                    IEnumerable<AppiumWebElement> elements = driver.FindElementsByClassName(type);
                    foreach (AppiumWebElement element in elements)
                    {
                        try
                        {
                            if (element != null && element.Text != null) _elements.Add(type, element); 
                        }
                        catch
                        {
                            //In case element is not completed
                        }
                    }
                    _main.Add(openWindow, _elements);
                }
            }

        }
    }
}
