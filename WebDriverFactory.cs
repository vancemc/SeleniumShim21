using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Reflection;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.IE;
//using OpenQA.Selenium.Safari;
using System.Diagnostics;
using System.Threading;

namespace SeleniumShim21
{
    public enum WebBrowserType
    {
        Ie,
        Chrome,
        FireFox, //Safari
    };


    public class WebDriverFactory
    {
        private delegate IWebDriver CreateWebDriverDelegate();
        private static Dictionary<WebBrowserType, CreateWebDriverDelegate> _webDrivers;

        static WebDriverFactory()
        {
            Initialize();
        }

        private static void Initialize()
        {
            _webDrivers = new Dictionary<WebBrowserType, CreateWebDriverDelegate>
            {
                {WebBrowserType.Chrome, CreateChromeDriver},
                {WebBrowserType.FireFox, CreateFireFoxDriver},
                {WebBrowserType.Ie, CreateIeDriver}
            };
        }

        public static IWebDriver CreateWebDriver(Assembly executingAssembly, WebBrowserType webBrowserType, string webDriverEmbededResourceManifestName = "")
        {
            if (webBrowserType != WebBrowserType.FireFox)
            {
                ExtractWebDriverFromEmbeddedResource(executingAssembly, webDriverEmbededResourceManifestName);
            }

            IWebDriver webDriver = _webDrivers[webBrowserType].Invoke();

            return webDriver;
        }

        private static IWebDriver CreateIeDriver()
        {
            IWebDriver driver = new InternetExplorerDriver();
            return driver;
        }

        private static IWebDriver CreateChromeDriver()
        {
            IWebDriver driver = new ChromeDriver();
            return driver;
        }

        private static IWebDriver CreateFireFoxDriver()
        {
            IWebDriver driver = new FirefoxDriver();
            return driver;
        }

        private static void ExtractWebDriverFromEmbeddedResource(Assembly executingAssembly, string fullResourceManifestName)
        {
            KillRunningDrivers();

            var assemblyPath = executingAssembly.Location;

            var executingPath = Path.GetDirectoryName(assemblyPath);

            var webDriverFileName = GetWebDriverNameFromEmbeddedResourceName(fullResourceManifestName);

            var fullFileOutputPath = string.Format(@"{0}\{1}", executingPath, webDriverFileName);

            using (var fileStream = new FileStream(fullFileOutputPath, FileMode.Create))
            {
                using (Stream stream = executingAssembly.GetManifestResourceStream(fullResourceManifestName))
                {
                    BinaryReader reader = new BinaryReader(stream);

                    reader.BaseStream.Position = 0;

                    var bytesFromResource = reader.ReadBytes(1048);

                    while (bytesFromResource.Length > 0)
                    {
                        fileStream.Write(bytesFromResource, 0, bytesFromResource.Length);

                        bytesFromResource = reader.ReadBytes(1048);
                    }
                }
            }
        }

        public static void CopyWebDriverToExecutingAssemblyDirectory(Assembly assembly, string sourceWebDriverPath)
        {
            string destinationDirectory = Path.GetDirectoryName(assembly.Location);

            FileStream sourceWebDriverStreamIn = null;

            if (string.IsNullOrWhiteSpace(sourceWebDriverPath) == false)
            {
                if (File.Exists(sourceWebDriverPath))
                {
                    sourceWebDriverStreamIn = new FileStream(sourceWebDriverPath, FileMode.Open);
                }

                if (sourceWebDriverStreamIn != null)
                {
                    string webDriverFileName = Path.GetFileName(sourceWebDriverPath);

                    string fullDestinationFilePath = String.Format(@"{0}\{1}", destinationDirectory, webDriverFileName);

                    if (File.Exists(fullDestinationFilePath) == false)
                    {
                        using (var webDriverStreamOut = new FileStream(fullDestinationFilePath, FileMode.Create))
                        {
                            sourceWebDriverStreamIn.CopyTo(webDriverStreamOut);
                        }
                    }
                }
                else
                {
                    throw new Exception(String.Format("Attempt to copy web driver file {0} failed.", sourceWebDriverPath));
                }
            }
        }

        public static string GetWebDriverNameFromEmbeddedResourceName(string embeddedResourceName)
        {
            if (String.IsNullOrWhiteSpace(embeddedResourceName))
            {
                throw new Exception(String.Format("A valid embedded resource name must be provided."));
            }

            string[] stringSegments = embeddedResourceName.Split('.');

            string fullFileName;

            if (stringSegments.Length > 1)
            {
                string fileName = stringSegments[stringSegments.Length - 2];

                string fileExtension = stringSegments[stringSegments.Length - 1];

                fullFileName = String.Format(@"{0}.{1}", fileName, fileExtension);
            }
            else
            {
                throw new Exception("Could not extract file name from embedded resource name.");
            }

            return fullFileName;
        }

        private static void KillRunningDrivers()
        {
            KillChromeDriver();
            KillIeDriver();
        }

        private static void KillChromeDriver()
        {
            KillWebDriver("chromedriver");
        }

        private static void KillIeDriver()
        {
            KillWebDriver("IEDriverServer");
        }

        private static void KillWebDriver(string processName)
        {
            try
            {
                var processes =
                    Process.GetProcessesByName(processName);

                processes.ToList().ForEach(p => p.Kill());

                Thread.Sleep(1000);
            }
            catch (Exception ex)
            {
                string msg =
                    String.Format("There was an error attempting to kill process named {0}", processName);

                Debug.WriteLine(msg);
                Debug.WriteLine(ex.Message);
            }
        }
    }
}
