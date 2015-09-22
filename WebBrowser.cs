using Microsoft.VisualBasic;
using System;
using System.Diagnostics;
using System.Threading;
using OpenQA.Selenium;
using System.Reflection;

namespace SeleniumShim21
{
    public class WebBrowser
    {
        private const WebBrowserType DefaultWebBrowserType = WebBrowserType.FireFox;

        private const int DefaultElementSearchTimeOut = 10000;
        private const int DefaultRetryDelayInMilliseconds = 1000;

        private string _sendKeysText;

        /// <summary>
        /// Implicit retry on timeout duration in milliseconds for FindElement() and ExecuteAction(). Default value is set to 10000ms (10 seconds)
        /// </summary>
        public int ElementSearchTimeout { get; set; }

        /// <summary>
        /// Retry delay to prevent CPU utilization from maxing out when internal retry loop searches for elements that have not been de-serialized by browser threads. Default is 1000ms (One second)
        /// </summary>
        public int RetryDelayInMilliseconds { get; set; }
        public enum UserAction
        {
            Click,
            Clear,
            TypeText
        };

        /// <summary>
        /// Public field to get at underlying WebDriver instance
        /// </summary>
        public IWebDriver WebDriver { get; set; }

        /// <summary>
        /// Overloaded constructor allows specifing the WebDriver (browser type) and WebDriver file path.
        /// Note: you can include WebDriver files (e.g. ChromeDriver.exe) as an embedded resource and 
        ///       stream it out to a destination directory. This works well with executing from a CICD pipeline.
        /// </summary>
        /// <param name="browserType">Default = Ie, Chrome or FireFox</param>
        /// <param name="webDriverFullFilePath">Full file path for Ie or Chrome drivers</param>
        public WebBrowser(Assembly executingAssembly, WebBrowserType browserType = DefaultWebBrowserType,  string webDriverEmbeddedResourceManifestName="")
            : this(browserType)
        {
            WebDriver = WebDriverFactory.CreateWebDriver(executingAssembly, browserType, webDriverEmbeddedResourceManifestName);
        }

        /// <summary>
        /// This constructor requires that Selenium Web Drivers already be copied to the excuting assembly path. 
        /// Use the WebDriverFactory to copy WebDrivers or to copy WebDrivers included in the VS project as embedded resources or
        /// use the overloaded constructor that takes a manifest resource name to automatically copy the WebDriver to 
        /// the executing directory.
        /// </summary>
        /// <param name="browserType"></param>
        public WebBrowser(WebBrowserType browserType = DefaultWebBrowserType)
        {
            ElementSearchTimeout = DefaultElementSearchTimeOut;
            RetryDelayInMilliseconds = DefaultRetryDelayInMilliseconds;
        }

        /// <summary>
        /// Navigates to a destination url
        /// </summary>
        /// <param name="url">A valid url</param>
        public virtual void Start(string url="")
        {
            if (string.IsNullOrWhiteSpace(url)==false)
            {
                WebDriver.Url = url;
                WebDriver.Navigate();
            }
        }

        /// <summary>
        /// Uses implicit retry until timeout for web elements. This overload finds element and sets the text.
        /// </summary>
        /// <param name="by">By locator</param>
        /// <param name="text">Text to send</param>
        /// <param name="includeReturnKeyPress">Appends a return key press to text</param>
        /// <returns></returns>
        public virtual object ExecuteAction(By by, string text, bool includeReturnKeyPress=false)
        {
            if (includeReturnKeyPress)
            {
                text = string.Format("{0}\n", text);
            }

            SendKeys(by, text, includeReturnKeyPress);

            return null;
        }

        /// <summary>
        /// Same as ClickElement(By) or ClearElement(By). Uses implicit retry until timeout for web elements. This overload finds element and executs a user action (Click or Clear)
        /// </summary>
        /// <param name="by">By locator</param>
        /// <param name="userAction">Click or Clear</param>
        public virtual void ExecuteAction(By by, UserAction userAction)
        {
            switch (userAction)
            {
                case UserAction.Click:
                     ClickElement(by);
                     break;
                case UserAction.Clear:
                     ClearElement(by);
                     break;
            }
        }

        /// <summary>
        /// SendKeys wrapper that includes implicit retry until timeout. Same as ExecuteAction(By, string, bool)
        /// </summary>
        /// <param name="by">By locator</param>
        /// <param name="text">Text to send</param>
        /// <param name="includeReturnKeyPress">Appends a return key press</param>
        /// <returns></returns>
        public virtual void SendKeys(By by, string text, bool includeReturnKeyPress=false)
        {
            var textElement = FindElement(by);

            _sendKeysText = text + (includeReturnKeyPress ? "\r\n" : "");

            Func<IWebElement, IWebElement> textElementDelegate = SendKeys;

            ImplicitWaitUntilTimeout<IWebElement, IWebElement>(textElementDelegate, textElement);
        }

        // Note: Private because this overloaded signature matches delegate Func<By, object> calling signature
        private IWebElement SendKeys(IWebElement textElement)
        {
            textElement.SendKeys(_sendKeysText);
            _sendKeysText = string.Empty;

            return textElement;
        }

        /// <summary>
        /// Sends a click to a click-able element
        /// </summary>
        /// <param name="by"></param>
        /// <returns>IWebElement</returns>
        public virtual IWebElement ClickElement(By by)
        {
            var clickableElement = FindElement(by);

            Func<IWebElement, IWebElement> clickElementDelegate = ClickElement;

            var element =
                ImplicitWaitUntilTimeout<IWebElement,IWebElement>(clickElementDelegate, clickableElement);

            return element;
        }

        /// <summary>
        /// Clears text based elements
        /// </summary>
        /// <param name="by"></param>
        /// <returns>IWebElement</returns>
        public virtual IWebElement ClearElement(By by)
        {
            var clearableElement = FindElement(by);

            Func<IWebElement, IWebElement> clearElementDelegate = ClearElement;

            var element =
                ImplicitWaitUntilTimeout<IWebElement, IWebElement>(clearElementDelegate, clearableElement);

            return element;
        }

        // For use as a delegate only
        private IWebElement ClickElement(IWebElement element)
        {
            element.Click();
            return element;
        }

        // For use as a delegate only
        private IWebElement ClearElement(IWebElement element)
        {
            element.Clear();
            return element;
        }

        /// <summary>
        /// FindElement() wrapper that includes implicit retry until timeout while browser is still in the de-serialization process
        /// </summary>
        /// <param name="by"></param>
        /// <returns></returns>
        public virtual IWebElement FindElement(By by)
        {
            Func<By, object> seleniumFindElement = WebDriver.FindElement;

            var element = ImplicitWaitUntilTimeout(seleniumFindElement, by) as IWebElement;

            return element;
        }

        // Note: purposely not using default parameter for ElementSearchTimeout 
        //       to allow use of property .vs a const required for default values.
        /// <summary>
        /// Use to verify that an element is visible. For negative verification (element is not visible), use overload to override the search timeout to shorten wait time.
        /// </summary>
        /// <param name="by">By locator</param>
        /// <returns>true if element found and element.Displayed property is set to true</returns>
        public virtual bool ElementIsVisible(By by)
        {
            var retVal = ElementIsVisible(by, ElementSearchTimeout);

            return retVal;
        }

        /// <summary>
        /// Use to verify that an element is visible. For negative verification (element is not visible), override the search timeout to shorten wait time.
        /// </summary>
        /// <param name="by">By locator<</param>
        /// <param name="overrideTimeout">Override timeout in milliseconds overrides ElementSearchTimeout property value</param>
        /// <returns>True if element is visible</returns>
        public virtual bool ElementIsVisible(By by, int overrideTimeout)
        {
            bool result = false;
            IWebElement element;
            int oldElementSearchTimeOut = ElementSearchTimeout;
            ElementSearchTimeout = overrideTimeout;

            try
            {
                element = FindElement(by);

                result = element != null && element.Displayed;
            }
            catch (Exception)
            {
                result = false;
            }
            finally
            {
                ElementSearchTimeout = oldElementSearchTimeOut;
            }

            return result;
        }

        /// <summary>
        /// Validates that an element's text field contains the expected text (case insensitive)
        /// </summary>
        /// <param name="by">Element with text field that contains expected text</param>
        /// <param name="expectedText">Expected text value</param>
        /// <returns>True if expected text matches element.text value</returns>
        public virtual bool TextIsVisible(By by, string expectedText)
        {
            var result = TextIsVisible(by, expectedText, ElementSearchTimeout);

            return result;
        }

        /// <summary>
        /// Validates that an element's text field contains the expected text (case insensitive)
        /// </summary>
        /// <param name="by">Element with text field that contains expected text</param>
        /// <param name="expectedText">Expected text value</param>
        /// <param name="overrideTimeout">Overrides WebBrowser.ElementSearchTimeout setting</param>
        /// <returns>True if expected text matches element.text value</returns>
        public virtual bool TextIsVisible(By by, string expectedText, int overrideTimeout)
        {
            bool result = false;
            string actualText = string.Empty;
            int oldElementSearchTimeOut = ElementSearchTimeout;
            ElementSearchTimeout = overrideTimeout;

            result = ElementIsVisible(by, overrideTimeout);

            var stopwatch = new Stopwatch();

            stopwatch.Start();

            try
            {
               if(result == true)
               {
                   do
                   {
                       var element = FindElement(by);

                       try
                       {
                           actualText = element.Text;

                           result = (expectedText.ToLower() == actualText.ToLower());
                       }
                       catch
                       {
                           result = false;
                       }

                   } while ((result==false) && (stopwatch.ElapsedMilliseconds < ElementSearchTimeout));
               }
            }
            catch
            {
                result = false;
            }
            finally
            {
                ElementSearchTimeout = oldElementSearchTimeOut;
            }

            return result;
        }


        public virtual string GetUserInput(string prompt = "SeleniumShim: Get User input",
                                   string title = "SeleniumShim 2.0",
                                   string defaultResponse = "")
        {
            var input = Interaction.InputBox(prompt, title, defaultResponse);

            return input;
        }

        private ReturnType ImplicitWaitUntilTimeout<ParamType, ReturnType>(Func<ParamType, ReturnType> methodDelegate, ParamType param)
        {
            ReturnType retVal = default(ReturnType);
            var timer = new Stopwatch();
            Exception exception = null;
            timer.Start();

            do
            {
                try
                {
                    retVal = methodDelegate(param);
                }
                catch (Exception ex)
                {
                    Thread.Sleep(RetryDelayInMilliseconds);
                    exception = ex;
                }
            } while (retVal == null && (timer.ElapsedMilliseconds < ElementSearchTimeout));

            if (retVal == null)
            {
                throw exception ?? new Exception("");
            }

            return retVal;
        }
    }
}
