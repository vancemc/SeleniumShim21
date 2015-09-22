using System;

namespace SeleniumShim21
{

    public abstract class WebPage
    {
        public WebBrowser Browser { get; set; }
        public string PageUrl { get; set; }

        /// <summary>
        /// Set page Title for implicit page validation when calling Load()
        /// </summary>
        public string Title { get; set; }

        protected WebPage(WebBrowser browser, string pageUrl="")
        {
            Browser = browser;
            PageUrl = pageUrl;
        }

        /// <summary>
        /// Loads web page
        /// </summary>
        /// <param name="validateTitle">Validates page title with expected title when set to true, no Title validation when set to false.</param>
        public virtual void Load(bool validateTitle=true)
        {
            if (Browser == null)
            {
                throw new Exception("WebBrowser object is null.");
            }

            Browser.Start(PageUrl);

            if (validateTitle)
            {
                VerifyPageTitle();
            }

            Title = Browser.WebDriver.Title;
        }

        public void Close()
        {
            try
            {
                Browser.WebDriver.Close();
            }
            finally
            {

            }
        }

        public virtual void ValidatePage()
        {
            VerifyPageTitle();
        }

        /// <summary>
        /// Verifies page title.
        /// </summary>
        public virtual void VerifyPageTitle()
        {
            string actualPageTitle = Browser.WebDriver.Title;

            if (String.Compare(Title, actualPageTitle, StringComparison.OrdinalIgnoreCase) !=
                0)
            {
                string pageErrorMsg = string.Format("Acutal page title '{0}' did not match expected page title '{1}'",
                    actualPageTitle, Title);

                throw new Exception(pageErrorMsg);
            }
        }
    }
}
