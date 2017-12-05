using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using System.Windows.Forms;
using System.Threading;
using System.Collections.ObjectModel;
using System.IO;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Interactions.Internal;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.HtmlControls;


namespace KeywordDrivenTest
{
    public static class UIOperation
    {
        static string oldWindowHandle;
        static string newWindowHandle;
        //default screen resolution is 1920*1080
        static Size screenResolution = new Size(1920, 1080);   

        public static bool PerformActions(string[] testData)
        {
            string[] testDataRow = testData;
            try
            {                              
                //translate array into selenium language
                TranslateTestData(testDataRow);                               
                return true;
            }
            catch (Exception e)
            {
                TestExecuter.sw.WriteLine(@"Test failed in step <"+ testDataRow[1] + ">, Reason: " + e.ToString());
                return false;
            }
          
        }

        private static void TranslateTestData(string[] testDataRow)
        {
            
            string[] dataRow = testDataRow;
            //check key word
            string keyWord = dataRow[2].Trim();
            //text of web element
            string elementText;
            //if steps begins with %, skip it
            if (dataRow[1].Trim().Substring(0, 1) == "%")
                keyWord= "%";
            switch (keyWord)
            {
                //if key word is "SetScreenResolution"
                case "SetScreenResolution":
                    string resolution = dataRow[5].Trim();
                    string[] resolutionArray = resolution.Split(new Char[] { ',', ';','*'});
                    screenResolution = new Size(Convert.ToInt32(resolutionArray[0].Trim()), Convert.ToInt32(resolutionArray[1].Trim()));
                    if (TestExecuter.driver != null)
                    {
                        TestExecuter.driver.Manage().Window.Size = screenResolution;
                        TestExecuter.driver.Manage().Window.Position = new Point(0, 0);
                    }
                    break;
                //if key word is "OpenBrowser"
                case "OpenBrowser":
                     //check Para
                     string browserName = dataRow[5].Trim().ToUpper();

                     if (browserName == "IE") //open IE
                     {
                        var options = new InternetExplorerOptions
                        {
                            BrowserAttachTimeout = TimeSpan.FromSeconds(30),
                            RequireWindowFocus = false,
                            IntroduceInstabilityByIgnoringProtectedModeSettings = true,
                            IgnoreZoomLevel = true,
                            EnsureCleanSession = true,
                            //PageLoadStrategy = InternetExplorerPageLoadStrategy.Normal,
                            //ValidateCookieDocumentType = true,
                            InitialBrowserUrl = "about:Tabs",
                            BrowserCommandLineArguments = "-noframemerging"
                        };
                        TestExecuter.driver = new InternetExplorerDriver(@"..\..\..\KeywordDrivenTest\Bin\Debug",options, TimeSpan.FromSeconds(30));                       
                     }
                     else
                     {
                         if (browserName == "CHROME") //open Chrome
                         {
                            TestExecuter.driver = new ChromeDriver(@"..\..\..\KeywordDrivenTest\Bin\Debug");
                         }
                         else  //error message
                         {
                             TestExecuter.sw.WriteLine("Browser is not supported ! Please input 'IE' or 'Chrome' in Excel file. ");
                             TestExecuter.testResult = false;
                             break;
                         }                  
                     }
                    TestExecuter.driver.Manage().Window.Size = screenResolution;
                    TestExecuter.driver.Manage().Window.Position = new Point(0, 0);
                    // Set implicit wait timeouts to 60 secs
                    TestExecuter.driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(30);
                    ICapabilities capabilities = ((RemoteWebDriver)TestExecuter.driver).Capabilities;
                    TestExecuter.sw.WriteLine(@"Browser Version: <" + browserName + " " + capabilities.Version + ">");
                    break;
                //if key word is "NavigateTo"
                case "NavigateTo":
                    //go to URL 
                    TestExecuter.driver.Navigate().GoToUrl(dataRow[5].Trim());
                    break;
                //if key word is "SendKeys"
                case "SendKeys":
                    GetElement(dataRow[3], dataRow[4]).Clear();
                    ((IJavaScriptExecutor)TestExecuter.driver).ExecuteScript("arguments[0].style.opacity='1'", GetElement(dataRow[3], dataRow[4]));
                    GetElement(dataRow[3], dataRow[4]).SendKeys(dataRow[5].Trim());
                    break;
                //if key word is "SendKeys"
                case "SendControlKeys":
                    SendControlKeys(dataRow[3], dataRow[4], dataRow[5]);
                    break;
                //if key word is "Click"
                case "Click":
                        if (TestExecuter.driver != null)
                        {
                            //get text from what you click 
                            oldWindowHandle = TestExecuter.driver.CurrentWindowHandle;
                            string txt = GetElement(dataRow[3], dataRow[4]).Text;
                            ((IJavaScriptExecutor)TestExecuter.driver).ExecuteScript("arguments[0].click();", GetElement(dataRow[3], dataRow[4]));
                            //if it is not download operation
                            if (!txt.ToUpper().Contains("DOWNLOAD"))
                            {
                                //go to latest window
                                TestExecuter.driver.SwitchTo().Window(TestExecuter.driver.WindowHandles.Last());
                                newWindowHandle = TestExecuter.driver.CurrentWindowHandle;
                            }
                            //if it is open button to load client
                            if (txt.ToUpper().Equals("OPEN"))
                            {
                                //maximize windows                        
                                List<string> lstWindow = TestExecuter.driver.WindowHandles.ToList();
                                foreach (var handle in lstWindow)
                                {
                                    if (TestExecuter.driver.SwitchTo().Window(handle).Manage().Window.Size.Height != screenResolution.Height)
                                    {
                                        TestExecuter.driver.SwitchTo().Window(handle).Manage().Window.Size = screenResolution;
                                        TestExecuter.driver.SwitchTo().Window(handle).Manage().Window.Position = new Point(0, 0);
                                    }
                                }
                            }
                            txt = null;                           
                        }              
                                                   
                    break;
                //if key word is "DoubleClick"
                case "DoubleClick":
                    //get old window handle before double click
                    if (TestExecuter.driver != null)
                        oldWindowHandle = TestExecuter.driver.CurrentWindowHandle;
                    //create Actions object
                    Actions doubleClickBuilder = new Actions(TestExecuter.driver);
                    //double click action
                    doubleClickBuilder.DoubleClick(GetElement(dataRow[3], dataRow[4])).Build().Perform();
                    //get new window handle after click
                    TestExecuter.driver.SwitchTo().Window(TestExecuter.driver.WindowHandles.Last());                   
                    newWindowHandle = TestExecuter.driver.CurrentWindowHandle;
                    //max new window
                    if (oldWindowHandle != newWindowHandle)
                    {
                      TestExecuter.driver.Manage().Window.Maximize();                                                
                    }
                    break;
                //if key word is "MoveToElement"
                case "MoveToElement":
                    //create Actions object
                    Actions moveToBuilder = new Actions(TestExecuter.driver);
                    //move to element action
                    moveToBuilder.MoveToElement(GetElement(dataRow[3], dataRow[4])).Build().Perform();                                 
                    break;
                //if key word is "Wait"
                case "Wait":
                    int s = Convert.ToInt32(dataRow[5].Trim()); //s is second
                    Thread.Sleep(s * 1000);
                    break;
                //if key word is "RefreshPage"
                case "RefreshPage":
                    Actions actions = new Actions(TestExecuter.driver);
                    actions.KeyDown(OpenQA.Selenium.Keys.Control).SendKeys(OpenQA.Selenium.Keys.F5).Perform();                    
                    break;
                //if key word is "ScrollToElement"
                case "ScrollToElement":
                    ((IJavaScriptExecutor)TestExecuter.driver).ExecuteScript("arguments[0].scrollIntoView();", GetElement(dataRow[3], dataRow[4]));
                    break;
                //if key word is "GetElement(dataRow[3], dataRow[4])"
                case "SwitchToFrame":
                    TestExecuter.driver.SwitchTo().Frame(GetElement(dataRow[3], dataRow[4]));
                    break;
                //if key word is "VerifyElementExist"
                case "VerifyElementExist":
                    if (!GetElement(dataRow[3], dataRow[4]).Displayed)
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason: ");
                        TestExecuter.sw.WriteLine("Web Element" + dataRow[4].ToString() + "is not displayed !");
                        TestExecuter.testResult = false;
                    }
                    else
                    {
                        //if it is client selection page
                        if (dataRow[4].Trim() == ".//*[@id='client-selection-form']/div[1]/span[1]") 
                        {
                            if (TestExecuter.driver.FindElement(By.XPath(".//*[@id='client-selection-form']/div[1]/span[1]")).Text.Trim().Length < 28)
                            {
                                TestExecuter.sw.WriteLine("RealSuite Version: < XXX >");
                            }
                            else
                            {
                                TestExecuter.buildVersion = TestExecuter.driver.FindElement(By.XPath(".//*[@id='client-selection-form']/div[1]/span[1]")).Text.Trim().Substring(18, 10);
                                //write build version to log
                                TestExecuter.sw.WriteLine("RealSuite Version: <" + TestExecuter.buildVersion + ">");
                            }
                           
                            
                        }                        
                    }
                    break;
                //if key word is "VerifyValueGreaterThan"
                case "VerifyValueGreaterThan":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text;
                    if (Convert.ToDouble(elementText)<=Convert.ToDouble(dataRow[5]))
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] + " : " + dataRow[4].ToString() + "> is not greater than " + "'" + dataRow[5] + "'");
                        TestExecuter.testResult = false;
                    }
                    break;
                //if key word is "VerifyValueLessThan"
                case "VerifyValueLessThan":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text;
                    if (Convert.ToDouble(elementText) >= Convert.ToDouble(dataRow[5]))
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] + " : " + dataRow[4].ToString() + "> is not less than " + "'" + dataRow[5] + "'");
                        TestExecuter.testResult = false;
                    }
                    break;
                //if key word is "VerifyValueEqualTo"
                case "VerifyValueEqualTo":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text.Trim();
                    if (elementText != dataRow[5].Trim())
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] + " : " + dataRow[4].ToString() + "> is not equal to " + "'" + dataRow[5] + "'");
                        TestExecuter.testResult = false;
                    }
                    break;
                //if key word is "VerifyValueNotEqualTo"
                case "VerifyValueNotEqualTo":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text.Trim();
                    if (elementText == dataRow[5].Trim())
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] + " : " + dataRow[4].ToString() + "> is equal to " + "'" + dataRow[5] + "'");
                        TestExecuter.testResult = false;
                    }
                    break;
                //if key word is "VerifyValueContain"
                case "VerifyValueContain":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text;
                    if (!elementText.Contains(dataRow[5]))
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1]+">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] +" : " + dataRow[4].ToString() + "> doesn't contain " +"'"+ dataRow[5]+"'");
                        TestExecuter.testResult = false;
                    }
                    break;
                //if key word is "VerifyValueNotContain"
                case "VerifyValueNotContain":
                    elementText = GetElement(dataRow[3], dataRow[4]).Text;
                    if (elementText.Contains(dataRow[5]))
                    {
                        TestExecuter.sw.WriteLine("Test failed in step: <" + dataRow[1] + ">, Reason:");
                        TestExecuter.sw.WriteLine("Web Element <" + dataRow[3] + " : " + dataRow[4].ToString() + "> doesn't contain " + "'" + dataRow[5] + "'");
                        TestExecuter.testResult = false;
                    }
                    break;               
                //if key word is "GoToNextWindow"
                case "GoToNextWindow":
                    TestExecuter.driver.SwitchTo().Window(newWindowHandle);
                    break;
                //if key word is "GoToNextWindow"
                case "GotoPreviousWindow":
                    TestExecuter.driver.SwitchTo().Window(oldWindowHandle);
                    break;
                //if key word is "CloseBrowser"
                case "CloseBrowser":
                    TestExecuter.driver.Close();
                    break;
               //if key word is "WinSaveFile" 
                case "WinSaveFile(IE11)":
                    string titleDownload = TestExecuter.driver.Title;
                    BrowserWindow browserWindowDownload = new BrowserWindow();
                    browserWindowDownload.SearchProperties[UITestControl.PropertyNames.ClassName] = "IEFrame";
                    browserWindowDownload.WindowTitles.Add(titleDownload);
                    WinToolBar notificationBar = new WinToolBar(browserWindowDownload);
                    notificationBar.SearchProperties.Add(WinToolBar.PropertyNames.Name, "Notification", PropertyExpressionOperator.EqualTo);
                    WinSplitButton saveButton = new WinSplitButton(notificationBar);
                    saveButton.SearchProperties.Add(WinButton.PropertyNames.Name, "Save", PropertyExpressionOperator.EqualTo);
                    Mouse.Click(saveButton);
                    break;
                //if key word is "WinClickOK" 
                case "WinClickOK(IE11)":
                    string titlePop = TestExecuter.driver.Title;
                    BrowserWindow browser = new BrowserWindow();
                    browser.SearchProperties[UITestControl.PropertyNames.ClassName] = "IEFrame";
                    browser.WindowTitles.Add(titlePop);
                    WinWindow pop = new WinWindow(null);
                    pop.SearchProperties.Add(WinWindow.PropertyNames.Name, "Message from webpage", PropertyExpressionOperator.EqualTo);
                    WinButton ok = new WinButton(pop);                   
                    ok.SearchProperties.Add(WinButton.PropertyNames.Name, "OK", PropertyExpressionOperator.EqualTo);
                    Mouse.Click(ok);
                    break;
                //if key word is "WinUploadFile"
                case "WinUploadFile":
                    WinWindow UploadWindow = new WinWindow(null);
                    UploadWindow.SearchProperties.Add("Name", "Choose File to Upload");
                    WinEdit FileNameEdit = new WinEdit(UploadWindow);
                    FileNameEdit.SearchProperties.Add("Name", "File name:");
                    WinButton OpenBtn = new WinButton(UploadWindow);
                    OpenBtn.SearchProperties.Add("Name", "Open");
                    Keyboard.SendKeys(FileNameEdit, @"C:\Users\pans\Downloads\53360.txt");
                    Mouse.Click(OpenBtn);
                    break;
                default:
                    break;
            }
        }

        private static IWebElement GetElement(string elementBy, string elementValue)
        {
            switch (elementBy.Trim())
            {
                case "Id":
                    return  TestExecuter.driver.FindElement(By.Id(elementValue));
                case "Name":
                    return TestExecuter.driver.FindElement(By.Name(elementValue));                    
                case "LinkText":
                    return TestExecuter.driver.FindElement(By.LinkText(elementValue));                  
                case "PartialLinkText":
                    return TestExecuter.driver.FindElement(By.PartialLinkText(elementValue));                   
                case "XPath":
                    return TestExecuter.driver.FindElement(By.XPath(elementValue));                    
                case "TagName":
                    return TestExecuter.driver.FindElement(By.TagName(elementValue));                   
                case "Class":
                    return TestExecuter.driver.FindElement(By.ClassName(elementValue));                  
                case "CSS":
                    return TestExecuter.driver.FindElement(By.CssSelector(elementValue));                    
                default:
                    return TestExecuter.driver.FindElement(By.Id(elementValue));                                 
            }
        }

        private static void SendControlKeys(string elementBy, string elementValue, string keyValue)
        {
            switch (keyValue.Trim().ToUpper())
            {
                case "TAB" :
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Tab);
                    break;
                case "ENTER":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Enter);
                    break;
                case "CONTROL":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Control);
                    break;
                case "ALT":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Alt);
                    break;
                case "DELETE":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Delete);
                    break;
                case "SHIFT":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Shift);
                    break;
                case "BACKSPACE":
                    GetElement(elementBy, elementValue).SendKeys(OpenQA.Selenium.Keys.Backspace);
                    break;               
            }
            
        }


    }
}
