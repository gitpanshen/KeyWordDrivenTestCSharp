using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Windows.Input;
using System.Windows.Forms;
using System.Drawing;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using OpenQA.Selenium.Support;
using OpenQA.Selenium;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace KeywordDrivenTest
{
    /// <summary>
    /// Summary description for CodedUITest1
    /// </summary>
    [CodedUITest]
    public class TestExecuter
    {
        public TestExecuter()
        {
        }

        //selenium driver
        public static IWebDriver driver = null;
        //Excel files for test cases
        string[] files;
        //test result for each test case
        public static bool testResult;
        //log file path
        public string logPath;
        //StreamWriter for log file     
        public static StreamWriter sw;
        //passed test case number
        int passedTCN;
        //failed test case number
        int failedTCN;
        //failed folder
        string failedFolder = @"C:\KeywordDriven\Failed";
        //passed folder
        string passedFolder = @"C:\KeywordDriven\Passed";
        //log folder
        string logFolder = @"C:\KeywordDriven\Log";
        //not executed folder
        string notExeFolder = @"C:\KeywordDriven\Not Execute";
        //test case folder
        string testCaseFolder = @"C:\KeywordDriven\Test Cases";
        //build version
        public static string buildVersion = "";
        //sub folder name
        string dirName = "";
      
        [TestInitialize]
        public void KeywordTestInitialize()
        {
            //check folders are there, otherwise create them
            System.IO.Directory.CreateDirectory(notExeFolder);
            System.IO.Directory.CreateDirectory(failedFolder);
            System.IO.Directory.CreateDirectory(passedFolder);
            System.IO.Directory.CreateDirectory(logFolder);
            System.IO.Directory.CreateDirectory(testCaseFolder);
            //clear passed and failed folders
            ClearFolder(passedFolder);
            ClearFolder(failedFolder);
            //get all files from "not executed" folder
            files = Directory.GetFiles(notExeFolder);
            //create a log file
            string logFileName = "KeywordDrivenTestLog-" + System.Environment.MachineName + "-" + DateTime.Now.ToString("yyyyMMddHHmmss") +".txt";
            logPath = System.IO.Path.Combine(logFolder,logFileName);
            //write the file
            sw = File.CreateText(logPath);
            sw.WriteLine(@"Keyword Driven Test began at: " + DateTime.Now.ToString());
            sw.WriteLine(@"Test Environment: " + System.Environment.MachineName);
            //set passed and failed test cases to 0
            passedTCN = 0;
            failedTCN = 0;
            //clear driver if exits
            if (driver != null)
                driver.Dispose();      
        }

        [TestMethod]
        [TestCategory("Keyword Driven Test")]      
        public void KeywordDrivenTestMethod()
        {
            // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
            //get all folders under "C:\KeywordDriven\Not execute"
            string dirPath = notExeFolder;
            List<string> dirs = new List<string>(Directory.EnumerateDirectories(dirPath));
            ProcessAllFiles();
            // if sub folder exists            
            foreach (var dir in dirs)
            {
                //get folder name
                dirName = new DirectoryInfo(dir).Name;
                //create corresponding folder under "Passed" and "Failed" folders
                System.IO.Directory.CreateDirectory(failedFolder+"\\"+dirName);
                System.IO.Directory.CreateDirectory(passedFolder+"\\"+dirName);
                //get files in one folder
                files = Directory.GetFiles(dir);
                //read all Excel files in folder
                ProcessAllFiles();                
            }
        }

        [TestCleanup]
        public void KeywordTestCleanup()
        {
            //delete all empty folders
            DeleteEmptySubdirectories(notExeFolder);
            DeleteEmptySubdirectories(failedFolder);
            DeleteEmptySubdirectories(passedFolder);
            //write log
            sw.WriteLine("----------------------------------------------------------------------------------------------");
            sw.WriteLine(@"Keyword Driven Test finished at: " + DateTime.Now.ToString());
            sw.WriteLine(@"Passed test cases number:" + passedTCN);
            sw.WriteLine(@"Failed test cases number:" + failedTCN);
            if ((passedTCN==0) & (failedTCN==0))
               sw.WriteLine(@"No test case is executed! Did you put correct test case file (xlsx) into 'Not execute' folder ?" );
            sw.Close();
            //clear all cache
            if (driver != null )
              driver.Quit();
        }

        //process all files (test cases) one by one under the folder
        private void ProcessAllFiles()
        {
            //read all Excel files in folder
            foreach (string file in files)
            {
                testResult = true;
                //get the test data array from Excel file
                string[,] testDataArray = LoadExcelFile.ImportSheet(file);
                // define 1-d array to save each row of array
                string[] row = new string[testDataArray.GetLength(1)];
                // write test case name to log
                sw.WriteLine("----------------------------------------------------------------------------------------------");
                sw.WriteLine("Test Case Path: " + file);
                sw.WriteLine("Test Case Name: <" + testDataArray[1, 0].ToString() + ">");
                //for each test data array
                for (int i = 1; i < testDataArray.GetLength(0); i++)
                {
                    //get one row from array
                    for (int j = 0; j < testDataArray.GetLength(1); j++)
                        row[j] = testDataArray[i, j];

                    //if any action failed, stop running this test case, set test result as "failed"  
                    if (UIOperation.PerformActions(row) == false)
                    {
                        testResult = false;
                        break;
                    }
                }              
                //clear array
                Array.Clear(testDataArray, 0, testDataArray.Length);
                //clear driver
                if (driver != null)
                {
                   driver.Quit();
                }
                //move file into "Passed" or "Failed" folder 
                MoveFile(file);
            }
        }

        private void MoveFile(string dir)
        {
            //get file name
            string fileName = Path.GetFileName(dir);
            string desFile;
            //if not pass,move file to "failed" folder
            if (testResult == false)
            {                
                sw.WriteLine("Test Failed !");
                //if there is no sub folder
                if (dirName=="")
                   desFile = System.IO.Path.Combine(failedFolder, fileName);
                else
                   desFile = System.IO.Path.Combine(failedFolder+"\\"+dirName, fileName);
                failedTCN = failedTCN + 1;
            }
            else //if pass, move file to "passed" folder
            {               
                sw.WriteLine("Test Passed !");
                //if there is no sub folder
                if (dirName == "")
                   desFile = System.IO.Path.Combine(passedFolder, fileName);
                else
                   desFile = System.IO.Path.Combine(passedFolder + "\\" + dirName, fileName);
                passedTCN = passedTCN + 1;
            }
            System.IO.File.Move(dir, desFile);
            driver = null;        
        }

        //delete all items under one folder
        private void ClearFolder(string folderName)
        {
            DirectoryInfo dir = new DirectoryInfo(folderName);

            foreach (FileInfo fi in dir.GetFiles())
            {
                fi.Delete();
            }

            foreach (DirectoryInfo di in dir.GetDirectories())
            {
                ClearFolder(di.FullName);
                di.Delete();
            }
        }

        //delete empty folders under one directory
        private void DeleteEmptySubdirectories(string dirPath)
        {
            foreach (var directory in Directory.GetDirectories(dirPath))
            {
                DeleteEmptySubdirectories(directory);
                if (Directory.GetFiles(directory).Length == 0 &&
                    Directory.GetDirectories(directory).Length == 0)
                {
                    Directory.Delete(directory, false);
                }
            }
        }

        #region Additional test attributes

        // You can use the following additional attributes as you write your tests:

        ////Use TestInitialize to run code before running each test 
        //[TestInitialize()]
        //public void MyTestInitialize()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        ////Use TestCleanup to run code after each test has run
        //[TestCleanup()]
        //public void MyTestCleanup()
        //{        
        //    // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
        //}

        #endregion

        /// <summary>
        ///Gets or sets the test context which provides
        ///information about and functionality for the current test run.
        ///</summary>
        public TestContext TestContext
        {
            get
            {
                return testContextInstance;
            }
            set
            {
                testContextInstance = value;
            }
        }
        private TestContext testContextInstance;
    }
}
