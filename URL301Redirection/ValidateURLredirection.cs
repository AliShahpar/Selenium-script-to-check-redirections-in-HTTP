using System;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;


namespace URL301Redirection
{
    public class ValidateURLredirection
    {
        public string outputURL = "";
        public HttpStatusCode statuscode;
        public ExcelDataReadWrite objectExcel = new ExcelDataReadWrite();
        public int excel_rows = 0, excel_cols = 0;
        public IWebDriver driver;

        // this would validate URL's
        public void ValidateUrlsMethod()
        {
            // first call Excel rows and columns count method 
            ExcelRowsColsCount();
            
            #region
            // create a driver instance 
            var chromedriverhttptimeout = TimeSpan.FromMinutes(5);
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");
            options.AddArguments("--disable-extensions");
            //options.PageLoadStrategy = PageLoadStrategy.Normal;
            driver = new ChromeDriver("E:\\chromedriverWin32", options, chromedriverhttptimeout);
            #endregion

            for (int i = 1; i <= excel_rows; i++)
            {
                // Read from the data from Sheet1
                string url = objectExcel.ReadSheatData(2, i, 1).ToString();  // first parameter is for sheet number
                driver.Navigate().GoToUrl(url);

                #region
                try
                {
                    ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                    // Create a new HttpWebRequest Object to the mentioned URL.
                    HttpWebRequest myHttpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                    myHttpWebRequest.MaximumAutomaticRedirections = 1;
                    myHttpWebRequest.AllowAutoRedirect = true;
                    HttpWebResponse response = (HttpWebResponse)myHttpWebRequest.GetResponse();

                    outputURL = driver.Url.ToString();
                    statuscode = response.StatusCode;
                    // Write the date into Sheet-1
                    objectExcel.WriteSheetData(2, i, 3, outputURL);  // first parameter is for sheet number
                    objectExcel.WriteSheetData(2, i, 4, statuscode.ToString());  // first parameter is for sheet number
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    // Write the data into Sheet-1
                    objectExcel.WriteSheetData(2, i, 4, ex.ToString());  // first parameter is for sheet number
                }
                #endregion
                url = "";
            }
            // Quit the webdriver after the loop finishes
            driver.Quit();
            // close the excel sheet after the loop finishes
            objectExcel.CloseExcel();
        }

        // this will tell count of rows and columns in excel that contains some data and are not null 
        public void ExcelRowsColsCount()
        {
            // this will tell how many rows and columns contains data in sheet1
            int excelrownullvalue = 0, excelcolnullvalue = 0;
            int k = 1;

            #region
            // this will count the number of rows in the excel 
            for (int i = 1; i <= k; i++)
            {
                if (excelrownullvalue != 1)
                {
                    for (int j = 1; j <= (excel_rows + 1); j++)
                    {
                        try
                        {
                            if (objectExcel.ReadSheatData(1, j, i).ToString() != null)
                            {
                                excel_rows++;
                            }
                        }
                        catch
                        {
                            System.Console.WriteLine("Excel rows = " + excel_rows + "..... ");
                            excelrownullvalue = 1;
                        }
                    }
                }
            }
            #endregion

            #region
            // this will count the number of columns in the excel 
            for (int i = 1; i <= (excel_cols + 1); i++)
            {
                if (excelcolnullvalue != 1)
                {
                    for (int j = 1; j <= k; j++)
                    {
                        try
                        {
                            if (objectExcel.ReadSheatData(1, j, i).ToString() != null)
                            {
                                excel_cols++;
                            }
                        }
                        catch
                        {
                            System.Console.WriteLine("Excel columns = " + excel_cols + "..... ");
                            excelcolnullvalue = 1;
                        }
                    }
                }
            }
            #endregion
        }

    }
}
