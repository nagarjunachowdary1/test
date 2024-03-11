using Newtonsoft.Json;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Web;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using OpenQA.Selenium.Chrome;
namespace ScrapperConsole
{
    public class Scrapper
    {

        public static void DownloadExcel(string path)
        {
            try
            {
                //if (DateTime.Now.ToString("yyyy-MM-dd_HH") == (DateTime.Today.ToString("yyyy-MM-dd") + "_02"))
                //{
                //remove all files from archive folder
                string archfol = path + "\\Archive\\";
                foreach(string str in Directory.GetFiles(archfol))
                {
                    File.Delete(str);
                }
                //Move files to archive folder before running the job
                //string path = ConfigurationManager.AppSettings["downloadpath"].ToString();
                string[] files = Directory.GetFiles(path);
                    if (files?.Length > 0)
                    {
                        foreach (string str in files)
                        {
                            string archpath = path + "\\Archive\\" + str;
                            File.Move(str, Path.Combine(path, "Archive", Path.GetFileName(str)));
                        }
                    }
                    IWebDriver driver;
                var options = new EdgeDriver();
                //driver = new EdgeDriver();
                var edgeoptions = new EdgeOptions();

                    //var options = new EdgeOptions();
                   //driver = new EdgeDriver();
                  edgeoptions.AddUserProfilePreference("download.default_directory", path);
                driver = new EdgeDriver(edgeoptions); 
                //driver = new EdgeDriver("msedgedriver.exe", options);
                //driver = new EdgeDriver(options);
                //WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));
                System.Threading.Thread.Sleep(2000);
               // driver = new ChromeDriver("C:\\Users\\TALLAPANENIN\\source\\chromedriver.exe");
                System.Threading.Thread.Sleep(2000);
                    //'Fleetwave
                    //driver.Navigate().GoToUrl("https://ebianalytics.conedison.net/ui/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2F0.%20Public%20Shared%2FSuper%20Users%2FFinance%2FProject%20Accounting%20%26%20Billing%2FTEAM%20L2s");
                    driver.Navigate().GoToUrl(ConfigurationManager.AppSettings["appurl"]);

                    System.Threading.Thread.Sleep(5000);

                    //driver.SwitchTo().Frame("main");

                    //driver.FindElement(By.Id("username")).SendKeys("tallapanenin");
                    //driver.FindElement(By.Id("password")).SendKeys("Murthy.9059");
                    driver.FindElement(By.Id("idcs-signin-idp-signin-form-idp-button-AAD")).Click();

                System.Threading.Thread.Sleep(50000);
                while (driver.FindElement(By.LinkText("Export")) == null)
                {
                    System.Threading.Thread.Sleep(5000);
                }

                //wait.Until(ExpectedConditions.)
                //    var aa = driver.FindElement(By.LinkText("Export"));
                //wait.Until(ExpectedConditions(driver.FindElement(By.LinkText("Export"))).Click();

                //driver.FindElement(By.Id("idDownloadDataMenu")).Click();
                driver.FindElement(By.LinkText("Export")).Click();
                    driver.FindElement(By.LinkText("Data")).Click();
                    driver.FindElement(By.LinkText("Excel")).Click();
                    System.Threading.Thread.Sleep(25000);

                    //string path = ConfigurationManager.AppSettings["downloadpath"].ToString();
                    string[] dfiles = Directory.GetFiles(path);
                    if (dfiles?.Length > 0)
                    {
                        foreach (string str in dfiles)
                        {
                            File.Move(str, Path.Combine(path, DateTime.Now.ToString("yyyy-MMM-dd_HH-mm-ss") + "_" + Path.GetFileName(str)));
                        }
                    }

                    driver.Quit();
                    driver.Dispose();
                //}

            }
            catch (Exception ex)
            {
                File.AppendAllText(path + "\\error.txt", ex.Message);
            }
        }
        //public static  async void callback(object state)
        //{
        //    await DownloadExcel();
        //}
    }
}
