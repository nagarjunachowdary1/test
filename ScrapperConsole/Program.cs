using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium;
using static System.Net.Mime.MediaTypeNames;
using System.Configuration;
using ScrapperConsole;
using System.Timers;
using System.IO;


string path = ConfigurationManager.AppSettings["downloadpath"] ?? "".ToString();

try
{
    
    //Scrapper scrapper = new Scrapper();
    Scrapper.DownloadExcel(path);
    //var start = TimeSpan.Zero;
    //var period = TimeSpan.FromHours(1);
    //var timer = new System.Threading.Timer((e) =>
    //{
    //    //string path = ConfigurationManager.AppSettings["downloadpath"] ?? "".ToString();
    //    //Scrapper scrapper = new Scrapper();
    //    Scrapper.DownloadExcel();// path); 
    //}, null, start, period);
    //TimerCallback _callback = Scrapper.callback;
    //var _timer = new System.Threading.Timer(
    //    _callback,//Scrapper.callback,
    //    null,
    //    TimeSpan.Zero,
    //    period);
    //var t = new System.Timers.Timer(10 * 1000);

    //t.Elapsed += Scrapper.callback;
    //t.AutoReset = true;
    //t.Enabled = true;
}
catch (Exception ex)
{
    File.AppendAllText(path + "\\error.txt", ex.Message);
    File.AppendAllText(path + "\\error.txt", (ex.StackTrace ?? ""));
}
//try
//{
//    //Move files to archive folder before running the job
//    string path = ConfigurationManager.AppSettings["downloadpath"].ToString();
//    string[] files = Directory.GetFiles(path);
//    if (files?.Length > 0)
//    {
//        foreach (string str in files)
//        {
//            //string archpath = path + "\\Archive\\" + str;
//            File.Move(str, Path.Combine(path , "Archive", Path.GetFileName(str)));
//        }
//    }
//    IWebDriver driver;
//    var options = new EdgeOptions();
//    options.AddUserProfilePreference("download.default_directory", path);

//    driver = new EdgeDriver("msedgedriver.exe", options);
//    WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

//    //'driver = New ChromeDriver("C:\Users\chenc\AppData\Local\SeleniumBasic\chromedriver.exe")
//    System.Threading.Thread.Sleep(2000);
//    //'Fleetwave
//    //driver.Navigate().GoToUrl("https://ebianalytics.conedison.net/ui/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2F0.%20Public%20Shared%2FSuper%20Users%2FFinance%2FProject%20Accounting%20%26%20Billing%2FTEAM%20L2s");
//    driver.Navigate().GoToUrl(ConfigurationManager.AppSettings["appurl"]);

//    System.Threading.Thread.Sleep(5000);

//    //driver.SwitchTo().Frame("main");

//    //driver.FindElement(By.Id("username")).SendKeys("tallapanenin");
//    //driver.FindElement(By.Id("password")).SendKeys("Murthy@9059");
//    driver.FindElement(By.Id("idcs-signin-idp-signin-form-idp-button-AAD")).Click();

//    System.Threading.Thread.Sleep(50000);
//    //while(driver.FindElement(By.LinkText("Export")) == null)
//    //{
//    //    System.Threading.Thread.Sleep(5000);
//    //}

//    //wait.Until(ExpectedConditions.)
//    //var aa = driver.FindElement(By.LinkText("Export"));
//    //wait.Until(ExpectedConditions( driver.FindElement(By.LinkText("Export"))).Click();

//    //driver.FindElement(By.Id("idDownloadDataMenu")).Click();
//    driver.FindElement(By.LinkText("Export")).Click();
//    driver.FindElement(By.LinkText("Data")).Click();
//    driver.FindElement(By.LinkText("Excel")).Click();
//    System.Threading.Thread.Sleep(25000);

//    //string path = ConfigurationManager.AppSettings["downloadpath"].ToString();
//    string[] dfiles = Directory.GetFiles(path);
//    if (dfiles?.Length > 0)
//    {
//        foreach (string str in dfiles)
//        {
//            File.Move(str, Path.Combine(path, DateTime.Now.ToString("yyyy-MMM-dd_HH-mm-ss") +"_"+ Path.GetFileName(str)));
//        }
//    }

//    driver.Quit();
//    driver.Dispose();
//}
//catch (Exception ex)
//{
//    File.AppendAllText("error.txt", ex.Message);
//}

