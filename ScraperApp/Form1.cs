using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Support.UI;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace ScraperApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                IWebDriver driver;
                driver = new EdgeDriver("msedgedriver.exe");
                WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

                //'driver = New ChromeDriver("C:\Users\chenc\AppData\Local\SeleniumBasic\chromedriver.exe")
                System.Threading.Thread.Sleep(5000);
                //'Fleetwave
                driver.Navigate().GoToUrl("https://ebianalytics.conedison.net/ui/analytics/saw.dll?PortalGo&Action=prompt&path=%2Fshared%2F0.%20Public%20Shared%2FSuper%20Users%2FFinance%2FProject%20Accounting%20%26%20Billing%2FTEAM%20L2s");

                System.Threading.Thread.Sleep(5000);

                //driver.SwitchTo().Frame("main");

                //driver.FindElement(By.Id("username")).SendKeys("tallapanenin");
                //driver.FindElement(By.Id("password")).SendKeys("Murthy@9059");
                driver.FindElement(By.Id("idcs-signin-idp-signin-form-idp-button-AAD")).Click();

                System.Threading.Thread.Sleep(25000);
                //while(driver.FindElement(By.LinkText("Export")) == null)
                //{
                //    System.Threading.Thread.Sleep(5000);
                //}

                //wait.Until(ExpectedConditions.)
                //var aa = driver.FindElement(By.LinkText("Export"));
                //wait.Until(ExpectedConditions( driver.FindElement(By.LinkText("Export"))).Click();

                //driver.FindElement(By.Id("idDownloadDataMenu")).Click();
                driver.FindElement(By.LinkText("Export")).Click();
                driver.FindElement(By.LinkText("Data")).Click();
                driver.FindElement(By.LinkText("Excel")).Click();
                System.Threading.Thread.Sleep(25000);

                driver.Quit();
                driver.Dispose();
            }
            catch(Exception ex)
            {
                File.AppendAllText("~/error.txt", ex.Message);
            }
            Application.Exit();

        }

        private void Form1_QueryAccessibilityHelp(object sender, QueryAccessibilityHelpEventArgs e)
        {


        }

    }
}
