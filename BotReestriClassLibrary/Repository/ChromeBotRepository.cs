using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using BotReestriClassLibrary.DTOs;
using BotReestriClassLibrary.Interface;
using BotReestriClassLibrary.Wrapper;
using System.ComponentModel;
using System.Data;
using System.Net.Http;
using System.Drawing;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.IO;
using System.Diagnostics;
using System.Net;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;

namespace BotReestriClassLibrary.Repository
{
    public class ChromeBotRepository : IChromeBot

    {
        
        /*test aleks*/
        public async Task<bool> CheckForInternetConnectionAsync()
        {
            using (var client = new HttpClient())
            {
                try
                {
                    using (var response = await client.GetAsync("http://clients3.google.com/generate_204"))
                    {
                        return response.IsSuccessStatusCode;
                    }
                }
                catch
                {
                    return false;
                }
            }
        }
        public Result<ChromeBotDTO> BotChromeArguments(string ExcelPath, string Destination)
        {

            if (!String.IsNullOrEmpty(ExcelPath) & !String.IsNullOrEmpty(Destination))
            {
                try
                {
                    // აფდეითდება დრაივერი ქრომის და ავტომატურად იღებს googleChrome-ს ბოლო ვერსიას 
                    new DriverManager().SetUpDriver(new ChromeConfig());

                    using (IWebDriver driver = new ChromeDriver())
                    {

                        //Navigate to google page
                        driver.Navigate().GoToUrl("https://www.my.gov.ge/ka-ge/services/5/service/176");

                        //Maximize the window
                        driver.Manage().Window.Maximize();


                        // ინტერნეტის ჩეკისთვის 
                        Thread.Sleep(10000);

                        //var wait1 = new WebDriverWait(driver, new TimeSpan(0, 0, 30));
                        //wait1.Until(c => c.FindElement(By.ClassName("iframe")));
                        while (true)
                        {
                            if (CheckForInternetConnectionAsync().GetAwaiter().GetResult())
                            {
                                System.Threading.Thread.Sleep(10000);
                                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(10);

                                IWebElement iframe = driver.FindElement(By.ClassName("iframe"));
                                driver.SwitchTo().Frame(iframe);


                                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);

                                string fullPath = "";

                                fullPath = Destination;
                                // System.IO.Directory.CreateDirectory(fullPath);

                                var dateTimeNow = DateTime.Now;

                                //   fullPath = fullPath.Replace("\\\\", "\\");
                                if (File.Exists(fullPath + "\\" + "statemantesText.xlsx"))
                                {
                                    fullPath = fullPath + "\\" + "statemantesText" + "(2)" + ".xlsx";
                                    xlWorkbook.SaveCopyAs(fullPath);
                                }
                                else
                                {
                                    fullPath = fullPath + "\\" + "statemantesText.xlsx";
                                    xlWorkbook.SaveCopyAs(fullPath);
                                }
                                xlWorkbook.Close(false, Missing.Value, false);
                                xlWorkbook = xlApp.Workbooks.Open(fullPath);
                                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                                Microsoft.Office.Interop.Excel.Range oRng = xlWorksheet.Range["J1"];
                                oRng.NumberFormat = "0";
                                oRng.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight,
                                        Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                oRng = xlWorksheet.Range["J1"];
                                oRng.Value2 = "შენიშვნათარიღი";

                                Microsoft.Office.Interop.Excel.Range oRng2 = xlWorksheet.Range["K1"];
                                oRng2.NumberFormat = "0";
                                oRng2.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight,
                                        Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                oRng2 = xlWorksheet.Range["K1"];
                                oRng2.Value2 = "შენიშვნახალი";

                                Microsoft.Office.Interop.Excel.Range oRng3 = xlWorksheet.Range["L1"];
                                oRng3.NumberFormat = "0";
                                oRng3.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight,
                                        Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                oRng3 = xlWorksheet.Range["L1"];
                                oRng3.Value2 = "სტატუსი";

                                Microsoft.Office.Interop.Excel.Range oRng4 = xlWorksheet.Range["M1"];
                                oRng4.NumberFormat = "0";
                                oRng4.EntireColumn.Insert(Microsoft.Office.Interop.Excel.XlInsertShiftDirection.xlShiftToRight,
                                        Microsoft.Office.Interop.Excel.XlInsertFormatOrigin.xlFormatFromRightOrBelow);
                                oRng4 = xlWorksheet.Range["M1"];
                                oRng4.Value2 = "დაინტერესებული პირი";


                                int rowCount = xlRange.Rows.Count;
                                int colCount = xlRange.Columns.Count;

                                // dt.Column = colCount;  
                                //  dataGridView1.ColumnCount = colCount;
                                // dataGridView1.RowCount = rowCount;
                                Random random = new Random();
                                int randomNumber;


                                //Microsoft.Office.Interop.Excel._Worksheet wks;
                                //templatePath = System.Windows.Forms.Application.StartupPath + @"/shabloni_StatemanetText.xlsx";
                                //Microsoft.Office.Interop.Excel._Workbook wkb = xlApp.Workbooks.Open(templatePath);

                                //Range wksRange;
                                //wks = (Microsoft.Office.Interop.Excel._Worksheet)wkb.Worksheets.get_Item(1);
                                //wks.Activate();
                                //wks.Visible = XlSheetVisibility.xlSheetVisible;
                                for (int i = 2; i <= rowCount; i++)
                                {
                                    //bool connectionStatus = CheckForInternetConnection();
                                    //ConnectionOut outcon = new ConnectionOut();
                                    //while (connectionStatus == false)
                                    //{
                                    //    connectionStatus = CheckForInternetConnection();

                                    //    if (outcon.Visible == false)
                                    //    {
                                    //        outcon.Show();
                                    //        outcon.Activate();
                                    //    }
                                    //}
                                    //if (outcon.Visible == true)
                                    //{
                                    //    outcon.Hide();
                                    //}
                                    var vallll = xlRange.Cells[i, 8].Value2;
                                    var vallll3 = xlRange.Cells[i, 7].Value2;
                                    var vallll4 = xlRange.Cells[i, 9].Value2;



                                    if (xlRange.Cells[i, 8] != null && xlRange.Cells[i, 8].Value2 != null)
                                    {
                                        string value = xlRange.Cells[i, 8].Value2.ToString();

                                        driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);
                                        IWebElement element = driver.FindElement(By.XPath("//*[@id='input_0']"));

                                        element.Click();
                                        element.Clear();

                                        if (!String.IsNullOrEmpty(xlRange.Cells[i, 8].Value2))
                                        {
                                            element.SendKeys(xlRange.Cells[i, 8].Value2);

                                            randomNumber = random.Next(30, 50);
                                            //System.Threading.Thread.Sleep(randomNumber);

                                            IWebElement searchButton = driver.FindElement(By.XPath("//*[@id='searchForm']/div[2]/button[1]"));
                                            searchButton.Click();


                                            randomNumber = random.Next(7000, 15000);
                                            System.Threading.Thread.Sleep(randomNumber);
                                            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(20);

                                            string statemantTexti = driver.FindElement(By.XPath("//*[@id='callCenter']/md-list/md-list-item/div/div[1]/div/p[2]/span[1]")).Text;

                                            string statusiTexti = driver.FindElement(By.XPath("//*[@id='callCenter']/md-list/md-list-item/div/div[1]/div/p[2]/span[2]")).Text;
                                            string dainteresebuliPiriTexti = driver.FindElement(By.XPath("//*[@id='callCenter']/md-list/md-list-item/div/div[1]/div/p[2]/span[3]")).Text;




                                            //გასატესტად ფორმის რომელიც არ ჩანს
                                            // bool isElementDisplayed = driver.FindElement(By.XPath("//*[@id='callCenter']/md-list/md-list-item/div/div[1]/div/p[2]/span[6]")).Displayed;


                                            //*[@id="callCenter"]/md-list/md-list-item/div/div[1]/div/p[2]/span[3]
                                            //*[@id="callCenter"]/md-list/md-list-item/div/div[1]/div/p[2]/span[4]

                                            //*[@id="callCenter"]/md-list/md-list-item/div/div[1]/div/p[2]/span[3]

                                            //*[@id="callCenter"]/md-list/md-list-item/div/div[1]/div/p[2]/span[6]

                                            //Microsoft.Office.Interop.Excel.Range oRng = oSheet.Range["I1"];
                                            //     string[] tokens = statemantTexti.Split(" - ");
                                            string[] tokens = statemantTexti.Split(new string[] { " - " }, StringSplitOptions.None);
                                            //  wks.Cells[i][1] = xlRange.Cells[i, 8].Value2;
                                            xlWorksheet.Cells[10][i] = tokens[0];
                                            xlWorksheet.Cells[11][i] = tokens[1];
                                            xlWorksheet.Cells[12][i] = statusiTexti;
                                            xlWorksheet.Cells[13][i] = dainteresebuliPiriTexti;
                                            randomNumber = random.Next(5000, 10000);
                                        }

                                    }
                                    else if (xlRange.Cells[i, 8].Value2 == null || xlRange.Cells[i, 7].Value2 == null || xlRange.Cells[i, 9].Value2 == null)
                                    {
                                        continue; // Skip processing if value is empty
                                    }
                                    xlApp.DisplayAlerts = false;
                                    xlWorkbook.SaveAs(fullPath);
                                }

                                //driver.SwitchTo().DefaultContent();
                                //Close the browser



                                xlWorkbook.Close(false, Missing.Value, false);
                                driver.Close();
                                driver.Quit();
                                //    MessageBox.Show(
                                //"წარმატებით დასრულდა",
                                //"შეტყობინება",
                                //MessageBoxButtons.OK,
                                // MessageBoxIcon.Information,
                                //MessageBoxDefaultButton.Button1,
                                //(MessageBoxOptions)0x40000); // this set TopMost
                                //    OpenFolder(ToTextBox.Text);
                                //    //  MessageBox.Show("წარმატებით დასრულდა");
                                //    this.Close();
                                break;
                            }
                            else
                            {
                                // იცდის 5 წამს შემდეგი შემოწმებისთვის მოვიდა თუ არა ინტერნეტი 
                                Thread.Sleep(5000);
                            }
                        }
                    }
                }

                catch (Exception ex)
                {

                    return new Result<ChromeBotDTO>
                    {
                        Success = false,
                        Message = "წარმატებით განხორციელდა ბოტის მუშაობა !" + ex.Message,
                        StatusCode = System.Net.HttpStatusCode.BadGateway
                    };
                }
            }

            else
            {
                return new Result<ChromeBotDTO>
                {
                    Success = false,
                    Message = "ველები ცარიელია\"",
                    StatusCode = System.Net.HttpStatusCode.BadGateway
                };
            }























            return new Result<ChromeBotDTO>
            {
                Success = true,
                Message = "წარმატებით განხორციელდა ბოტის მუშაობა !",
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }
    }
}
