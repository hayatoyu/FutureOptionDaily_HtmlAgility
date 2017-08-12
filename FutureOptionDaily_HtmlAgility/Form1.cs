using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using HtmlAgilityPack;
using FutureOptionDaily_HtmlAgility.Class;
using WatiN.Core;

namespace FutureOptionDaily_HtmlAgility
{
    
    public partial class Form1 : System.Windows.Forms.Form
    {
        public List<DateTime> duration;
        

        public Form1()
        {
            InitializeComponent();
            //threads = new List<Thread>();
        }

        

        private void btnSelectFilePath_Click(object sender, EventArgs e)
        {
            OpenFileDialog selectFilePath = new OpenFileDialog();
            string filepath = string.Empty;
            selectFilePath.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            selectFilePath.RestoreDirectory = true;
            if (selectFilePath.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                filepath = selectFilePath.FileName.ToString();
                txtFilePath.Text = filepath;

            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            txtFilePath.Text = txtFilePath.Text.Replace("\"", string.Empty);
            BrowserSession browser = new BrowserSession();
            if(File.Exists(txtFilePath.Text))
            {
                Microsoft.Office.Interop.Excel.Application xlsApp = null;
                Workbook wb = null;
                Worksheet ws_Template = null;
                Worksheet ws_New = null;
                IE ie = new IE("http://www.twse.com.tw/zh/page/trading/exchange/FMTQIK.html");
                IE ieBig3 = new IE("http://www.twse.com.tw/zh/page/trading/fund/BFI82U.html");
                IE ieMargin = new IE("http://www.twse.com.tw/zh/page/trading/exchange/MI_MARGN.html");

                if (!ckBoxDuration.Checked)
                {
                    try
                    {
                        xlsApp = new Microsoft.Office.Interop.Excel.Application();
                        xlsApp.DisplayAlerts = false;
                        xlsApp.AskToUpdateLinks = false;
                        wb = xlsApp.Workbooks.Open(txtFilePath.Text);
                        ws_Template = wb.Worksheets[1];
                        ws_Template.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                        ws_New = wb.Worksheets[wb.Worksheets.Count];
                        bool isNoData = false;

                        // 三大法人期貨
                        lblStatus.Text = "正在抓取本日三大法人期貨盤後資料...";
                        isNoData = FutureBig3(new BrowserSession(), dateTimePicker1.Value, ws_New, false);
                        if (!isNoData)
                        {

                            lblStatus.Text = "正在抓取上一交易日三大法人期貨盤後資料...";
                            FutureBig3(new BrowserSession(), dateTimePicker2.Value, ws_New, true);


                            // 期貨漲跌點數
                            lblStatus.Text = "正在抓取今日期貨漲跌點數...";
                            getFuturePoint(new BrowserSession(), dateTimePicker1.Value, ws_New);

                            // 三大選擇權
                            lblStatus.Text = "正在抓取今日三大法人選擇權盤後資料...";
                            getOptionBig3(new BrowserSession(), dateTimePicker1.Value, ws_New, false);
                            lblStatus.Text = "正在抓取上一交易日三大法人選擇權盤後資料...";
                            getOptionBig3(new BrowserSession(), dateTimePicker2.Value, ws_New, true);

                            // 五大十大期貨
                            lblStatus.Text = "正在抓取五大十大今日期貨盤後資料...";
                            FutureFive_Ten(new BrowserSession(), dateTimePicker1.Value, ws_New, false);
                            lblStatus.Text = "正在抓取上一交易日五大十大期貨盤後資料...";
                            FutureFive_Ten(new BrowserSession(), dateTimePicker2.Value, ws_New, true);

                            // 五大十大選擇權
                            lblStatus.Text = "正在抓取五大十大今日選擇權盤後資料...";
                            OptionFive_Ten(new BrowserSession(), dateTimePicker1.Value, ws_New, false);
                            lblStatus.Text = "正在抓取五大十大上一交易日選擇權盤後資料...";
                            OptionFive_Ten(new BrowserSession(), dateTimePicker2.Value, ws_New, true);

                            // 三大現貨買賣超
                            lblStatus.Text = "正在抓取今日三大法人現貨買賣超...";
                            StockBig3(ieBig3, dateTimePicker1.Value, ws_New);

                            // 現貨成交資訊
                            lblStatus.Text = "正在抓取現貨成交資訊...";
                            TradingInfo(ie,dateTimePicker1.Value, ws_New);


                            // 前一日信用交易
                            lblStatus.Text = "正在抓取前一日信用交易資料...";
                            MarginTrading(ieMargin, dateTimePicker2.Value, ws_New);

                            // 改日期
                            DateTimeModify(ws_New, dateTimePicker1.Value);
                        }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    finally
                    {
                        //WaitUntilThreadsEnd();
                        ie.Dispose();
                        ieBig3.Dispose();
                        ieMargin.Dispose();
                        wb.Close(true);
                        if (ws_New != null)
                            Marshal.FinalReleaseComObject(ws_New);
                        if (ws_Template != null)
                            Marshal.FinalReleaseComObject(ws_Template);
                        if (wb != null)
                            Marshal.FinalReleaseComObject(wb);
                        if (xlsApp != null)
                        {
                            xlsApp.DisplayAlerts = true;
                            xlsApp.AskToUpdateLinks = true;
                            Marshal.FinalReleaseComObject(xlsApp);
                        }
                        
                    }
                }

                else
                {
                    try
                    {
                        xlsApp = new Microsoft.Office.Interop.Excel.Application();
                        xlsApp.DisplayAlerts = false;
                        xlsApp.AskToUpdateLinks = false;
                        wb = xlsApp.Workbooks.Open(txtFilePath.Text);
                        ws_Template = wb.Worksheets[1];

                        int index = 1;
                        bool isNoData = false;

                        // 如果碰到沒有交易的日期，自動跳過。
                        while(index < duration.Count)
                        {
                            ws_Template.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                            ws_New = wb.Worksheets[wb.Worksheets.Count];

                            // 三大法人期貨
                            lblStatus.Text = "正在抓取" + duration[index].ToShortDateString() + "三大法人期貨盤後資料...";
                            isNoData = FutureBig3(new BrowserSession(), duration[index], ws_New, false);
                            if (isNoData)
                            {
                                duration.RemoveAt(index);
                                ws_New.Delete();
                            }
                            else
                            {
                                lblStatus.Text = "正在抓取" + duration[index - 1].ToShortDateString() + "三大法人期貨盤後資料...";
                                FutureBig3(new BrowserSession(), duration[index - 1], ws_New, true);

                                // 期貨漲跌點數
                                lblStatus.Text = "正在抓取" + duration[index].ToShortDateString() + "期貨漲跌點數...";
                                getFuturePoint(new BrowserSession(), duration[index], ws_New);

                                // 三大選擇權
                                lblStatus.Text = "正在抓取" + duration[index].ToShortDateString() + "三大法人選擇權盤後資料...";
                                getOptionBig3(new BrowserSession(), duration[index], ws_New, false);
                                lblStatus.Text = "正在抓取" + duration[index - 1].ToShortDateString() + "三大法人選擇權盤後資料...";
                                getOptionBig3(new BrowserSession(), duration[index - 1], ws_New, true);

                                // 五大十大期貨
                                lblStatus.Text = "正在抓取五大十大" + duration[index].ToShortDateString() + "期貨盤後資料...";
                                FutureFive_Ten(new BrowserSession(), duration[index], ws_New, false);
                                lblStatus.Text = "正在抓取五大十大" + duration[index - 1].ToShortDateString() + "期貨盤後資料...";
                                FutureFive_Ten(new BrowserSession(), duration[index - 1], ws_New, true);

                                // 五大十大選擇權
                                lblStatus.Text = "正在抓取五大十大" + duration[index].ToShortDateString() + "選擇權盤後資料...";
                                OptionFive_Ten(new BrowserSession(), duration[index], ws_New, false);
                                lblStatus.Text = "正在抓取五大十大" + duration[index - 1].ToShortDateString() + "選擇權盤後資料...";
                                OptionFive_Ten(new BrowserSession(), duration[index - 1], ws_New, true);

                                // 三大現貨買賣超
                                lblStatus.Text = "正在抓取" + duration[index].ToShortDateString() + "三大法人現貨買賣超...";
                                StockBig3(ieBig3, duration[index], ws_New);

                                // 現貨成交資訊
                                lblStatus.Text = "正在抓取" + duration[index].ToShortDateString() + "現貨成交資訊...";
                                TradingInfo(ie,duration[index], ws_New);


                                // 前一日信用交易
                                lblStatus.Text = "正在抓取" + duration[index - 1].ToShortDateString() + "信用交易資料...";
                                MarginTrading(ieMargin, duration[index - 1], ws_New);

                                // 改日期
                                DateTimeModify(ws_New, duration[index]);

                                index++;
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                        MessageBox.Show(ex.StackTrace);
                    }
                    finally
                    {
                        //WaitUntilThreadsEnd();
                        ie.Dispose();
                        ieBig3.Dispose();
                        ieMargin.Dispose();
                        wb.Close(true);
                        if (ws_New != null)
                            Marshal.FinalReleaseComObject(ws_New);
                        if (ws_Template != null)
                            Marshal.FinalReleaseComObject(ws_Template);
                        if (wb != null)
                            Marshal.FinalReleaseComObject(wb);
                        if (xlsApp != null)
                        {
                            xlsApp.DisplayAlerts = true;
                            xlsApp.AskToUpdateLinks = true;
                            Marshal.FinalReleaseComObject(xlsApp);
                        }
                    }
                }
                MessageBox.Show("完成！");
                this.Close();
            }
            else
            {
                MessageBox.Show("Excel檔案不存在，請確認。");
            }
        }


        private bool FutureBig3(BrowserSession browser,DateTime date,Worksheet ws_New,bool isYesterday)
        {
            browser.Get("http://www.taifex.com.tw/chinese/3/7_12_3.asp");
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_Y")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_M")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_D")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("syear")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("smonth")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("sday")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("datestart")).First().SetAttributeValue("value", date.ToShortDateString());
            string response = browser.Post("http://www.taifex.com.tw/chinese/3/7_12_3.asp");

            // 判斷有沒有資料
            if (response.Contains("查無資料"))
                return true;

            // 先抓大台

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("table_f")).First();

            // 自營商
            HtmlNode cell = table.SelectSingleNode("tbody[1]/tr[4]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[13,3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[13,6].Value2 = cell.InnerText.Replace(",", string.Empty);

            cell = table.SelectSingleNode("tbody[1]/tr[4]/td[12]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[13,4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[13,7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 投信
            cell = table.SelectSingleNode("tbody[1]/tr[5]/td[8]/div[1]/font[1]");
            if(isYesterday)
                ws_New.Cells[14, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[14, 6].Value2 = cell.InnerText.Replace(",", string.Empty);

            cell = table.SelectSingleNode("tbody[1]/tr[5]/td[10]/div[1]/font[1]");
            if(isYesterday)
                ws_New.Cells[14, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[14, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            

            // 外資
            cell = table.SelectSingleNode("tbody[1]/tr[6]/td[8]/div[1]/font[1]");
            if(isYesterday)
                ws_New.Cells[15, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[15, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            
            cell = table.SelectSingleNode("tbody[1]/tr[6]/td[10]/div[1]/font[1]");
            if(isYesterday)
                ws_New.Cells[15, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[15, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 金融期

            // 自營
            cell = table.SelectSingleNode("tbody[1]/tr[10]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[29, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[29, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[10]/td[12]/div[1]/font[1]");
            if(isYesterday)
                ws_New.Cells[29, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[29, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 投信
            cell = table.SelectSingleNode("tbody[1]/tr[11]/td[8]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[30, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[30, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[11]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[30, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[30, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 外資
            cell = table.SelectSingleNode("tbody[1]/tr[12]/td[8]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[31, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[31, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[12]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[31, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[31, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 小台

            // 自營商
            cell = table.SelectSingleNode("tbody[1]/tr[13]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[23, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[23, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[13]/td[12]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[23, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[23, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 投信
            cell = table.SelectSingleNode("tbody[1]/tr[14]/td[8]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[24, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[24, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[14]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[24, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[24, 7].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 外資
            cell = table.SelectSingleNode("tbody[1]/tr[15]/td[8]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[25, 3].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[25, 6].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[15]/td[10]/div[1]/font[1]");
            if (isYesterday)
                ws_New.Cells[25, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[25, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            return false;
            
        }

        

        private void getFuturePoint(BrowserSession browser,DateTime date,Worksheet ws_New)
        {
            browser.Get("http://www.taifex.com.tw/chinese/3/3_1_1.asp");
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("commodity_id")).First().SetAttributeValue("value", "TX");
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_Y")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_M")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_D")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("syear")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("smonth")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("sday")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("datestart")).First().SetAttributeValue("value", date.ToShortDateString());
            string response = browser.Post("http://www.taifex.com.tw/chinese/3/3_1_1.asp");

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("table_f")).First();
            HtmlNode cell = table.SelectSingleNode("tbody[1]/tr[2]/td[7]/font[1]");
            ws_New.Cells[5, 11].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[2]/td[6]");
            ws_New.Cells[7, 12].Value2 = cell.InnerText;
            if(ws_New.Cells[5,11].Text.Contains("-"))
                ws_New.Cells[5,11].Font.Color = -11489280;
            else
                ws_New.Cells[5,11].Font.Color = -16776961;
        }

        private void getOptionBig3(BrowserSession browser,DateTime date,Worksheet ws_New,bool isYesterday)
        {
            browser.Get("http://www.taifex.com.tw/chinese/3/7_12_5.asp");
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_Y")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_M")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("DATA_DATE_D")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("syear")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("smonth")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("sday")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("datestart")).First().SetAttributeValue("value", date.ToShortDateString());
            string response = browser.Post("http://www.taifex.com.tw/chinese/3/7_12_5.asp");

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("table_f")).First();
            
            // 自營買權
            HtmlNode cell = table.SelectSingleNode("tbody[1]/tr[4]/td[11]/font[1]");
            if (isYesterday)
                ws_New.Cells[35, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[35, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[4]/td[13]/font[1]");
            if (isYesterday)
                ws_New.Cells[35, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[35, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 投信買權
            cell = table.SelectSingleNode("tbody[1]/tr[5]/td[8]/font[1]");
            if (isYesterday)
                ws_New.Cells[36, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[36, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[5]/td[10]/font[1]");
            if (isYesterday)
                ws_New.Cells[36, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[36, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 外資買權
            cell = table.SelectSingleNode("tbody[1]/tr[6]/td[8]/font[1]");
            if (isYesterday)
                ws_New.Cells[37, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[37, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[6]/td[10]/font[1]");
            if (isYesterday)
                ws_New.Cells[37, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[37, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 自營賣權
            cell = table.SelectSingleNode("tbody[1]/tr[7]/td[9]/font[1]");
            if (isYesterday)
                ws_New.Cells[42, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[42, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[7]/td[11]/font[1]");
            if(isYesterday)
                ws_New.Cells[42, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[42, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 投信賣權
            cell = table.SelectSingleNode("tbody[1]/tr[8]/td[8]/font[1]");
            if (isYesterday)
                ws_New.Cells[43, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[43, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[8]/td[10]/font[1]");
            if (isYesterday)
                ws_New.Cells[43, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[43, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

            // 外資賣權
            cell = table.SelectSingleNode("tbody[1]/tr[9]/td[8]/font[1]");
            if (isYesterday)
                ws_New.Cells[44, 4].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[44, 7].Value2 = cell.InnerText.Replace(",", string.Empty);
            cell = table.SelectSingleNode("tbody[1]/tr[9]/td[10]/font[1]");
            if (isYesterday)
                ws_New.Cells[44, 5].Value2 = cell.InnerText.Replace(",", string.Empty);
            else
                ws_New.Cells[44, 8].Value2 = cell.InnerText.Replace(",", string.Empty);

        }

        private void FutureFive_Ten(BrowserSession browser,DateTime date,Worksheet ws_New,bool isYesterday)
        {
            browser.Get("http://www.taifex.com.tw/chinese/3/7_8.asp");
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("yytemp")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("mmtemp")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("ddtemp")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_yy")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_mm")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_dd")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("chooseitemtemp")).First().SetAttributeValue("value", "TX     ");
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("datestart")).First().SetAttributeValue("value", date.ToShortDateString());
            string response = browser.Post("http://www.taifex.com.tw/chinese/3/7_8.asp");

            string Date;
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("table_f")).First();
            
            // 五大期貨
            HtmlNode cell = table.SelectSingleNode("tr[5]/td[1]/div[1]");
            Date = cell.InnerText.Replace("\n", string.Empty);
            cell = table.SelectSingleNode("tr[5]/td[2]/div[1]");
            if (isYesterday)
                ws_New.Cells[16, 3].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[16, 6].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[5]/td[6]/div[1]");
            if (isYesterday)
                ws_New.Cells[16, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[16, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[2]/div[1]");
            if (isYesterday)
                ws_New.Cells[18, 3].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[18, 6].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[6]/div[1]");
            if (isYesterday)
                ws_New.Cells[18, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[18, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            // 十大期貨
            cell = table.SelectSingleNode("tr[5]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[17, 3].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[17, 6].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[5]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[17, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[17, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[19, 3].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[19, 6].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[19, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[19, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            if (!isYesterday)
            {
                ws_New.Cells[16, 2].Value2 = ws_New.Cells[16, 2].Value2.ToString().Replace("yyyyMM", Date);
                ws_New.Cells[17, 2].Value2 = ws_New.Cells[17, 2].Value2.ToString().Replace("yyyyMM", Date);
            }

        }

        private void OptionFive_Ten(BrowserSession browser,DateTime date,Worksheet ws_New,bool isYesterday)
        {
            browser.Get("http://www.taifex.com.tw/chinese/3/7_9.asp");
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("yytemp")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("mmtemp")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("ddtemp")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_yy")).First().SetAttributeValue("value", date.Year.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_mm")).First().SetAttributeValue("value", date.Month.ToString());
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("choose_dd")).First().SetAttributeValue("value", date.ToString("dd"));
            browser.FormElements.Where(x => x.GetAttributeValue("name", "").Equals("chooseitemtemp")).First().SetAttributeValue("value", "TXO    ");
            browser.FormElements.Where(x => x.GetAttributeValue("id", "").Equals("datestart")).First().SetAttributeValue("value", date.ToShortDateString());
            string response = browser.Post("http://www.taifex.com.tw/chinese/3/7_9.asp");

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("table_f")).First();
            string Date;
            // 五大買權
            HtmlNode cell = table.SelectSingleNode("tr[5]/td[1]/div[1]");
            Date = cell.InnerText.Replace("\n", string.Empty);
            cell = table.SelectSingleNode("tr[5]/td[2]/div[1]");
            if(isYesterday)
                ws_New.Cells[38,4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[38, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[5]/td[6]/div[1]");
            if(isYesterday)
                ws_New.Cells[38, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[38, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[2]/div[1]");
            if(isYesterday)
                ws_New.Cells[40, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[40, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[6]/div[1]");
            if(isYesterday)
                ws_New.Cells[40, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[40, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            // 十大買權
            cell = table.SelectSingleNode("tr[5]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[39, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[39, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[5]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[39,5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[39, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[41, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[41, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[6]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[41, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[41, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            // 五大賣權
            cell = table.SelectSingleNode("tr[8]/td[2]/div[1]");
            if(isYesterday)
                ws_New.Cells[45, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[45, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[8]/td[6]/div[1]");
            if(isYesterday)
                ws_New.Cells[45, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[45, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[9]/td[2]/div[1]");
            if(isYesterday)
                ws_New.Cells[47, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[47, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[9]/td[6]/div[1]");
            if(isYesterday)
                ws_New.Cells[47, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[47, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            // 十大賣權
            cell = table.SelectSingleNode("tr[8]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[46, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[46, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[8]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[46, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[46, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[9]/td[4]/div[1]");
            if(isYesterday)
                ws_New.Cells[48, 4].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[48, 7].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            cell = table.SelectSingleNode("tr[9]/td[8]/div[1]");
            if(isYesterday)
                ws_New.Cells[48, 5].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();
            else
                ws_New.Cells[48, 8].Value2 = cell.InnerText.Substring(0, cell.InnerText.IndexOf("(")).Replace(",", string.Empty).Trim();

            if (!isYesterday)
            {
                // 買權五大十大
                ws_New.Cells[38, 3].Value2 = ws_New.Cells[38, 3].Value2.ToString().Replace("yyyyMM", Date);
                ws_New.Cells[39, 3].Value2 = ws_New.Cells[39, 3].Value2.ToString().Replace("yyyyMM", Date);

                // 賣權五大十大
                ws_New.Cells[45, 3].Value2 = ws_New.Cells[45, 3].Value2.ToString().Replace("yyyyMM", Date);
                ws_New.Cells[46, 3].Value2 = ws_New.Cells[46, 3].Value2.ToString().Replace("yyyyMM", Date);
            }
        }

        #region oldBig3
        /*
        private void StockBig3(BrowserSession browser,DateTime date,Worksheet ws_New)
        {
            //string response = browser.Get("http://www.twse.com.tw/ch/trading/fund/BFI82U/BFI82U.php?input_date=" + ROCDate(date));
            //string response = browser.Get(string.Format("http://www.twse.com.tw/ch/trading/fund/BFI82U/BFI82U.php?input_date={0}",ROCDate(date)));
            string response = browser.Get(string.Format("http://www.twse.com.tw/ch/trading/fund/BFI82U/BFI82U_print.php?begin_date={0}&end_date={1}&report_type=day&language=ch", date.ToString("yyyyMMdd"), date.ToString("yyyyMMdd")));

            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            
            doc.LoadHtml(response);
            HtmlNode table = doc.DocumentNode.Descendants("table").Where(t => t.Attributes.Contains("class") && t.Attributes["class"].Value.Equals("board_trad")).First();
            HtmlNode cell;
            for (int i = 3; i < 7; i++)
            {
                cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[1]/div[1]");

                // 自營商(避險)
                if (cell.InnerText.Contains("避險"))
                {
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[2]");
                    ws_New.Cells[6, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[3]");
                    ws_New.Cells[6, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                }
                else if (cell.InnerText.Contains("自營"))
                {
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[2]");
                    ws_New.Cells[5, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[3]");
                    ws_New.Cells[5, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                }
                else if (cell.InnerText.Contains("投信"))
                {
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[2]");
                    ws_New.Cells[7, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[3]");
                    ws_New.Cells[7, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                }
                else if (cell.InnerText.Contains("外資"))
                {
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[2]");
                    ws_New.Cells[8, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                    cell = table.SelectSingleNode("tr[" + i.ToString() + "]/td[3]");
                    ws_New.Cells[8, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
                }
            }
            // 104(2015)之前沒有自營商(避險)
            //// 自營商
            //HtmlNode cell = table.SelectSingleNode("tr[3]/td[2]");
            //ws_New.Cells[5,3].Value2 = cell.InnerText.Replace(",",string.Empty).Trim();
            //cell = table.SelectSingleNode("tr[3]/td[3]");
            //ws_New.Cells[5,4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();

            //// 自營商(避險)
            //cell = table.SelectSingleNode("tr[4]/td[2]");
            //ws_New.Cells[6, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
            //cell = table.SelectSingleNode("tr[4]/td[3]");
            //ws_New.Cells[6, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();

            //// 投信
            //cell = table.SelectSingleNode("tr[5]/td[2]");
            //ws_New.Cells[7, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
            //cell = table.SelectSingleNode("tr[5]/td[3]");
            //ws_New.Cells[7, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();

            //// 外資
            //cell = table.SelectSingleNode("tr[6]/td[2]");
            //ws_New.Cells[8, 3].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();
            //cell = table.SelectSingleNode("tr[6]/td[3]");
            //ws_New.Cells[8, 4].Value2 = cell.InnerText.Replace(",", string.Empty).Trim();

            
        }
        */
        #endregion
        
        private void StockBig3(IE ieBig3,DateTime date,Worksheet ws_New)
        {
            ieBig3.SelectList(Find.ByName("yy")).SelectByValue(date.Year.ToString());
            ieBig3.SelectList(Find.ByName("mm")).SelectByValue(date.Month.ToString());
            ieBig3.SelectList(Find.ByName("dd")).SelectByValue(date.Day.ToString());
            ieBig3.Link(Find.ByClass("button search")).Click();

            WaitforDataLoading(ieBig3, 20);
            Table table = ieBig3.Table(Find.ById("report-table"));
            foreach (TableRow row in table.TableRows)
            {
                if (row.TableCells.Count > 0)
                {
                    switch (row.TableCells[0].Text)
                    {
                        case "自營商(自行買賣)":
                            ws_New.Cells[5, 3].Value2 = row.TableCells[1].Text.Replace(",", string.Empty);
                            ws_New.Cells[5, 4].Value2 = row.TableCells[2].Text.Replace(",", string.Empty);
                            break;
                        case "自營商(避險)":
                            ws_New.Cells[6, 3].Value2 = row.TableCells[1].Text.Replace(",", string.Empty);
                            ws_New.Cells[6, 4].Value2 = row.TableCells[2].Text.Replace(",", string.Empty);
                            break;
                        case "投信":
                            ws_New.Cells[7, 3].Value2 = row.TableCells[1].Text.Replace(",", string.Empty);
                            ws_New.Cells[7, 4].Value2 = row.TableCells[2].Text.Replace(",", string.Empty);
                            break;
                        case "外資及陸資":
                            ws_New.Cells[8, 3].Value2 = row.TableCells[1].Text.Replace(",", string.Empty);
                            ws_New.Cells[8, 4].Value2 = row.TableCells[2].Text.Replace(",", string.Empty);
                            break;
                        default:
                            break;
                    }
                }
            }
        }
        private void TradingInfo(IE ie,DateTime date,Worksheet ws_New)
        {
            // 成交量
            string AllSold = "";
            string Point = string.Empty;
            
            if (!isThisMonth(ie,date))
            {
                ie.SelectList(Find.ByName("yy")).SelectByValue(date.Year.ToString());
                ie.SelectList(Find.ByName("mm")).SelectByValue(date.Month.ToString());
                ie.Link(Find.ByClass("button search")).Click();
                WaitforDataLoading(ie, 10);
            }
            
            Table table = ie.Table(Find.ById("report-table"));

            for (int i = 1; i < table.TableRows.Count; i++)
            {
                if (table.TableRows[i].TableCells[0].Text.Contains(ROCDate(date,false,1)))
                {
                    AllSold = table.TableRows[i].TableCells[2].Text.Replace(",", string.Empty).Trim();
                    Point = table.TableRows[i].TableCells[4].Text.Replace(",", string.Empty).Trim();
                    break;
                }
            }
            if (!string.IsNullOrEmpty(AllSold))
            {
                ws_New.Range["F5"].Formula = "=(C9+D9)/" + AllSold + "/2";
                ws_New.Cells[5, 12].Value2 = Point;
                ws_New.Calculate();
                if (Convert.ToInt32(ws_New.Cells[8,12].Value2) > 0)
                    ws_New.Range["L8"].Font.Color = -16776961;
                else
                    ws_New.Range["L8"].Font.Color = -11489280;
            }
            ie = null;
        }

      

        private void MarginTrading(IE ie,DateTime date,Worksheet ws_New)
        {
            long t_Buy, t_Sell, y_Buy, y_Sell;
            double Buy;
            
            if (!ie.Div(Find.ByClass("title")).Text.Contains(ROCDate(date,true,2)))
            {
                ie.SelectList(Find.ByName("yy")).SelectByValue(date.Year.ToString());
                ie.SelectList(Find.ByName("mm")).SelectByValue(date.Month.ToString());
                ie.SelectList(Find.ByName("dd")).SelectByValue("1");
                ie.Link(Find.ByClass("button search")).Click();
            }
            ie.SelectList(Find.ByName("dd")).SelectByValue(date.Day.ToString());
            ie.Link(Find.ByClass("button search")).Click();
            
            Table table = ie.Table(Find.ById("credit-table"));
            if (WaitForMarginTable(ref table, ie, 10))
            {
                long.TryParse(table.TableRows[2].TableCells[4].Text.Replace(",", string.Empty), out y_Sell);
                long.TryParse(table.TableRows[2].TableCells[5].Text.Replace(",", string.Empty), out t_Sell);
                long.TryParse(table.TableRows[3].TableCells[4].Text.Replace(",", string.Empty), out y_Buy);
                long.TryParse(table.TableRows[3].TableCells[5].Text.Replace(",", string.Empty), out t_Buy);
                Buy = Math.Round((double)((t_Buy - y_Buy) / 100000), 3);

                // 資增減
                ws_New.Cells[9, 10].Value2 = (Buy >= 0) ? "資增" + Buy.ToString() + "億" : "資減" + Math.Abs(Buy).ToString() + "億";

                // 券增減
                ws_New.Cells[9, 11].Value2 = (t_Sell > y_Sell) ? "券增" + (t_Sell - y_Sell).ToString() + "張" : "券減" + (y_Sell - t_Sell).ToString() + "張";

                // 上色
                if (ws_New.Range["J9"].Value2.ToString().Contains("資增"))
                    ws_New.Range["J9"].Font.Color = -16776961;
                else
                    ws_New.Range["J9"].Font.Color = -11489280;

                if (ws_New.Range["K9"].Value2.ToString().Contains("券增"))
                    ws_New.Range["K9"].Font.Color = -16776961;
                else
                    ws_New.Range["K9"].Font.Color = -11489280;
            }
            else
            {
                MessageBox.Show("信用交易資料抓取發生逾時，請重新再試");
                throw new TimeoutException();
            }
        }

        private string ROCDate(DateTime date,bool noDate,int Format = 1)
        {
            TaiwanCalendar twCalendar = new TaiwanCalendar();
            if (Format == 1)
                return twCalendar.GetYear(date).ToString() + "/" + date.ToString("MM/dd");
            else if (Format == 2)
            {
                string rocdate;
                rocdate = twCalendar.GetYear(date).ToString() + "年" + date.ToString("MM") + "月";
                if (!noDate)
                    rocdate += date.ToString("dd") + "日";
                return rocdate;
            }
            else
                return string.Empty;
            
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            if (duration == null)
                duration = new List<DateTime>();
            DateTime StartDate = e.Start;
            DateTime EndDate = e.End;
            StringBuilder stbr = new StringBuilder();

            while (StartDate <= EndDate)
            {
                if(!duration.Exists(d => d.Equals(StartDate)))
                    duration.Add(StartDate);
                StartDate = StartDate.AddDays(1);
            }
            rtbSelectedDate.Text = "";
            duration.Sort();
            foreach(DateTime d in duration)
            {
                stbr.AppendLine(d.ToShortDateString());
            }
            rtbSelectedDate.Text = stbr.ToString();
        }

        private void btnResetDuration_Click(object sender, EventArgs e)
        {
            duration.Clear();
            rtbSelectedDate.Text = string.Empty;
        }
        private void DateTimeModify(Worksheet ws_New,DateTime date)
        {
            // 標題
            ws_New.Cells[1, 2].Value2 = ws_New.Cells[1, 2].Value2.ToString().Replace("yyyy/MM/dd", date.ToShortDateString());

            
            // 工作頁名稱
            ws_New.Name = date.ToString("yyyyMMdd");
        }

        private void WaitforDataLoading(IE ie,int Timeout)
        {
            DateTime StartTime = DateTime.Now;
            Div LoadingMessage;
            Thread.Sleep(1500);
            while (true)
            {
                LoadingMessage = ie.Div(Find.ById("loading"));
                if (LoadingMessage.Style.Display.ToLower().Contains("none"))
                    break;
                
                if (DateTime.Now.Subtract(StartTime).Seconds > Timeout)
                    break;
            }
        }

        private void btn_SelectDir_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            DialogResult result = fbd.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
            {
                txtSaveDir.Text = fbd.SelectedPath;
            }

        }

        private void btn_Split_Click(object sender, EventArgs e)
        {
            string FilePath = txtFilePath.Text.Replace("\"", string.Empty);
            string SaveFolder = txtSaveDir.Text.Replace("\"", string.Empty);

            if (File.Exists(FilePath) && Directory.Exists(SaveFolder))
            {
                List<string> Months = new List<string>();
                string pattern = @"\d{8}";
                Regex r = new Regex(pattern);
                Microsoft.Office.Interop.Excel.Application xlsApp = null;
                Workbook wb = null;
                Workbook wb_New = null;
                bool isSuccess = true;
                try
                {
                    xlsApp = new Microsoft.Office.Interop.Excel.Application();
                    wb = xlsApp.Workbooks.Open(FilePath);

                    // 先記錄所有的月份
                    foreach (Worksheet ws in wb.Worksheets)
                    {
                        if (r.IsMatch(ws.Name))
                        {
                            if (!Months.Exists(x => x.Equals(ws.Name.Substring(0, 6))))
                                Months.Add(ws.Name.Substring(0, 6));
                        }
                    }

                    // 依照月份把Excel拆分到新Excel
                    for (int i = Months.Count - 1; i >= 0; i--)
                    {
                        wb_New = xlsApp.Workbooks.Add();
                        for (int j = wb.Worksheets.Count; j > 0; j--)
                        {
                            if (wb.Worksheets[j].Name.Contains(Months[i]))
                            {
                                //ws_Template.Copy(Type.Missing, wb.Worksheets[wb.Worksheets.Count]);
                                wb.Worksheets[j].Move(wb_New.Worksheets[1], Type.Missing);
                            }
                        }
                        
                        // 刪除不需要的工作頁
                        for (int k = wb_New.Worksheets.Count; k > 0; k--)
                        {
                            if (!r.IsMatch(wb_New.Worksheets[k].Name))
                                wb_New.Worksheets[k].Delete();
                        }
                        wb_New.SaveAs(SaveFolder + "\\" + Months[i] + "_Finance.xlsx");
                        wb_New.Close(false);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    isSuccess = false;
                }
                finally
                {
                    wb.Close(true);
                    if (wb != null)
                        Marshal.FinalReleaseComObject(wb);
                    if (wb_New != null)
                        Marshal.FinalReleaseComObject(wb_New);
                    if (xlsApp != null)
                        Marshal.FinalReleaseComObject(xlsApp);
                    if (isSuccess)
                        MessageBox.Show("拆分完成！");
                    else
                        MessageBox.Show("拆分過程出現錯誤，請檢查後重試");
                    this.Close();
                }
            }
            else
            {
                MessageBox.Show("檔案不存在或儲存路徑有誤，請檢查！");
            }
        }

        private void btn_Duration_Click(object sender, EventArgs e)
        {
            F_InputDuration Window = new F_InputDuration(this);
            Window.Show();
        }

        private bool isThisMonth(IE ie, DateTime date)
        {
            Div div = ie.Div(Find.ByClass("title"));
            return div.Text.Contains(ROCDate(date, true, 2));
        }

        private bool WaitForDaySelectList(DateTime date,SelectList list,int Timeout)
        {
            int daycount = DateTime.DaysInMonth(date.Year, date.Month);
            DateTime StartTime = DateTime.Now;
            while(DateTime.Now.Subtract(StartTime).Seconds < Timeout)
            {
                if (list.Options.Count == daycount)
                    return true;
            }
            return false;
        }

        private bool WaitForMarginTable(ref Table table,IE ie,int Timeout)
        {
            DateTime StartTime = DateTime.Now;
            while(DateTime.Now.Subtract(StartTime).Seconds < Timeout)
            {
                if (table.TableRows.Count > 2)
                    return true;
                table = ie.Table(Find.ById("credit-table"));
            }
            return false;
        }
    }
}
