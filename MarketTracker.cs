using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Net;
using System.Windows.Forms;
using System.Xml;
using HtmlAgilityPack;
using Newtonsoft.Json.Linq;
using Data = System.Collections.Generic.KeyValuePair<string, double>;

namespace NSEDayMarketTracker
{
    /// <summary>
    /// Contains logic for the option tracker form.
    /// </summary>
    public partial class MarketTracker : Form
    {
        private const string NSEIndiaWebsiteURL = "https://www.nseindia.com";
        private const string NIFTYStockWatchURL = NSEIndiaWebsiteURL + "/live_market/dynaContent/live_watch/stock_watch/niftyStockWatch.json";
        private const string BankNIFTYStockWatchURL = NSEIndiaWebsiteURL + "/live_market/dynaContent/live_watch/stock_watch/bankNiftyStockWatch.json";
        private const string NavigationMenuURL = NSEIndiaWebsiteURL + "/common/xml/navigation.xml";
        private const string VIXDetailsJSONURL = NSEIndiaWebsiteURL + "/live_market/dynaContent/live_watch/VixDetails.json";

        private decimal baseNumber = 0;
        private decimal baseNumberPlus50 = 0;
        private decimal baseNumberPlus100 = 0;
        private decimal baseNumberPlus150 = 0;
        private decimal baseNumberPlus200 = 0;
        private decimal baseNumberMinus50 = 0;
        private decimal baseNumberMinus100 = 0;
        private decimal percentageThreshold = 5;

        List<Data> gannlevels = new List<Data>();
        /// <summary>
        /// The constructor.
        /// </summary>
        public MarketTracker()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Defines what to do when the form is loaded.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The current event object</param>
        private void MarketTracker_Load(object sender, EventArgs e)
        {
            refreshMarketButton.Visible = false;
        }

        /// <summary>
        /// Downloads JSON data from the URL.
        /// </summary>
        /// <param name="webResourceURL">The web resource URL of the JSON file.</param>
        /// <returns>JObject to readily read from.</returns>
        private JObject DownloadJSONDataFromURL(string webResourceURL)
        {
            string stockWatchJSONString = string.Empty;

            using(var webClient = new WebClient())
            {
                // Set headers to download the data
                webClient.Headers.Add("Accept: text/html, application/xhtml+xml, */*");
                webClient.Headers.Add("User-Agent: Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)");

                // Download the data
                stockWatchJSONString = webClient.DownloadString(webResourceURL);

                // Serialise it into a JObject
                JObject jObject = JObject.Parse(stockWatchJSONString);

                return jObject;
            }
        }

        /// <summary>
        /// Sets market open and close values.
        /// </summary>
        /// <param name="equitiesStockWatchJObject">The JObject to read from.</param>
        /// <returns>The open market base number.</returns>
        private int SetMarketOpenCloseValues(JObject equitiesStockWatchJObject)
        {
            int openMarketBaseNumber = 0;
            openValueLabel.Text = equitiesStockWatchJObject["latestData"][0]["open"].ToString();
            currentValueLabel.Text = equitiesStockWatchJObject["latestData"][0]["ltp"].ToString();

            // Calculate percentage difference
            decimal difference = Convert.ToDecimal(currentValueLabel.Text) - Convert.ToDecimal(openValueLabel.Text);
            decimal percentage = Math.Round(difference / Convert.ToDecimal(openValueLabel.Text) * 100, 2);
            string percentageDifference = "" + percentage;
            currentValuePercentageLabel.Text = percentageDifference;

            // Set colours according to result
            if(Convert.ToDecimal(currentValueLabel.Text) >= Convert.ToDecimal(openValueLabel.Text))
            {
                currentValueLabel.BackColor = Color.Green;
                currentValuePercentageLabel.BackColor = Color.Green;
            }
            else
            {
                currentValueLabel.BackColor = Color.Red;
                currentValuePercentageLabel.BackColor = Color.Red;
            }

            // Calculate the base open market value
            string precedingNumber = openValueLabel.Text.Split('.')[0];
            precedingNumber = precedingNumber.Replace(",", "");

            if(precedingNumber.Length == 5)
            {
                openMarketBaseNumber = Int32.Parse(precedingNumber) - (Int32.Parse(precedingNumber) % 100);
            }
            // Set Gann Details

            if (gannTableDataGridView.Rows.Count < 9)
            {
                CalculateGannSQRT9();
                gannTableDataGridView.Rows.Clear();
                foreach (var item in gannlevels)
                {
                    gannTableDataGridView.Rows.Add(item.Key, item.Value.ToString());
                }
            }
            //gannTableDataGridView.Rows.Add(data[0],data[1], data[2],data[3], data[4], data[5], data[6], data[7]);
            return openMarketBaseNumber;
        }

        /// <summary>
        /// Sets the date, time and the week for the data.
        /// </summary>
        /// <param name="equitiesStockWatchJObject">The JObject to read from.</param>
        private void SetDateTimeWeek(JObject equitiesStockWatchJObject)
        {
            dateLabel.Text = equitiesStockWatchJObject["time"].ToString();
            int weekNumber = 1 | DateTime.Now.Day / 7;
            weekNumber = (DateTime.Now.Day % 7 == 0) ? weekNumber - 1 : weekNumber;
            weekLabel.Text = "Week " + weekNumber;
        }

        /// <summary>
        /// Refreshes the data and resets all the values to the UI.
        /// </summary>
        /// <param name="sender">The sender object</param>*
        /// <param name="e">The current event object</param>
        private void RefreshMarketButton_Click(object sender, EventArgs e)
        {
            refreshMarketButton.Text = "Refreshing...";
            JObject equitiesStockWatchDataJObject = null;

            if(marketSelectComboBox.SelectedItem.ToString() == "NIFTY")
            {
                equitiesStockWatchDataJObject = DownloadJSONDataFromURL(NIFTYStockWatchURL);
            }
            else if(marketSelectComboBox.SelectedItem.ToString() == "Bank NIFTY")
            {
                equitiesStockWatchDataJObject = DownloadJSONDataFromURL(BankNIFTYStockWatchURL);
            }
            int openMarketBaseNumber = SetMarketOpenCloseValues(equitiesStockWatchDataJObject);

            SetDateTimeWeek(equitiesStockWatchDataJObject);

            string liveMarketURL = GetLiveMarketURL();

            if(marketSelectComboBox.SelectedItem.ToString() == "NIFTY")
            {
                HtmlNodeCollection workSetRows = DownloadMarketData(liveMarketURL, openMarketBaseNumber);
                RenderStrikePriceDayTable(workSetRows);
            }
            else if(marketSelectComboBox.SelectedItem.ToString() == "Bank NIFTY")
            {
                string bankNIFTYMarketURL = GetBankNIFTYMarketURL(liveMarketURL);
                HtmlNodeCollection workSetRows = DownloadMarketData(bankNIFTYMarketURL, openMarketBaseNumber);
                RenderStrikePriceDayTable(workSetRows);
            }
            
            refreshMarketButton.Text = "Refresh";
        }

        /// <summary>
        /// Downloads the navigation XML file and returns the live market URL.
        /// </summary>
        /// <returns>The live market URL</returns>
        private string GetLiveMarketURL()
        {
            string marketvalue = marketSelectComboBox.SelectedItem.ToString();
            string navigationXML = string.Empty;
            string liveMarketURL = string.Empty;

            using(var webClient = new WebClient())
            {
                // Set headers to download the data
                webClient.Headers.Add("Accept: text/html, application/xhtml+xml, */*");
                webClient.Headers.Add("User-Agent: Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)");

                // Download the data
                navigationXML = webClient.DownloadString(NavigationMenuURL);
                XmlReaderSettings xmlReaderSettings = new XmlReaderSettings
                {
                    IgnoreWhitespace = true
                };

                using(XmlReader xmlReader = XmlReader.Create(new StringReader(navigationXML), xmlReaderSettings))
                {
                    xmlReader.MoveToContent();
                    xmlReader.ReadToDescendant("item");
                    xmlReader.ReadToNextSibling("item");
                    xmlReader.ReadToDescendant("submenu");
                    xmlReader.ReadToNextSibling("submenu");
                    xmlReader.ReadToDescendant("submenuitem");
                    liveMarketURL = xmlReader.GetAttribute("link");

                    return NSEIndiaWebsiteURL + liveMarketURL;
                }
            }
        }

        /// <summary>
        /// Grabs the required trs from the market table after calculating the range from the base number.
        /// </summary>
        /// <param name="marketURL">The market URL</param>
        /// <param name="openMarketBaseNumber">The open market base number</param>
        /// <returns>HtmlNodeCollection</returns>
        private HtmlNodeCollection DownloadMarketData(string marketURL, int openMarketBaseNumber)
        {
            // Define the range
            baseNumber = Math.Round(Convert.ToDecimal(openMarketBaseNumber), 2);
            baseNumberPlus50 = baseNumber + 100;
            baseNumberPlus100 = baseNumber + 200;
            baseNumberPlus150 = baseNumber + 300;
            baseNumberPlus200 = baseNumber - 100;
            baseNumberMinus50 = baseNumber - 200;
            baseNumberMinus100 = baseNumber - 300;

            // Grab all rows
            var htmlWeb = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDocument = htmlWeb.Load(marketURL);

            HtmlNodeCollection tableRows = htmlDocument.DocumentNode.SelectNodes("//table[@id=\"octable\"]//tr");
            tableRows.RemoveAt(tableRows.Count - 1);
            tableRows.RemoveAt(0);
            tableRows.RemoveAt(0);

            // Get only those rows which contain values for the defined tange
            HtmlNodeCollection workSetRows = new HtmlNodeCollection(null);
            foreach(var currentTableRow in tableRows)
            {
                if(currentTableRow.InnerHtml.Contains(baseNumber.ToString()) || currentTableRow.InnerHtml.Contains(baseNumberPlus50.ToString())
                    || currentTableRow.InnerHtml.Contains(baseNumberPlus100.ToString()) || currentTableRow.InnerHtml.Contains(baseNumberMinus50.ToString())
                    || currentTableRow.InnerHtml.Contains(baseNumberMinus100.ToString()) || currentTableRow.InnerHtml.Contains(baseNumberPlus150.ToString())
                    || currentTableRow.InnerHtml.Contains(baseNumberPlus200.ToString()))
                {
                    workSetRows.Add(currentTableRow);
                }
            }

            return workSetRows;
        }

        /// <summary>
        /// Returns the Bank NIFTY page URL from the main market URL JavaScript snippet.
        /// </summary>
        /// <param name="marketURL">The market URL for Bank NIFTY.</param>
        /// <returns>The Bank NIFTY URL.</returns>
        private string GetBankNIFTYMarketURL(string marketURL)
        {
            // Load the web page and get the JavaScript
            var htmlWeb = new HtmlWeb();
            HtmlAgilityPack.HtmlDocument htmlDocument = htmlWeb.Load(marketURL);
            HtmlNodeCollection scriptTags = htmlDocument.DocumentNode.SelectNodes("//script[@type=\"text/javascript\"]");
            string bankNIFTYMarketURL = scriptTags[5].InnerHtml;

            // Process the JavaScript and get the Bank NIFTY URL
            bankNIFTYMarketURL = bankNIFTYMarketURL.Replace("\r\n", "");
            bankNIFTYMarketURL = bankNIFTYMarketURL.Replace("\t", "");
            bankNIFTYMarketURL = bankNIFTYMarketURL.Remove(0, bankNIFTYMarketURL.IndexOf("BANKNIFTY"));
            bankNIFTYMarketURL = bankNIFTYMarketURL.Remove(0, 28);
            bankNIFTYMarketURL = bankNIFTYMarketURL.Remove(bankNIFTYMarketURL.IndexOf(";"), bankNIFTYMarketURL.Length - bankNIFTYMarketURL.IndexOf(";"));
            bankNIFTYMarketURL = bankNIFTYMarketURL.Replace("'", "");

            return NSEIndiaWebsiteURL + "/" + bankNIFTYMarketURL;
        }

        /// <summary>
        /// Renders data into the Strike Price Day Table.
        /// </summary>
        /// <param name="workSetRows">A collection of HTML nodes</param>
        private void RenderStrikePriceDayTable(HtmlNodeCollection workSetRows)
        {
            strikePriceTableDataGridView.Rows.Clear();
            List<List<string>> strikePriceDayTableValues = new List<List<string>>();
            List<int> indicesOfRowsToHighlight = new List<int>();

            // Fetch the list of tds in the row
            foreach(HtmlNode currentWorkSetRow in workSetRows)
            {
                List<string> strikePriceDayTableRowValues = new List<string>();
                foreach(HtmlNode currentNodeInWorkSetRow in currentWorkSetRow.ChildNodes)
                {
                    if(currentNodeInWorkSetRow.Name == "td" && currentNodeInWorkSetRow.InnerText != "")
                    {
                        strikePriceDayTableRowValues.Add(currentNodeInWorkSetRow.InnerText.Trim());
                    }
                }
                strikePriceDayTableValues.Add(strikePriceDayTableRowValues);
            }

            // Remove unwanted data
            foreach(List<string> currentValuesSet in strikePriceDayTableValues)
            {
                currentValuesSet.RemoveAt(0);
                currentValuesSet.RemoveAt(1);
                currentValuesSet.RemoveAt(1);
                currentValuesSet.RemoveAt(2);
                currentValuesSet.RemoveAt(2);
                currentValuesSet.RemoveAt(2);
                currentValuesSet.RemoveAt(2);
                currentValuesSet.RemoveAt(2);
                currentValuesSet.RemoveAt(3);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(4);
                currentValuesSet.RemoveAt(5);
                currentValuesSet.Insert(0, "Dummy");
                currentValuesSet.Add("Put Writers");

                // And render
                strikePriceTableDataGridView.Rows.Add(currentValuesSet[0], currentValuesSet[1], currentValuesSet[2], currentValuesSet[3], currentValuesSet[4],
                    currentValuesSet[5], currentValuesSet[6]);
            }

            // Find rows to highlight
            foreach(DataGridViewRow currentRowCells in strikePriceTableDataGridView.Rows)
            {
                for(int currentRowCellsIndex = 0; currentRowCellsIndex < currentRowCells.Cells.Count; currentRowCellsIndex++)
                {
                    if(currentRowCells.Cells[1].Value.ToString().Contains("-"))
                    {
                        currentRowCells.Cells[0].Style.BackColor = Color.LightGreen;
                        currentRowCells.Cells[0].Value = "CEW Exiting";
                    }
                    else
                    {
                        currentRowCells.Cells[0].Style.BackColor = Color.PaleVioletRed;
                        currentRowCells.Cells[0].Value = "Call Writers";
                    }
                    currentRowCells.Cells[6].Style.BackColor = Color.LightGreen;
                }

                foreach(DataGridViewCell currentRowCell in currentRowCells.Cells)
                {
                    if(currentRowCell.Value.Equals(baseNumber.ToString() + ".00") || currentRowCell.Value.Equals(baseNumberPlus50.ToString() + ".00")
                        || currentRowCell.Value.Equals(baseNumberPlus100.ToString() + ".00"))
                    {
                        indicesOfRowsToHighlight.Add(currentRowCell.RowIndex);
                    }
                }
            }

            // Highlight range cells with blue
            foreach(int rowIndex in indicesOfRowsToHighlight)
            {
                DataGridViewRow currentRowCells = strikePriceTableDataGridView.Rows[rowIndex];

                foreach(DataGridViewCell currentRowCell in currentRowCells.Cells)
                {
                    if(!(currentRowCell.Value.ToString().Equals("CEW Exiting") || currentRowCell.Value.ToString().Equals("Call Writers")
                        || currentRowCell.Value.ToString().Equals("Put Writers")))
                    {
                        currentRowCell.Style.BackColor = Color.LightBlue;
                    }
                }
            }

            UpdateDayTable(indicesOfRowsToHighlight);
            // Check for Option strategy
            if(OptionStrategyGridView.Rows.Count <= 0)
                LookForOptionStrategy(strikePriceDayTableValues);
            //
        }

        private void LookForOptionStrategy(List<List<string>> strikePriceDayTableValues)
        {        
            string trade = "";
            //Get Current ATM
            var currvalue = currentValueLabel.Text.Split('.')[0];
            currvalue = currvalue.Replace(",", "");
            int close = Convert.ToInt32(currvalue);
            int mod = (close % 100);
            var atm = close - mod;
            if(mod > 50)
            {
                atm = atm + 100;
            }

            decimal ceval, peval;
            string strtime = DateTime.Now.ToString("hh:mm");
            int cesp = 0;
            int pesp = 0;
            // check short straddle        
            if (string.IsNullOrEmpty(trade))
            {
                cesp = atm;
                pesp = atm;
                bool retval = CheckOptionStrategy( atm,atm, strikePriceDayTableValues,out ceval, out peval);
                if (retval)
                {
                    trade = "STRADDLE";
                    OptionStrategyGridView.Rows.Add(ceval.ToString(), cesp, trade, pesp, peval.ToString(), strtime);
                    Console.Beep(5000, 5000);
                }

            }
            // check short strangle
            if (string.IsNullOrEmpty(trade))
            {
                cesp = atm + 100;
                pesp = atm - 100;
                bool retval = CheckOptionStrategy( cesp, pesp, strikePriceDayTableValues, out ceval, out peval);
                if (retval)
                {
                    trade = "NEAR STRANGLE";
                    OptionStrategyGridView.Rows.Add(ceval.ToString(), cesp, trade, pesp, peval.ToString(), strtime);
                    Console.Beep(5000, 5000);
                }
            }
            // check short strangle
            if (string.IsNullOrEmpty(trade))
            {
                cesp = atm + 200;
                pesp = atm - 200;
                bool retval = CheckOptionStrategy( cesp, pesp, strikePriceDayTableValues, out ceval, out peval);
                if (retval)
                {
                    trade = "FAR STRADDLE";
                    OptionStrategyGridView.Rows.Add(ceval.ToString(), cesp, trade, pesp, peval.ToString(), strtime);
                    Console.Beep(5000, 5000);
                }
            }

            


        }

        private bool CheckOptionStrategy(int sp1,int sp2, List<List<string>> strikePriceDayTableValues,
            out decimal ce, out decimal pe)
        {
            bool bsp1 = false;
            bool bsp2 = false;
            ce = 0;
            pe = 0;
            foreach (var spdata in strikePriceDayTableValues)
            {
                //bsp1 = false;
                //bsp2 = false;
                var strsp = spdata[3];
                strsp = strsp.Split('.')[0];
                strsp = strsp.Replace(",", "");
                var sp = Convert.ToInt32(strsp);
               
                if (sp == sp1)
                {
                    ce = Convert.ToDecimal(spdata[2]);
                    bsp1 = true;
                }

                if (sp == sp2)
                {
                    pe = Convert.ToDecimal(spdata[4]);
                    bsp2 = true;
                }

                if (bsp1 && bsp2)
                {
                    var percentage = CalcualatePercentageDiff(ce, pe);
                    if (percentage <= percentageThreshold)
                    {
                        return true;
                    }
                    else
                        return false;

                }

            }

           

            return false;
        }

        private decimal CalcualatePercentageDiff(decimal ce, decimal pe)
        {
            decimal difference = Math.Abs(ce - pe);
            decimal percentage = 0;
            if (ce > pe)
            {
                //percentage = Math.Round(difference / ce * 100, 2);
                percentage = CalCulatePercentageIncrease(ce, pe);
            }
            else
            {
                //percentage = Math.Round(difference / pe * 100, 2);
                percentage = CalCulatePercentageIncrease(pe, ce);
            }
            return percentage;
        }
        private decimal CalCulatePercentageIncrease(decimal big, decimal small)
        {
            var incr = big - small;
            if (small == 0)
                small = 1;
            var perincr = (incr / small) * 100;

            var inc = Math.Round(perincr, 2);
            //string strinc = String.Format("{0:F2}", inc);
            return inc;// perincr;
        }

        /// <summary>
        /// Updates day table with rows.
        /// </summary>
        /// <param name="rowsIndex">The list of rows indices with which to work upon.</param>
        private void UpdateDayTable(List<int> rowsIndex)
        {
            List<string> ceValues = new List<string>();
            List<string> peValues = new List<string>();

            foreach(int currentRowIndex in rowsIndex)
            {
                DataGridViewRow currentHighlightedRowCells = strikePriceTableDataGridView.Rows[currentRowIndex];
                ceValues.Add(currentHighlightedRowCells.Cells[1].Value.ToString());
                peValues.Add(currentHighlightedRowCells.Cells[5].Value.ToString());
            }

            int ceTotal = 0;
            int peTotal = 0;

            foreach(string currentCEValue in ceValues)
            {
                ceTotal += Int32.Parse(currentCEValue.Replace(",", ""));
            }

            foreach(string currentPEValue in peValues)
            {
                peTotal += Int32.Parse(currentPEValue.Replace(",", ""));
            }

            int percentage = (peTotal - ceTotal) / ceTotal * 100;
            dayTableDataGridView.Rows.Add(DateTime.Now.ToLocalTime().ToLongTimeString(), currentValueLabel.Text, ceTotal, percentage, peTotal);
        }

        /// <summary>
        /// Refreshes the data after every 5 seconds.
        /// </summary>
        /// <param name="sender">The sender object</param>
        /// <param name="e">The current event object</param>
        private void RefreshTimer_Tick(object sender, EventArgs eventArgs)
        {
            if(marketSelectComboBox.SelectedItem != null)
            {
                refreshMarketButton.Visible = true;
                refreshMarketButton.PerformClick();
            }
        }

        /// <summary>
        /// Sets values for VIX.
        /// </summary>
        /// <param name="vixJObject"></param>
        private void SetVIXValues(JObject vixJObject)
        {
            vixValueLabel.Text = vixJObject["currentVixSnapShot"][0]["CURRENT_PRICE"].ToString();
            vixValuePercentageLabel.Text = vixJObject["currentVixSnapShot"][0]["PERC_CHANGE"].ToString();
            string previousVIXClose = vixJObject["currentVixSnapShot"][0]["PREV_CLOSE"].ToString();

            // Calculate percentage difference
            decimal difference = Convert.ToDecimal(vixValueLabel.Text) - Convert.ToDecimal(previousVIXClose);
            decimal percentage = Math.Round(difference / Convert.ToDecimal(vixValueLabel.Text) * 100, 2);
            string percentageDifference = "" + percentage;

            // Set colours according to result
            if(Convert.ToDecimal(vixValueLabel.Text) >= Convert.ToDecimal(previousVIXClose))
            {
                vixValueLabel.BackColor = Color.Green;
                vixValuePercentageLabel.BackColor = Color.Green;
            }
            else
            {
                vixValueLabel.BackColor = Color.Red;
                vixValuePercentageLabel.BackColor = Color.Red;
            }
        }

        private void CalculateGannSQRT9()
        {
            var sqrt = 0.0;
            List<double> gann = new List<double>();
            double open = Convert.ToDouble(openValueLabel.Text);
            sqrt = Math.Sqrt(open);
            //console.log('SQRT ', sqrt, element.master.open);
            //return;

            var val1 = Math.Round(sqrt); //try converting to int
            var val2 = val1 - 1;
            var val3 = val1 + 1;
            var val4 = val3 + 1;
            var flo = 0.0;
            //gann[0] = val2 * val2; 
            gann.Add(val2 * val2);
            var xyz = 1;
            for (var i = 1; i <= 8; ++i)
            {
                flo = val2 + i * 0.125;
                //gann[xyz] = flo * flo;
                gann.Add(flo * flo);
                ++xyz;
            }
            for (var i = 1; i <= 8; ++i)
            {
                flo = val1 + i * 0.125;
                //this.gann[xyz] = flo * flo;
                gann.Add(flo * flo);
                ++xyz;
            }
            for (var i = 1; i <= 8; ++i)
            {
                flo = val3 + i * 0.125;
                //this.gann[xyz] = flo * flo;
                gann.Add(flo * flo);
                ++xyz;
            }
            for (var i = 1; i <= 8; ++i)
            {
                flo = val4 + i * 0.125;
                //this.gann[xyz] = flo * flo;
                gann.Add(flo * flo);
                ++xyz;
            }

            SetGannTable(gann);

        }

        private void SetGannTable(List<double> gann)
        {
            //console.log('setting Gann data....', element.quote);
            //finalGann = [];
            double open = Convert.ToDouble(openValueLabel.Text);
            for (var i = 0; i <= 33; ++i)
            {
                if (open < gann[i])
                {
                    gannlevels.Add(new Data("Resistance 4", gann[i + 3]));

                    gannlevels.Add(new Data("Resistance 3", gann[i + 2]));

                    gannlevels.Add(new Data("Resistance 2", gann[i + 1]));

                    gannlevels.Add(new Data("Resistance 1", gann[i]));

                    gannlevels.Add(new Data("Open", open));

                    gannlevels.Add(new Data("Supprt 1", gann[i - 1]));

                    gannlevels.Add(new Data("Supprt 2", gann[i - 2]));

                    gannlevels.Add(new Data("Supprt 3", gann[i - 3]));

                    gannlevels.Add(new Data("Supprt 4", gann[i - 4]));
                    //element.gannlevels = finalGann;                   
                    break;
                }
            }
        }
    }
}