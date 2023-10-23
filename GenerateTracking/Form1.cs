using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Google.Apis.Services;

namespace GenerateTracking
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        public void GenerateTrackingYO()
        {
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            string ApplicationName = "Your Application Name";
            string SpreadsheetId = "1WlPa_ANWw3pf96Gbc_vIVUoNdwHjl13JdSEhiESpcZA";
            string CredentialsFilePath = "./client_secret.json";

            string filePath = "timeYO.txt";
            
            UserCredential credential;

            using (var stream = new FileStream("gg.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("client_secret.json", true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string range = "工作表1!A5000:E1000000";

            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;

            if (values != null && values.Count > 0)
            {

                var csvContent_SA = new StringBuilder();
                var csvContent_SZ = new StringBuilder();
                var csvContent_AG = new StringBuilder();
                var csvContent_AS = new StringBuilder();
                var csvContent_AP = new StringBuilder();
                var csvContent_FL = new StringBuilder();
                string header = "Order ID,Order Item ID,Carrier Slug,Tracking Number";
                csvContent_SA.AppendLine(header);
                csvContent_SZ.AppendLine(header);
                csvContent_AG.AppendLine(header);
                csvContent_AS.AppendLine(header);
                csvContent_AP.AppendLine(header);
                csvContent_FL.AppendLine(header);

                bool[] checks = new bool[] { false, false, false, false, false, false };

                string startDate = File.ReadAllText(filePath);
                DateTime currentDate = DateTime.Now.AddDays(-6);
                int month = currentDate.Month;
                int day = currentDate.Day;
                string endDate = month.ToString() + day.ToString();
                foreach (var row in values)
                {

                    do { row.Add(""); } while (row.Count < 5);
                    string time = row[0].ToString();
                    if (string.Compare(time, startDate) >= 0 && string.Compare(time, endDate) <= 0)
                    {
                        string orderNumber = row[3].ToString();
                        string logisticTracking = row[4].ToString();
                        string carrier = getCarrier(logisticTracking);

                        string orderNumberAfter = orderNumber.Length > 2 ? orderNumber.Substring(2) : "0_0";
                        string shopName = orderNumber.Length > 2 ? orderNumber.Substring(0, 2) : "00";

                        var order = orderNumberAfter.Split('_');
                        string textExport = order[0] + "," + order[1] + "," + carrier + "," + logisticTracking;
                        if (shopName == "SA")
                        {
                            checks[0] = true;
                            csvContent_SA.AppendLine(textExport);
                        }
                        else if (shopName == "SZ")
                        {
                            checks[1] = true;
                            csvContent_SZ.AppendLine(textExport);
                        }
                        else if (shopName == "AG")
                        {
                            checks[2] = true;
                            csvContent_AG.AppendLine(textExport);
                        }
                        else if (shopName == "AS")
                        {
                            checks[3] = true;
                            csvContent_AS.AppendLine(textExport);
                        }
                        else if (shopName == "AP")
                        {
                            checks[4] = true;
                            csvContent_AP.AppendLine(textExport);
                        }
                        else if (shopName == "FL")
                        {
                            checks[5] = true;
                            csvContent_FL.AppendLine(textExport);
                        }
                    }
                }
                if (checks[0] == true)
                {
                    File.WriteAllText("tracking_SA.csv", csvContent_SA.ToString());
                }
                if (checks[1] == true)
                {
                    File.WriteAllText("tracking_SZ.csv", csvContent_SZ.ToString());
                }
                if (checks[2] == true)
                {
                    File.WriteAllText("tracking_AG.csv", csvContent_AG.ToString());
                }
                if (checks[3] == true)
                {
                    File.WriteAllText("tracking_AS.csv", csvContent_AS.ToString());
                }
                if (checks[4] == true)
                {
                    File.WriteAllText("tracking_AP.csv", csvContent_AP.ToString());
                }
                if (checks[5] == true)
                {
                    File.WriteAllText("tracking_FL.csv", csvContent_FL.ToString());
                }

                File.WriteAllText(filePath, endDate);

                MessageBox.Show("Export Success");
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
        public void GenerateTrackingFM()
        {
            string[] Scopes = { SheetsService.Scope.SpreadsheetsReadonly };
            string ApplicationName = "Your Application Name";
            string SpreadsheetId = "16QWM0OhOsicTJKaOCp4NzwNxeJNIV1MXVIoTtfvgV9E";
            string CredentialsFilePath = "./client_secret.json";

            string filePath = "timeFM.txt";

            UserCredential credential;

            using (var stream = new FileStream("gg.json", FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore("client_secret.json", true)).Result;
            }

            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            string range = "Orders!A900:G1000000";

            SpreadsheetsResource.ValuesResource.GetRequest request =
                service.Spreadsheets.Values.Get(SpreadsheetId, range);
            ValueRange response = request.Execute();
            IList<IList<object>> values = response.Values;

            if (values != null && values.Count > 0)
            {

                var csvContent_SA = new StringBuilder();
                var csvContent_SZ = new StringBuilder();
                var csvContent_AG = new StringBuilder();
                var csvContent_AS = new StringBuilder();
                var csvContent_AP = new StringBuilder();
                var csvContent_FL = new StringBuilder();
                string header = "Order ID,Order Item ID,Carrier Slug,Tracking Number";
                csvContent_SA.AppendLine(header);
                csvContent_SZ.AppendLine(header);
                csvContent_AG.AppendLine(header);
                csvContent_AS.AppendLine(header);
                csvContent_AP.AppendLine(header);
                csvContent_FL.AppendLine(header);

                bool[] checks = new bool[] { false, false, false, false, false, false };

                string startDate = File.ReadAllText(filePath);
                DateTime currentDate = DateTime.Now.AddDays(-6);
                int month = currentDate.Month;
                int day = currentDate.Day;
                string endDate = month.ToString() + day.ToString();
                foreach (var row in values)
                {

                    do { row.Add(""); } while (row.Count < 7);
                    string time = row[0].ToString();
                    if (string.Compare(time, startDate) >= 0 && string.Compare(time, endDate) <= 0)
                    {
                        string orderNumber = row[3].ToString();
                        string logisticTracking = row[6].ToString();
                        string carrier = getCarrier(logisticTracking);

                        string orderNumberAfter = orderNumber.Length > 2 ? orderNumber.Substring(2) : "0_0";
                        string shopName = orderNumber.Length > 2 ? orderNumber.Substring(0, 2) : "00";

                        var order = orderNumberAfter.Split('_');
                        string textExport = order[0] + "," + order[1] + "," + carrier + "," + logisticTracking;
                        if (shopName == "SA")
                        {
                            checks[0] = true;
                            csvContent_SA.AppendLine(textExport);
                        }
                        else if (shopName == "SZ")
                        {
                            checks[1] = true;
                            csvContent_SZ.AppendLine(textExport);
                        }
                        else if (shopName == "AG")
                        {
                            checks[2] = true;
                            csvContent_AG.AppendLine(textExport);
                        }
                        else if (shopName == "AS")
                        {
                            checks[3] = true;
                            csvContent_AS.AppendLine(textExport);
                        }
                        else if (shopName == "AP")
                        {
                            checks[4] = true;
                            csvContent_AP.AppendLine(textExport);
                        }
                        else if (shopName == "FL")
                        {
                            checks[5] = true;
                            csvContent_FL.AppendLine(textExport);
                        }
                    }
                }
                if (checks[0] == true)
                {
                    File.WriteAllText("tracking_SA.csv", csvContent_SA.ToString());
                }
                if (checks[1] == true)
                {
                    File.WriteAllText("tracking_SZ.csv", csvContent_SZ.ToString());
                }
                if (checks[2] == true)
                {
                    File.WriteAllText("tracking_AG.csv", csvContent_AG.ToString());
                }
                if (checks[3] == true)
                {
                    File.WriteAllText("tracking_AS.csv", csvContent_AS.ToString());
                }
                if (checks[4] == true)
                {
                    File.WriteAllText("tracking_AP.csv", csvContent_AP.ToString());
                }
                if (checks[5] == true)
                {
                    File.WriteAllText("tracking_FL.csv", csvContent_FL.ToString());
                }

                File.WriteAllText(filePath, endDate);

                MessageBox.Show("Export Success");
            }
            else
            {
                MessageBox.Show("No data");
            }
        }
        private void btnExportYO_Click(object sender, EventArgs e)
        {
            GenerateTrackingYO();
        }
        

        public string getCarrier(string trackingNumber)
        {
            var carrierSlug = "";
            if (trackingNumber.StartsWith("YT"))
            {
                carrierSlug = "yun-express-cn";
            }
            else if (trackingNumber.StartsWith("007") && trackingNumber.Length == 20)
            {
                carrierSlug = "new-zealand-post";
            }
            else if (trackingNumber.StartsWith("YDH"))
            {
                carrierSlug = "ydh";
            }
            else if (trackingNumber.StartsWith("JPS"))
            {
                carrierSlug = "jpsgj";
            }
            else if (trackingNumber.StartsWith("JY"))
            {
                carrierSlug = "uniuni";
            }
            else if (trackingNumber.StartsWith("SF"))
            {
                carrierSlug = "sf-express-cn";
            }
            else if (trackingNumber.StartsWith("JJD"))
            {
                carrierSlug = "yodel";
            }
            else if (trackingNumber.StartsWith("DY") || trackingNumber.StartsWith("3A"))
            {
                carrierSlug = "cne";
            }
            else if (trackingNumber.StartsWith("FJ"))
            {
                carrierSlug = "royal-mail";
            }
            else if (trackingNumber.StartsWith("H100"))
            {
                carrierSlug = "hermes-de";
            }
            else if (trackingNumber.StartsWith("LR"))
            {
                carrierSlug = "china-post";
            }
            else if (trackingNumber.StartsWith("34W") || (trackingNumber.StartsWith("36") && trackingNumber.Length == 23))
            {
                carrierSlug = "ubi";
            }
            else if (trackingNumber.StartsWith("33XH"))
            {
                carrierSlug = "australia-post";
            }
            else if (trackingNumber.StartsWith("33XJ"))
            {
                carrierSlug = "australia-post";
            }
            else if (trackingNumber.StartsWith("YHT"))
            {
                carrierSlug = "yht";
            }
            else if (trackingNumber.StartsWith("JC"))
            {
                carrierSlug = "jcex";
            }
            else if (trackingNumber.StartsWith("LY"))
            {
                carrierSlug = "china-post";
            }
            else if (trackingNumber.StartsWith("UG") || trackingNumber.StartsWith("UH"))
            {
                carrierSlug = "yanwen";
            }
            else if (trackingNumber.StartsWith("LV"))
            {
                carrierSlug = "china-post";
            }
            else if (trackingNumber.StartsWith("LH"))
            {
                carrierSlug = "china-post";
            }
            else if (trackingNumber.StartsWith("4PX"))
            {
                carrierSlug = "4px";
            }
            else if (trackingNumber.StartsWith("1Z"))
            {
                carrierSlug = "ups";
            }
            else if (trackingNumber.StartsWith("US"))
            {
                carrierSlug = "jt-international";
            }
            else if (trackingNumber.Length == 10)
            {
                carrierSlug = "dhl";
            }
            else if (trackingNumber.Length == 12)
            {
                carrierSlug = "fedex";
            }
            else if (trackingNumber.Length == 16)
            {
                carrierSlug = "canada-post";
            }
            else if (trackingNumber.Length >= 26)
            {
                carrierSlug = "usps";
            }
            else if (trackingNumber.Split('|').Length > 1)
            {
                carrierSlug = trackingNumber.Split('|')[1];
            }
            else
            {
                carrierSlug = "NOTFOUND";
            }
            return carrierSlug;
        }

        private void btnExportFM_Click(object sender, EventArgs e)
        {
            GenerateTrackingFM();
        }
    }
}
