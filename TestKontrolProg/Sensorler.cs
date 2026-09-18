using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing;

namespace TestKontrolProg
{
    public partial class Sensorler : Form
    {
        private Form1 _form1;
        private Button _c600ReadButton;

        public Sensorler(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
            AddC600ReadButton();
        }

        private void AddC600ReadButton()
        {
            _c600ReadButton = new Button
            {
                Name = "buttonC600Read",
                Text = "C600 Verileri Çek",
                Size = new Size(120, 30),
                Location = new System.Drawing.Point(282, 360),
                UseVisualStyleBackColor = true
            };
            _c600ReadButton.Click += C600ReadButton_Click;
            Controls.Add(_c600ReadButton);
        }

        private async void C600ReadButton_Click(object sender, EventArgs e)
        {
            if (_c600ReadButton != null)
            {
                _c600ReadButton.Enabled = false;
                _c600ReadButton.Text = "Okunuyor...";
            }

            try
            {
                var fields = new Dictionary<TextBox, string>
                {
                    { textBox1, "5-TMPVAL" },          // Fresh air temperature
                    { textBox2, "FRESHHUM" },           // Fresh air humidity/function block
                    { textBox6, "SUPPLY_AIR_TEMP" },    // Supply air temperature
                    { textBox7, "1-HUMVAL" },           // Supply air humidity
                    { textBox11, "3-TMPVAL" },          // Return air temperature
                    { textBox12, "HUMVAL" },            // Return air humidity
                    { textBox13, "CO2VAL" },            // Return air CO2
                    { textBox16, "TMPVAL" },            // Exhaust air temperature
                    { textBox19, "1-TMPVAL" },          // After coil temperature
                    { textBox22, "7-TMPVAL" },          // Mixing air temperature
                    { textBox25, "ROOM_TEMP" },         // Room temperature
                    { textBox26, "ROOM_HUM" },          // Room humidity
                    { textBox28, "1-CO2VALUE" },        // Return CO2 sensor
                    { textBox29, "TEMPCONT_WATERT" }    // Water temperature
                };

                var tasks = fields.Select(async item =>
                {
                    try
                    {
                        string value = await C600Communication.ReadValueAsync(item.Value);
                        return new KeyValuePair<TextBox, string>(item.Key, value);
                    }
                    catch
                    {
                        return new KeyValuePair<TextBox, string>(item.Key, "-");
                    }
                }).ToArray();

                KeyValuePair<TextBox, string>[] results = await Task.WhenAll(tasks);

                int ok = 0;
                foreach (var result in results)
                {
                    result.Key.Text = result.Value;
                    if (result.Value != "-")
                        ok++;
                }

                // GenericJSON'da karşılığı bulunmayan alanlar bilinçli olarak boş bırakılır.
                textBox3.Text = "-";   // Fresh CO2
                textBox8.Text = "-";   // Supply CO2
                textBox17.Text = "-";  // Exhaust humidity
                textBox18.Text = "-";  // Exhaust CO2
                textBox20.Text = "-";  // After coil humidity
                textBox21.Text = "-";  // After coil CO2
                textBox23.Text = "-";  // Mixing humidity
                textBox24.Text = "-";  // Mixing CO2
                textBox27.Text = "-";  // Room sensor 2
                textBox30.Text = "-";  // Return CO2 air temperature
                textBox31.Text = "-";  // Return CO2 air CO2 (not a duplicate)

                MessageBox.Show(
                    ok + " adet C600 değeri okundu.\n\n" +
                    "Bağlantı: SCOPE / 127.0.0.1:4242 / USB",
                    "C600 Haberleşme",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "C600 / SCOPE üzerinden veriler okunamadı.\n\n" + ex.Message,
                    "C600 Haberleşme",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                if (_c600ReadButton != null)
                {
                    _c600ReadButton.Enabled = true;
                    _c600ReadButton.Text = "C600 Verileri Çek";
                }
            }
        }

        private void label1_Click(object sender, EventArgs e) { }
        private void groupBox2_Enter(object sender, EventArgs e) { }

        private void button1_Click(object sender, EventArgs e)
        {
            SensorDatasTake(groupBox1, Datas.FreshAirSensorValues);
            SensorDatasTake(groupBox2, Datas.SupplyAirSensorValues);
            SensorDatasTake(groupBox3, Datas.ReturnAirSensorValues);
            SensorDatasTake(groupBox4, Datas.ExhaustAirSensorValues);
            SensorDatasTake(groupBox5, Datas.AfterCoilAirSensorValues);
            SensorDatasTake(groupBox6, Datas.MixAirSensorValues);
            SensorDatasTake(groupBox7, Datas.RoomTempSensor1Values);
            SensorDatasTake(groupBox8, Datas.RoomTempSensor2Values);
            SensorDatasTake(groupBox9, Datas.ReturnCO2SensorValues);
            SensorDatasTake(groupBox10, Datas.WaterTempSensorValues);
            SensorDatasTake(groupBox11, Datas.ReturnCO2AirSensorValues);
            _form1.SetLabelSensorKontrolText();
            this.Hide();
        }

        private void SensorDatasTake(GroupBox groupBox, Dictionary<string, string> sensorDict)
        {
            var textBoxes = groupBox.Controls.OfType<TextBox>()
                .OrderBy(tb => tb.Tag?.ToString()).ToArray();

            foreach (TextBox tb in textBoxes)
            {
                string key = tb.Tag?.ToString();
                if (string.IsNullOrEmpty(key)) continue;
                string value1 = string.IsNullOrWhiteSpace(tb.Text) ? "-" : tb.Text.Trim();
                if (value1 != "-" && !double.TryParse(value1, out _)) value1 = "-";
                if (sensorDict.ContainsKey(key)) sensorDict[key] = value1;
                else sensorDict.Add(key, value1);
            }
        }

        private void Sensorler_Load(object sender, EventArgs e) { }
    }

    internal static class C600Communication
    {
        private const string ScopeBaseUrl = "http://127.0.0.1:4242";
        private const string JsonUsername = "ADMIN";
        private static readonly string JsonPassword = "SBT" + "Admin" + "!";
        private const string JsonPin = "6000";
        private const string JsonLanguage = "0";
        private const string JsonUser = "2";

        public static async Task<string> ReadValueAsync(string jsonId)
        {
            if (string.IsNullOrWhiteSpace(jsonId))
                throw new ArgumentException("JSON ID boş olamaz.", "jsonId");

            string url = ScopeBaseUrl
                + "/json.html?callback=?&fn=Read"
                + "&pin=" + Uri.EscapeDataString(JsonPin)
                + "&lng=" + Uri.EscapeDataString(JsonLanguage)
                + "&us=" + Uri.EscapeDataString(JsonUser)
                + "&id=" + Uri.EscapeDataString(jsonId);

            using (var client = new WebClient())
            {
                client.Encoding = Encoding.UTF8;
                client.Credentials = new NetworkCredential(JsonUsername, JsonPassword);
                string response = await client.DownloadStringTaskAsync(new Uri(url));
                return ParseValue(response);
            }
        }

        private static string ParseValue(string response)
        {
            if (string.IsNullOrWhiteSpace(response))
                throw new InvalidOperationException("SCOPE boş cevap döndürdü.");

            Match match = Regex.Match(
                response,
                "\"value\"\\s*:\\s*(-?\\d+(?:[.,]\\d+)?)",
                RegexOptions.IgnoreCase);

            if (!match.Success)
                throw new InvalidOperationException(
                    "SCOPE cevabında 'value' alanı bulunamadı: " + response);

            string raw = match.Groups[1].Value.Replace(',', '.');
            double number;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                throw new InvalidOperationException(
                    "Okunan değer sayı olarak çözümlenemedi: " + raw);

            return number.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
