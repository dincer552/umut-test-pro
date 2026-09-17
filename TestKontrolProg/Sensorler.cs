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
                Text = "Verileri Çek",
                Size = new Size(100, 30),
                Location = new System.Drawing.Point(282, 360),
                UseVisualStyleBackColor = true
            };
            _c600ReadButton.Click += C600ReadButton_Click;
            Controls.Add(_c600ReadButton);
        }

        private async void C600ReadButton_Click(object sender, EventArgs e)
        {
            if (_c600ReadButton != null)
                _c600ReadButton.Enabled = false;

            try
            {
                string value = await C600Communication.ReadValueAsync("SUPPLY_AIR_TEMP");
                textBox6.Text = value;
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "C600 / SCOPE üzerinden Supply Air sıcaklığı okunamadı.\n\n" + ex.Message,
                    "C600 Haberleşme",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
            finally
            {
                if (_c600ReadButton != null)
                    _c600ReadButton.Enabled = true;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }
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
            // TextBox'ları Tag (key adı) sırasına göre al – ister alfabetik ister başka kritere göre
            var textBoxes = groupBox.Controls
                                    .OfType<TextBox>()
                                    .OrderBy(tb => tb.Tag?.ToString())
                                    .ToArray();

            foreach (TextBox tb in textBoxes)
            {
                string key = tb.Tag?.ToString();
                if (string.IsNullOrEmpty(key)) continue;

                // Boşsa "-", doluysa kendi string değeri
                string value1 = string.IsNullOrWhiteSpace(tb.Text) ? "-" : tb.Text.Trim();

                if (value1 != "-" && !double.TryParse(value1, out _))
                {
                    value1 = "-";
                }

                if (sensorDict.ContainsKey(key))
                {
                    sensorDict[key] = value1;
                }
                else
                {
                    sensorDict.Add(key, value1);
                }
            }
        }
        private void Sensorler_Load(object sender, EventArgs e)
        {

        }
    }

    /// <summary>
    /// Siemens Climatix SCOPE local JSON tunnel üzerinden C600 okuma altyapısı.
    /// İlk aşamada yalnızca JSON Read kullanılır; PLC'ye yazma yapılmaz.
    /// </summary>
    internal static class C600Communication
    {
        private const string ScopeBaseUrl = "http://127.0.0.1:4242";

        public static async Task<string> ReadValueAsync(string jsonId)
        {
            if (string.IsNullOrWhiteSpace(jsonId))
                throw new ArgumentException("JSON ID boş olamaz.", "jsonId");

            string url = ScopeBaseUrl + "/json.html?callback=?&fn=Read&id=" + Uri.EscapeDataString(jsonId);

            using (var client = new WebClient())
            {
                client.Encoding = Encoding.UTF8;
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
                "\\\"value\\\"\\s*:\\s*(-?\\d+(?:[.,]\\d+)?)",
                RegexOptions.IgnoreCase);

            if (!match.Success)
                throw new InvalidOperationException("SCOPE cevabında 'value' alanı bulunamadı: " + response);

            string raw = match.Groups[1].Value.Replace(',', '.');
            double number;
            if (!double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                throw new InvalidOperationException("Okunan değer sayı olarak çözümlenemedi: " + raw);

            return number.ToString("0.##", CultureInfo.InvariantCulture);
        }
    }
}
