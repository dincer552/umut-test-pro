using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Drawing;

namespace TestKontrolProg
{
    public partial class Sensorler : Form
    {
        private Form1 _form1;
        public Sensorler(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
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
                if (string.IsNullOrEmpty(key) ) continue;   // Tag yoksa atla

                // Boşsa "-", doluysa kendi string değeri
                string value1 = string.IsNullOrWhiteSpace(tb.Text) ? "-" : tb.Text.Trim();

                if (value1 != "-" && !double.TryParse(value1,out _))
                {
                    value1 = "-";
                }


                if (sensorDict.ContainsKey(key))
                {
                    sensorDict[key] = value1;   // Value’yu güncelle
                }
                else
                {
                    // İsterseniz olmayan key için ekleme de yapabilirsiniz
                    sensorDict.Add(key, value1);
                }
            }
        }
        private void Sensorler_Load(object sender, EventArgs e)
        {

        }
    }
}
