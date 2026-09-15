using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.AxHost;

namespace TestKontrolProg
{
    public partial class Moduller : Form
    {
        private Form1 _form1;
        public Moduller(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //Rotor
            if (checkBox1.Checked)
            {
                Datas.RotorData[0, 0] = 1;
                if (radioButton1.Checked) 
                {
                    Datas.RotorData[0, 0] = 1;
                }
                else
                {
                    Datas.RotorData[0, 1] = 0;
                }
            }
            else
            {
                Datas.RotorData[0, 0] = 0;
                Datas.RotorData[0, 1] = 0;
            }
            //Rotor


            //RunAround
            if (checkBox2.Checked)
            {
                Datas.RunAroundData = 1;
            }
            else
            {
                Datas.RunAroundData = 0;
            }
            //RunAround


            //RoomBMS
            if (checkBox7.Checked)
            {
                Datas.RoomBMS = 1;
            }
            else
            {
                Datas.RoomBMS = 0;
            }
            //RoomBMS


            //TempAvgEn
            if (checkBox8.Checked)
            {
                Datas.TempAvgEn = 1;
            }
            else
            {
                Datas.TempAvgEn = 0;
            }
            //TempAvgEn


            //ChangeOver 
            if (checkBox6.Checked)
            {
                Datas.ChangeOverData = 1;
            }
            else
            {
                Datas.ChangeOverData = 0;
            }
            //ChangeOver 


            //DX Batarya
            if (checkBox3.Checked)
            {
                Datas.DXData[0, 0] = 1;
                int value;
                if (!int.TryParse(textBox1.Text, out value))
                {
                    value = 0; // Sayı değilse 0 ata
                }

                if (value < 0) { value = 1; }else if (value > 5) {  value = 5; }
                Datas.DXData[0, 1] = value;

            }
            else
            {
                Datas.DXData[0, 0] = 0;
                Datas.DXData[0, 1] = 0;
            }
            //DX Batarya


            //Nemlendirici
            if (checkBox4.Checked)
            {
                Datas.NemlendiriciData[0, 0] = 1;
                int value;
                if (!int.TryParse(textBox2.Text, out value))
                {
                    value = 0; // Sayı değilse 0 ata
                }
                if (value < 0) { value = 1; } else if (value > 8) { value = 8; }
                Datas.NemlendiriciData[0, 1] = value;
            }
            else
            {
                Datas.NemlendiriciData[0, 0] = 0;
                Datas.NemlendiriciData[0, 1] = 0;
            }
            //Nemlendirici


            //Vanalar
            for (int i = 0; i < 4; i++)
            {
                Datas.ValveData[i] = checkedListBox1.GetItemChecked(i) ? 1 : 0;
            }
            //Vanalar


            //Komponentler
            for (int i = 0; i < 7; i++)
            {
                Datas.Components[i] = checkedListBox2.GetItemChecked(i) ? 1 : 0;
            }
            //Komponentler

           
            //Elektrikli ısıtıcı
            if (checkBox5.Checked)
            {
                Datas.ElectricalHeater = 1;
            }
            else
            {
                Datas.ElectricalHeater = 0;
            }
            //Elektrikli ısıtıcı
          

            //Elektrikli Data
            int index = 0;
            for (int i = 0; i < 3; i++) // satırlar
            {
                for (int j = 0; j < 3; j++) // sütunlar
                {
                    var value = dataGridView1.Rows[i].Cells[j].Value;

                    float floatValue = 0;
                    if (value != null && float.TryParse(value.ToString(), out floatValue))
                    {
                        Datas.ElectricalData[index] = floatValue;
                    }
                    else
                    {
                        Datas.ElectricalData[index] = 0; // boş veya geçersizse 0
                    }
                    index++; // tek boyutlu dizide sıradaki konuma geç
                }
            }
            //Elektrikli Data


            _form1.SetLabelModulKontrolText();
            this.Hide();
        }

        private void Moduller_Load(object sender, EventArgs e)
        {
            
            string[] HTTKademeName = { "1. Kademe", "2. Kademe", "3. Kademe" };
            string[] HTTColumnsName = { "R(A)", "S(A)", "T(A)" };

            helperFuncs.HazirlaTablo(dataGridView1, HTTColumnsName, HTTKademeName, 70, 100);
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkBox5_CheckedChanged(object sender, EventArgs e)
        {
            
        }
    }
}
