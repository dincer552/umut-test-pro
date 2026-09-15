using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestKontrolProg
{
    public partial class DamperKontrol : Form
    {
        private Form1 _form1;
        public DamperKontrol(Form1 form)
        {
            InitializeComponent();
            _form1 = form;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            TextBox[] tboxes = { textBox1, textBox2, textBox3, textBox4, textBox5, textBox6 };
           // RadioButton[] Rbuttons = { radioButton2, radioButton4, radioButton6, radioButton8, radioButton10, radioButton12 };
            
        
            //Damper sayıları
            for (int i = 0; i < tboxes.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(tboxes[i].Text) && int.TryParse(tboxes[i].Text, out int number)) // boş kontrolü
                {
                    Datas.DamperDatas[i, 0] = int.Parse(tboxes[i].Text);
                }
                else
                {
                    Datas.DamperDatas[i, 0] = 0; // boş bırakılırsa 0 yaz
                }
            }
            //Damper sayıları


            //Oransal Kontrol
            //for (int i = 0; i < Rbuttons.Length; i++)
            //{
            //    if (Rbuttons[i].Checked)   // RadioButton seçili mi?
            //    {
            //        Datas.DamperDatas[i, 1] = 1;
            //    }
            //    else
            //    {
            //        Datas.DamperDatas[i, 1] = 0;
            //    }
            //}
            //Oransal Kontrol


            _form1.SetLabelDamperKontrolText();
            this.Hide();
        }


        private void radioButton6_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void DamperKontrol_Load(object sender, EventArgs e)
        {

        }
    }
}
