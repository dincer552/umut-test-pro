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
    public partial class FanKontrol : Form
    {
        private Form1 _form1;
        public FanKontrol(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                Datas.FanTip = 1;
            else if (radioButton2.Checked)
                Datas.FanTip = 2;
            else if (radioButton3.Checked)
                Datas.FanTip = 3;


            int value;
            if (int.TryParse(textBox1.Text, out value))
            {
                Datas.SupplyFanNumber = value;
            }
            else
            {
                Datas.SupplyFanNumber = 0; // boş veya geçersiz giriş olursa varsayılan değer
            }

            if (int.TryParse(textBox2.Text, out value))
            {
                Datas.ReturnFanNumber = value;
            }
            else
            {
                Datas.ReturnFanNumber = 0; // boş veya geçersiz giriş olursa varsayılan değer
            }
       
            Datas.AirFlowControlOk = checkBox1.Checked;
            Datas.PressureControlOk = checkBox2.Checked;

            if (int.TryParse(textBox3.Text, out value))
            {
                Datas.SupplyAirFlow = value;
            }
            else
            {
                Datas.SupplyAirFlow = 0; // boş veya geçersiz giriş olursa varsayılan değer
            }
            if (int.TryParse(textBox4.Text, out value))
            {
                Datas.ReturnAirFlow = value;
            }
            else
            {
                Datas.ReturnAirFlow = 0; // boş veya geçersiz giriş olursa varsayılan değer
            }
            
            _form1.SetLabelFanKontrolText();
            this.Hide();
        }

        private void FanKontrol_Load(object sender, EventArgs e)
        {

        }
    }
}
