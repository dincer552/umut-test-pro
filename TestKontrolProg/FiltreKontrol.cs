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
    public partial class FiltreKontrol : Form
    {
        private Form1 _form1;
        public FiltreKontrol(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 18; i++)
            {
                Datas.Filtreler[i, 0] = (checkedListBox1.GetItemChecked(i) ? 1 : 0).ToString();
                Datas.Filtreler[i, 1] = checkedListBox1.Items[i].ToString();
            }
            _form1.SetLabelFiltreKontrolText();
            this.Hide();
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        private void FiltreKontrol_Load(object sender, EventArgs e)
        {

        }
    }
}
