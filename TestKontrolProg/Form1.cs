using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Net.Mime.MediaTypeNames;

using Microsoft.VisualBasic;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using static System.Windows.Forms.AxHost;
using System.Diagnostics;
using com.itextpdf.text.pdf;

using ClosedXML.Excel;
using System.Drawing.Text;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;
using Microsoft.Office.Core;


namespace TestKontrolProg
{
    public partial class Form1 : Form
    {
        private FanKontrol FanKontrolForm;
        private DamperKontrol DamperKontrolForm;
        private FiltreKontrol FiltreKontrolForm;
        private Moduller ModullerForm;
        private Sensorler SensorlerForm;
        private UserControl UserControlForm;

        public Form1()
        {
            InitializeComponent();
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Icon = Properties.Resources.logo;   // Kaynaktan Icon tipi olarak al

            var items = Properties.Settings.Default.SavedText;
            if (items != null)
            {
                textBox1.Text = items[0];
                textBox2.Text = items[1];
                textBox3.Text = items[2];
            }
            Datas.UserName = Properties.Settings.Default.UserName;
            string base64 = Properties.Settings.Default.DefaultImageBase64;
            Datas.imgBytes = Convert.FromBase64String(base64);



        }
        private void Form1_Load(object sender, EventArgs e)
        {
            Datas.UserName = Properties.Settings.Default.UserName;
            string base64 = Properties.Settings.Default.DefaultImageBase64;
            Datas.imgBytes = Convert.FromBase64String(base64);
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }
        private void label2_Click(object sender, EventArgs e)
        {

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            var itemsToSave = new System.Collections.Specialized.StringCollection();




            // Önce TextBox'lardaki verileri ekle
            string[] savedTexts = new string[]
            {
                textBox1.Text,
                textBox2.Text,
                textBox3.Text,
            };

            foreach (var text in savedTexts)
            {
                itemsToSave.Add(text);
            }

            // Ayara kaydet
            Properties.Settings.Default.SavedText = itemsToSave;
            if (UserControlForm != null)
            {
                Properties.Settings.Default.UserName = UserControlForm.label1.Text;
            }
            
            Properties.Settings.Default.DefaultImageBase64 = Convert.ToBase64String(Datas.imgBytes);
            Properties.Settings.Default.Save();

        }
        private void button2_Click(object sender, EventArgs e)
        {
            if (FanKontrolForm == null || FanKontrolForm.IsDisposed)
                FanKontrolForm = new FanKontrol(this);
            FanKontrolForm.Show();
            FanKontrolForm.BringToFront();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            if (DamperKontrolForm == null || DamperKontrolForm.IsDisposed)
                DamperKontrolForm = new DamperKontrol(this);
            DamperKontrolForm.Show();
            DamperKontrolForm.BringToFront();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (FiltreKontrolForm == null || FiltreKontrolForm.IsDisposed)
                FiltreKontrolForm = new FiltreKontrol(this);
            FiltreKontrolForm.Show();
            FiltreKontrolForm.BringToFront();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            if (ModullerForm == null || ModullerForm.IsDisposed)
                ModullerForm = new Moduller(this);
            ModullerForm.Show();
            ModullerForm.BringToFront();
        }
        private void button6_Click(object sender, EventArgs e)
        {
            if (SensorlerForm == null || SensorlerForm.IsDisposed)
                SensorlerForm = new Sensorler(this);
            SensorlerForm.Show();
            SensorlerForm.BringToFront();
        }
        private void label16_Click(object sender, EventArgs e)
        {

        }
        public void SetLabelFanKontrolText(string text = "Kontrol Edildi")
        {
            label12.Text = text;
            label12.BackColor = System.Drawing.Color.Chartreuse;
        }
        public void SetLabelDamperKontrolText(string text = "Kontrol Edildi")
        {
            label13.Text = text;
            label13.BackColor = System.Drawing.Color.Chartreuse;
        }
        public void SetLabelModulKontrolText(string text = "Kontrol Edildi")
        {
            label14.Text = text;
            label14.BackColor = System.Drawing.Color.Chartreuse;
        }
        public void SetLabelFiltreKontrolText(string text = "Kontrol Edildi")
        {
            label15.Text = text;
            label15.BackColor = System.Drawing.Color.Chartreuse;
        }
        public void SetLabelSensorKontrolText(string text = "Kontrol Edildi")
        {
            label16.Text = text;
            label16.BackColor = System.Drawing.Color.Chartreuse;
        }


        private void button8_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Boş Buton");
            Datas.ProjectInfo[0] = textBox1.Text;
            Datas.ProjectInfo[1] = textBox2.Text;
            Datas.ProjectInfo[2] = textBox3.Text;
            Datas.Not.Clear(); // varsa eskileri temizle
            Datas.Not.AddRange(textBox19.Lines);
            helperFuncs.ExcelOpen();
            helperFuncs.ExcelWrite();
            helperFuncs.ExcelSave();
        }


        private void button1_Click_1(object sender, EventArgs e)
        {
            string klasor = AppDomain.CurrentDomain.BaseDirectory;
            Datas.pdfPath = System.IO.Path.Combine(klasor, "Test_Kontrol_Listesi.pdf");
            string dosyaAdi = textBox3.Text + "_Test_Kontrol_Listesi.pdf";
            string input = Interaction.InputBox("Dosya adını giriniz:", "", dosyaAdi);
            Datas.newPdfPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), input);
            Datas.ProjectInfo[0] = textBox1.Text;
            Datas.ProjectInfo[1] = textBox2.Text;
            Datas.ProjectInfo[2] = textBox3.Text;
            Datas.Not.Clear(); // varsa eskileri temizle
            Datas.Not.AddRange(textBox19.Lines);
            helperFuncs2.TestControlReport();
        }

        private void button7_Click(object sender, EventArgs e)
        {
             if (UserControlForm == null || UserControlForm.IsDisposed)
                UserControlForm = new UserControl(this);
            UserControlForm.Show();
            UserControlForm.BringToFront();
        }
    }
}
