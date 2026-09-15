using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;


using System.IO;
using Microsoft.VisualBasic;
using System.Windows.Media;
using Org.BouncyCastle.Asn1;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Security.Cryptography;

namespace TestKontrolProg
{
    public partial class UserControl : Form
    {
        private Form1 _form1;
        string userFolderPath;
        public UserControl(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;

            userFolderPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "user");
            Directory.CreateDirectory(userFolderPath);

        }

        private void UserControl_Load(object sender, EventArgs e)
        {
            string base64 = Properties.Settings.Default.DefaultImageBase64;
            label1.Text = Properties.Settings.Default.UserName;
        
            if (!string.IsNullOrEmpty(base64))
            {
                byte[] imgBytes = Convert.FromBase64String(base64);
                using (var ms = new MemoryStream(imgBytes))
                {
                    pictureBox1.Image = System.Drawing.Image.FromStream(ms);
                }
            }
            else
            {
                string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "null.png");
                if (File.Exists(filePath))
                {
                    pictureBox1.Image = System.Drawing.Image.FromFile(filePath);
                }
            }
            LoadImageList();
        }
        private void LoadImageList()
        {
            listBox1.Items.Clear();

            string[] files = Directory.GetFiles(userFolderPath, "*.*").Where(
                f=>f.EndsWith(".png",StringComparison.OrdinalIgnoreCase)).ToArray();
            string file1 = "";
            foreach (var file in files) 
            {
                if (file.Length > 4)
                {
                    file1 = file.Substring(0,file.Length- 4);
                   
                }
                else
                {
                    file1 = file;
                }
                    listBox1.Items.Add(Path.GetFileName(file1));
            }
        }

      
        
        private void button2_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Resim Dosyaları|*.png";

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                byte[] imgBytes = File.ReadAllBytes(ofd.FileName);
                string base64 = Convert.ToBase64String(imgBytes);

                Properties.Settings.Default.DefaultImageBase64 = base64;
                Properties.Settings.Default.Save();
                
                using (var ms = new MemoryStream(imgBytes))
                {
                    pictureBox1.Image = System.Drawing.Image.FromStream(ms);
                }
                string ImageName = Interaction.InputBox("Ad Soyad: ");
                Properties.Settings.Default.UserName = ImageName;
                label1.Text = ImageName;
                Datas.UserName = ImageName;
                Datas.imgBytes = imgBytes;
                iTextSharp.text.Image img = iTextSharp.text.Image.GetInstance(imgBytes);
                string debugKlasoru = Application.StartupPath;
                string userKlasoru = Path.Combine(debugKlasoru, "User");
                if (!Directory.Exists(userKlasoru))
                {
                    Directory.CreateDirectory(userKlasoru);
                }
                string dosyaYolu = Path.Combine(userKlasoru, ImageName + ".png");
                File.WriteAllBytes(dosyaYolu, imgBytes);
                LoadImageList();
                MessageBox.Show("User Kaydedildi");
            }

           
        }

        private void button1_Click(object sender, EventArgs e)
        {


            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Resim Dosyaları|*.png";

            if(ofd.ShowDialog()==DialogResult.OK)
            {
                byte[] imgBytes = File.ReadAllBytes(ofd.FileName);
                string base64 = Convert.ToBase64String(imgBytes);

                Properties.Settings.Default.DefaultImageBase64 = base64;
                Properties.Settings.Default.Save();

                using (var ms = new MemoryStream(imgBytes))
                {
                    pictureBox1.Image=System.Drawing.Image.FromStream(ms);
                }
              
            }


            
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(listBox1.SelectedItem != null)
            {
                string selectedItem = listBox1.SelectedItem.ToString();
                label1.Text = listBox1.SelectedItem.ToString();
                Datas.UserName = selectedItem;
                
                selectedItem = selectedItem + ".png";


                Datas.imgBytes = File.ReadAllBytes((Path.Combine(Application.StartupPath+"/user", selectedItem)));
                string fullPath = Path.Combine(userFolderPath,selectedItem);
                if (File.Exists(fullPath))
                {
                    using (var ms=new MemoryStream(File.ReadAllBytes(fullPath)))
                    {
                        pictureBox1.Image = System.Drawing.Image.FromStream(ms);
                    }
                }
               
            }
        }

        private void UserControl_FormClosing(object sender, FormClosingEventArgs e)
        {
            string selectedItem = listBox1.SelectedItem.ToString();
            label1.Text = listBox1.SelectedItem.ToString();
            Datas.UserName = selectedItem;

            selectedItem = selectedItem + ".png";

            Datas.imgBytes = File.ReadAllBytes((Path.Combine(Application.StartupPath + "/user", selectedItem)));

            Properties.Settings.Default.DefaultImageBase64 = Convert.ToBase64String(Datas.imgBytes);
 
            Properties.Settings.Default.UserName = label1.Text;
            Properties.Settings.Default.Save();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            if(listBox1.SelectedItem!=null)
            {
                DialogResult sonuc = MessageBox.Show(
                    listBox1.SelectedItem.ToString()+" Sil...",
                    "User Sil",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                
                if (sonuc == DialogResult.Yes)
                {
                    string deleteImg = listBox1.SelectedItem.ToString() + ".png";
                    File.Delete(Path.Combine(Application.StartupPath + "/user", deleteImg));
                    Datas.imgBytes = File.ReadAllBytes(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "null.png")); 
                    Datas.UserName = "";
                    label1.Text = Datas.UserName;
                //    pictureBox1.Image = ""; 
                    LoadImageList();
                }
                
            }
            else
            {
                MessageBox.Show("Listeden seçiniz...");
            }
        }
    }
}
