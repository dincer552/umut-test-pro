using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Microsoft.Office.Core;
using Microsoft.VisualBasic;
using static System.Windows.Forms.AxHost;



namespace TestKontrolProg
{
    public static class helperFuncs2
    {
        public static void PDFWrite(PdfContentByte cb, string[] metinDizisi, float x, float yBaslangic, float satirAraligi)
        {
            for (int i = 0; i < metinDizisi.Length; i++)
            {
                cb.SetTextMatrix(x, yBaslangic - i * satirAraligi);
                cb.ShowText(metinDizisi[i]);
            }
        }
        public static void SensorChecked(PdfContentByte cb, Dictionary<string, string> dict,
         int checkRow, float checkCol)
        {
            if (dict == null)
            {
                cb.SetTextMatrix(checkRow, checkCol);
                cb.ShowText("-");
                return;
            }
            

            bool checkedWritten = false; // "Checked" yazıldı mı

            foreach (var kvp in dict)
            {
                string value = kvp.Value;
                // Değer geçerli mi?
                if (!string.IsNullOrEmpty(value) && value != "-" && value != "0")
                {
                    if (!checkedWritten)
                    {
                        cb.SetTextMatrix(checkRow, checkCol);
                        cb.ShowText("Checked");
                        checkedWritten = true;
                    }
                    // İsterseniz burada key veya value ile başka işlem de yapabilirsiniz
                }
                else
                {
                    cb.SetTextMatrix(checkRow, checkCol);
                    cb.ShowText("-");
                }
            }
        }
        public static void DamperWrite(PdfContentByte cb, float x, float y, int[,] sensor, int index)
        {
            cb.SetTextMatrix(x, y);
            if (sensor[index, 0] > 0)
            {
                cb.ShowText("Checked");  
            }
            else
            {
                cb.ShowText("-");
            }
        }
        public static void TableDraw(PdfContentByte cb,int x, int y, int w, int h, int satir, int sutun)
        {
            cb.SetLineWidth(0.9f);
            cb.SetColorStroke(BaseColor.BLACK);

            int wTemp = w;
            for (int i = 0; i < sutun; i++) 
            {
                if (i == 0) { w = w + 25;} else { w = wTemp; }
                    for (int j = 0; j < satir; j++)
                    {
                        cb.Rectangle(x + (i * w), y + (j * h), w, h);
                    }
                if (i == 0) { x = x + 25; }
            }
            cb.Stroke();
        }
        public static void Imza(PdfContentByte cb)
        {


            byte[] imgBytes = Datas.imgBytes;
            Image img = Image.GetInstance(imgBytes);

            img.SetAbsolutePosition(455, 120);
            img.ScaleToFit(100, 200);
            cb.AddImage(img);
            cb.Rectangle(455, 120, 100, 100);
            cb.Stroke();
            cb.SetTextMatrix(470, 200);
            cb.ShowTextAligned(Element.ALIGN_CENTER,Datas.UserName,505,200,0);
        }
        public static void LineDraw(PdfContentByte cb, int Start_x, int Start_y, int Finish_x, int Finis_y)
        {
            cb.SetLineWidth(0.8f);
            cb.SetColorStroke(BaseColor.BLACK);

            cb.MoveTo(Start_x, Start_y);
            cb.LineTo(Finish_x, Finis_y);

            cb.Stroke();
        }
        public static int SensorValueWrite2(PdfContentByte cb, Dictionary<string, string> dict,
            string sensorName, int detailStartRow, int detailCol)
        {
            if (dict == null || dict.Values.All(v => v == "-")) return detailStartRow;

            // --- Eğer dict'teki tüm değerler null veya boş ise tabloyu yazma ---
            bool hasValue = dict.Values.Any(v => !string.IsNullOrWhiteSpace(v));
            if (!hasValue) return detailStartRow;


           
            int startRow = detailStartRow;
            LineDraw(cb, 35, startRow+12, 325, startRow+12);
            // --- Başlık satırı ---
            cb.SetTextMatrix(detailCol,startRow);
            cb.ShowText("Sensor");
            
         
            int col = detailCol + 90;
            foreach (var key in dict.Keys)
            {
                if (dict[key] != "-")
                {
                    cb.SetTextMatrix(col, startRow);
                    cb.ShowText(key);
                   
                    cb.SetTextMatrix(col, startRow - 11);
                    cb.ShowText(dict[key]);
                  
                    col += 65;
                }
            }

            // --- Sensör adı + değerler --- //
            cb.SetTextMatrix(detailCol,startRow - 11);
            cb.ShowText(sensorName);
          
            col = detailCol + 50;
            return startRow - 30; // +3: 2 satır tablo + 1 satır boşluk
        }


        public static void TestControlReport()
        {
            try
            {
                using (PdfReader reader = new PdfReader(Datas.pdfPath))
                using (FileStream fs = new FileStream(Datas.newPdfPath, FileMode.Create, FileAccess.Write))
                using (PdfStamper stamper = new PdfStamper(reader, fs))
                {
                    PdfContentByte cb = stamper.GetOverContent(1); // 1. sayfa

                    // türkçe font
                    string fontPath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "timesbd.ttf");
                    BaseFont bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);

                    //Font ve size ayarla
                    cb.BeginText();
                    cb.SetFontAndSize(bf, 9);
                    //Font ve size ayarla

                    int x = 430;
                    float y = 695.5f;
                    // Tarih
                    cb.SetTextMatrix(x, y);
                    cb.ShowText(": " + DateTime.Now.ToString());
                    // Tarih

                    //Santral bilgisi yaz
                    helperFuncs2.PDFWrite(cb, Datas.ProjectInfo, 180, 677, 23.5f);
                    //Santral bilgisi yaz
                    x = 180 ;y = 583;
                    // Fan Kontrol yaz
                    cb.SetTextMatrix(x, y);
                    if (Datas.FanTip == 1)
                    {
                        cb.ShowText("Danfoss Ziehl-Abegg");
                    }
                    else if (Datas.FanTip == 2)
                    {
                        cb.ShowText("EC Ziehl-Abegg");
                    }
                    else if (Datas.FanTip == 3)
                    {
                        cb.ShowText("EC EBM-Papst");
                    }
                    else
                    {
                        cb.ShowText("-");
                    }

                    cb.SetTextMatrix(x, y-23.5f);
                    if (Datas.AirFlowControlOk == true)
                    {
                        
                        cb.ShowText("Checked");
                    }
                    else
                    {
                        cb.ShowText("-");
                    }

                    cb.SetTextMatrix(x, y-47);
                    if (Datas.PressureControlOk == true)
                    { 
                        cb.ShowText("Checked");
                    }
                    else
                    {
                        cb.ShowText("-");
                    }
                    x = 443; y = 583;
                    cb.SetTextMatrix(x, y);
                    if (Datas.SupplyFanNumber != 0)
                    {
                        cb.ShowText(Datas.SupplyFanNumber.ToString());
                    }
                    else
                    {
                        cb.ShowText("-");
                    }

                    cb.SetTextMatrix(x, y-23.5f);
                    if (Datas.ReturnFanNumber != 0)
                    {     
                        cb.ShowText(Datas.ReturnFanNumber.ToString());
                    }
                    else
                    {
                        cb.ShowText("-");
                    }

                    cb.SetTextMatrix(x, y-47);
                    if (Datas.SupplyAirFlow != 0)
                    {
                        cb.ShowText(Datas.SupplyAirFlow.ToString());
                    }
                    else
                    {
                        cb.ShowText("-");
                    }

                    cb.SetTextMatrix(x, y-70.5f);
                    if (Datas.ReturnAirFlow != 0)
                    {
                        cb.ShowText(Datas.ReturnAirFlow.ToString());
                    }
                    else
                    {
                        cb.ShowText("-");
                    }
                    // Fan Kontrol yaz


                    // Rotor yaz      
                    if (Datas.RotorData[0, 0] == 1)
                    {
                        Datas.Modules1[0] = "Checked";
                    }
                    else
                    {
                        Datas.Modules1[0] = "-";
                    }
                    // Rotor yaz

                    // RunAround ve ChangeOver yaz 
                    if (Datas.RunAroundData == 1)
                    {
                        Datas.Modules1[1] = "Checked";
                    }
                    else
                    {
                        Datas.Modules1[1] = "-";
                    }

                    if (Datas.ChangeOverData == 1)
                    {
                        Datas.Modules1[2] = "Checked";
                    }
                    else
                    {
                        Datas.Modules1[2] = "-";
                    }
                    // RunAround ve ChangeOver yaz

                    // DX ve Nemlendirici yaz
                    if (Datas.DXData[0, 0] == 1)
                    {
                        Datas.Modules1[4] = "Checked";
                        Datas.Modules1[5] = Datas.DXData[0, 1].ToString();
                    }
                    else
                    {
                        Datas.Modules1[4] = "-";
                        Datas.Modules1[5] = "-";
                    }

                    if (Datas.NemlendiriciData[0, 0] == 1)
                    {
                        Datas.Modules1[6] = "Checked";
                        Datas.Modules1[7] =  Datas.NemlendiriciData[0, 1].ToString();
                    }
                    else
                    {
                        Datas.Modules1[6] = "-";
                        Datas.Modules1[7] = "-";
                    }
                        // DX ve Nemlendirici yaz

                        // Vana yaz
                        for (int i = 0; i < 4; i++)
                        {
                            if (i < 2)
                            {
                                if (Datas.ValveData[i] == 1)
                                {
                                    Datas.Modules1[i + 8] = "Checked";
                                }
                                else
                                {
                                    Datas.Modules1[i + 8] = "-";
                                }
                            }
                            else
                            {
                                if (Datas.ValveData[i] == 1)
                                {
                                    Datas.Modules2[i - 2] = "Checked";
                                }
                                else
                                {
                                    Datas.Modules2[i - 2] = "-";
                                }
                            }
                        }
                    // Vana yaz

                    //elektrikli varsa   
                    if (Datas.ElectricalHeater == 1)
                    {
                        Datas.Modules1[3] = "Checked";
                    }
                    else
                    {
                        Datas.Modules1[3] = "-";
                    }
                        //elektrikli varsa


                    //Modules 1 yaz
                    helperFuncs2.PDFWrite(cb, Datas.Modules1, 180, 460, 26);
                    //Modules 1 yaz


                    // Komponent yaz
                    for (int i = 0; i < 7; i++)
                    {
                        if (Datas.Components[i] == 1)
                        {
                            Datas.Modules2[i+2] = "Checked";
                        }
                        else 
                        {
                            Datas.Modules2[i + 2] = "-";
                        }
                    }
                    helperFuncs2.PDFWrite(cb, Datas.Modules2, 443, 460, 26);
                    // Komponent yaz

                    // Damper data yaz
                    x = 180;y = 172.5f;
                    helperFuncs2.DamperWrite(cb, x, y, Datas.DamperDatas, 0);
                    helperFuncs2.DamperWrite(cb, x, y-25.5f, Datas.DamperDatas, 1);
                    helperFuncs2.DamperWrite(cb, x, y-51, Datas.DamperDatas, 2);
                    x = 443; y = 172.5f;
                    helperFuncs2.DamperWrite(cb, x, y, Datas.DamperDatas, 3);
                    helperFuncs2.DamperWrite(cb, x, y-25.5f, Datas.DamperDatas, 4);
                    helperFuncs2.DamperWrite(cb, x, y-51, Datas.DamperDatas, 5);
                    // Damper data yaz

                    //Not yaz
                    int row = 0;
                    foreach (string line in Datas.Not)
                    {
                        //if (row < 4)
                        //{
                        cb.SetTextMatrix(300, 76 - (row * 12));
                        cb.ShowText(line);
                        row++;
                        //}
                    }
                    //Not yaz

                    //Filtre yaz
                    int satir = 0;
                    int sutun = 0;
                    int filterCount = 0;
                    for (int i = 0; i < 18; i++)
                    {
                        if (Datas.Filtreler[i, 0] == "1")
                        {
                            filterCount++;
                            if (filterCount < 9)
                            {
                                cb.SetTextMatrix(38 + satir, 76 - sutun);
                                cb.ShowText(Datas.Filtreler[i, 1] + ": Checked");
                                sutun += 12;
                                if (sutun % 48 == 0)
                                {
                                    sutun = 0;
                                    satir += 130;
                                }
                            }
                        }
                    }
                    //Filtre yaz



                    // 2.sayfaya geç
                    cb = stamper.GetOverContent(2); // 2.sayfa
                    cb.SetFontAndSize(bf, 9);
                    // 2.sayfaya geç

                    // Tarih
                    cb.SetTextMatrix(430, 696.5f);
                    cb.ShowText(": " + DateTime.Now.ToString());
                    // Tarih

                    // sensor kontrol et
                    x = 180; y = 680;
                    helperFuncs2.SensorChecked(cb, Datas.FreshAirSensorValues, x, y);
                    helperFuncs2.SensorChecked(cb, Datas.SupplyAirSensorValues, x, y-23.5f);
                    helperFuncs2.SensorChecked(cb, Datas.ExhaustAirSensorValues, x, y - (23.5f * 2));
                    helperFuncs2.SensorChecked(cb, Datas.ReturnAirSensorValues, x, y - (23.5f * 3));
                    helperFuncs2.SensorChecked(cb, Datas.AfterCoilAirSensorValues, x, y - (23.5f * 4));
                    helperFuncs2.SensorChecked(cb, Datas.MixAirSensorValues, x, y - (23.5f * 5));
                    x = 443; y = 680;
                    helperFuncs2.SensorChecked(cb, Datas.RoomTempSensor1Values, x, y);
                    helperFuncs2.SensorChecked(cb, Datas.RoomTempSensor2Values, x, y- 23.5f);
                    helperFuncs2.SensorChecked(cb, Datas.ReturnCO2SensorValues, x, y - (23.5f * 2));
                    helperFuncs2.SensorChecked(cb, Datas.WaterTempSensorValues, x, y - (23.5f * 3));
                    helperFuncs2.SensorChecked(cb, Datas.ReturnCO2AirSensorValues, x, y - (23.5f * 4));
                    // sensor kontrol et

                    // elektrikli var mı
                    if (Datas.ElectricalHeater == 1)
                    {
                        helperFuncs2.TableDraw(cb, 146, 450, 70, 20, 4, 4);

                        cb.SetTextMatrix(155, 517);
                        cb.ShowText("Elektrical Heater");

                        cb.SetTextMatrix(155, 497);
                        cb.ShowText("1.Kademe");

                        cb.SetTextMatrix(155, 477);
                        cb.ShowText("2.Kademe");

                        cb.SetTextMatrix(155, 457);
                        cb.ShowText("3.Kademe");

                        cb.SetTextMatrix(265, 517);
                        cb.ShowText("R(A)");

                        cb.SetTextMatrix(335, 517);
                        cb.ShowText("S(A)");

                        cb.SetTextMatrix(405, 517);
                        cb.ShowText("T(A)");

                        int sayac = 0;
                        for (int i = 0; i < 3; i++)
                        {
                            for (int j = 0; j < 3; j++)
                            {
                                cb.SetTextMatrix(265 + (j * 70), 497 - (i * 20));
                                cb.ShowText(Datas.ElectricalData[sayac].ToString());
                                sayac += 1;
                            }
                        }
                    }
                    // elektrikli var mı


                    // sensor değerlerini yaz
                    int detailRow = 510;
                    if (Datas.ElectricalHeater == 1) { detailRow = 415; }
                    int detailCol = 50;
                    cb.SetFontAndSize(bf, 8);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.FreshAirSensorValues, "FreshAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.SupplyAirSensorValues, "SupplyAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.ReturnAirSensorValues, "ReturnAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.ExhaustAirSensorValues, "ExhaustAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.AfterCoilAirSensorValues, "AfterCoilAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.MixAirSensorValues, "MixCoilAirSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.RoomTempSensor1Values, "RoomTempSensor1", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.RoomTempSensor2Values, "RoomTempSensor2", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.ReturnCO2SensorValues, "ReturnCO2Sensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.WaterTempSensorValues, "WaterTempSensor", detailRow, detailCol);
                    detailRow = helperFuncs2.SensorValueWrite2(cb, Datas.ReturnCO2AirSensorValues, "ReturnCO2AirSensor", detailRow, detailCol);

                    // sensor değerlerini yaz
                    cb.SetFontAndSize(bf, 10);
                    Imza(cb);

                    cb.EndText();
                }
                MessageBox.Show("PDF üzerine yazıldı:\n" + Datas.newPdfPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Hata: " + ex.Message + "\n\nDetay:\n" + ex.StackTrace);
            }

        }
    }
}      
 

