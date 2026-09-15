using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.PowerPoint;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;


namespace TestKontrolProg
{
    public static class helperFuncs
    {

        public static string debugPath = AppDomain.CurrentDomain.BaseDirectory;
        public static string filePath = Path.Combine(debugPath, "Test_Report_BACnet_Modbus.xlsx");
        public static string endFilePath = Path.Combine(debugPath, "Test_Report_BACnet_Modbus.xlsx");
        public static XLWorkbook workbook;
        public static string bosluk = " ";

        public static void HazirlaTablo(DataGridView dgv, string[] kolonIsimleri, string[] satirIsimleri, int kolonGenisligi = 70, int satirBaslikGenisligi = 100)
        {
            dgv.Columns.Clear();
            dgv.Rows.Clear();
            dgv.ColumnCount = kolonIsimleri.Length;
            for (int i = 0; i < kolonIsimleri.Length; i++)
            {
                dgv.Columns[i].Name = kolonIsimleri[i];
                dgv.Columns[i].Width = kolonGenisligi;
            }
            dgv.RowHeadersWidth = satirBaslikGenisligi;
            foreach (string satir in satirIsimleri)
            {
                int rowIndex = dgv.Rows.Add();
                dgv.Rows[rowIndex].HeaderCell.Value = satir;
            }
            dgv.ScrollBars = System.Windows.Forms.ScrollBars.None;
        }
        public static void ExcelOpen()
        {
            try
            {
                if (File.Exists(filePath))
                {
                    // Var olan dosyayı aç
                    workbook = new XLWorkbook(filePath);
                }
                else
                {
                    // Yeni dosya oluştur
                    workbook = new XLWorkbook();
                }
            }
            catch (IOException)
            {
                MessageBox.Show("Excel dosyası şu anda açık. Lütfen kapatıp tekrar deneyin.",
                                "Dosya Kullanımda", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // İşleme devam etme
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Beklenmedik bir hata oluştu: {ex.Message}",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public static void TestReportWrite(IXLWorksheet worksheet)
        {
            // "TestRaporu" sayfasını al ya da oluştur
            try
            {
                // Santral Info
                worksheet.Cell(10, 4).Value = bosluk + Datas.ProjectInfo[0];
                worksheet.Cell(11, 4).Value = bosluk + Datas.ProjectInfo[1];
                worksheet.Cell(12, 4).Value = bosluk + Datas.ProjectInfo[2];
                // Santral Info
                // **********************Fan Kontrol****************************
                // Fan Tip
                if (Datas.FanTip == 1)
                {
                   
                    worksheet.Cell(15, 4).Value =  "Danfoss Ziehl-Abegg";
                }
                else if (Datas.FanTip == 2)
                {
                    worksheet.Cell(15, 4).Value =  "EC Ziehl-Abegg";
                }
                else if (Datas.FanTip == 3)
                {
                    worksheet.Cell(15, 4).Value =  "EC EBM-Papst";
                }
                else
                {
                    worksheet.Cell(15, 4).Value = "-";
                }
                // Fan Tip
                // Fan Control Option
                if (Datas.AirFlowControlOk == true)
                {
                    worksheet.Cell(16, 4).Value =  "Checked";
                }
                else
                {
                    worksheet.Cell(16, 4).Value = "-";
                }
                if (Datas.PressureControlOk == true)
                {
                    worksheet.Cell(17, 4).Value = "Checked";
                }
                else
                {
                    worksheet.Cell(17, 4).Value = "-";
                }
                // Fan Control Option
                // Fan Number
                worksheet.Cell(15, 10).Value = Datas.SupplyFanNumber;
                worksheet.Cell(16, 10).Value = Datas.ReturnFanNumber;
                // Fan Number
                // Fan Speed
                worksheet.Cell(17, 10).Value =  Datas.SupplyAirFlow;
                worksheet.Cell(18, 10).Value =  Datas.ReturnAirFlow;
                // Fan Speed
                //********************** Fan Kontrol ****************************
                //********************** Moduller ****************************
                if (Datas.RotorData[0, 0] == 1)
                {
                    worksheet.Cell(21, 4).Value = "Checked";  
                }
                else
                {
                    worksheet.Cell(21, 4).Value = "-";
                }
                if (Datas.RunAroundData == 1)
                {
                    worksheet.Cell(22, 4).Value = "Checked";
                }
                else
                {
                    worksheet.Cell(22, 4).Value = "-";
                }
                if (Datas.DXData[0, 0] == 1)
                {
                    worksheet.Cell(25, 4).Value = "Checked";
                    worksheet.Cell(26, 4).Value = Datas.DXData[0, 1];
                }
                else
                {
                    worksheet.Cell(25, 4).Value = "-";
                    worksheet.Cell(26, 4).Value = "-";
                }
                if (Datas.NemlendiriciData[0, 0] == 1)
                {
                    worksheet.Cell(27, 4).Value = "Checked";
                    worksheet.Cell(28, 4).Value =  Datas.NemlendiriciData[0, 1];
                }
                else
                {
                    worksheet.Cell(27, 4).Value = "-";
                    worksheet.Cell(28, 4).Value = "-";
                }
                if (Datas.ElectricalHeater == 1)
                {
                    worksheet.Cell(24, 4).Value = "Checked";
                }
                else
                {
                    worksheet.Cell(24, 4).Value = "-";
                }

                for (int i = 0; i < 4; i++)
                {
                    if (i < 2)
                    {
                        if (Datas.ValveData[i] == 1)
                        {
                            worksheet.Cell(29 + i, 4).Value = "Checked";
                        }
                        else
                        {
                            worksheet.Cell(29 + i, 4).Value = "-";
                        }
                    }
                    else
                    {
                        if (Datas.ValveData[i] == 1)
                        {
                            worksheet.Cell(19 + i, 10).Value = "Checked";
                        }
                        else
                        {
                            worksheet.Cell(20 + i, 10).Value = "-";
                        }
                    }

                }

                if (Datas.ChangeOverData == 1)
                {
                    worksheet.Cell(23, 4).Value = "Checked";
                }
                else
                {
                    worksheet.Cell(23, 4).Value = "-";
                }

                for (int i = 0; i < 7; i++)
                {
                    if (Datas.Components[i] == 1)
                    {
                        worksheet.Cell(23 + i, 10).Value = "Checked";
                    }
                    else
                    {
                        worksheet.Cell(23 + i, 10).Value = "-";
                    }
                }
                //********************** Moduller ****************************
                //********************** Damper ****************************
                int satir = 0;
                int sutun = 0;
                for (int i = 0; i < 6; i++)
                {
                    if (Datas.DamperDatas[i, 0] != 0)
                    {
                        worksheet.Cell(33 + satir, 4 + sutun).Value = "Checked"; 
                    }
                    else
                    {
                        worksheet.Cell(33 + satir, 4 + sutun).Value = "-";
                    }
                    satir += 2;
                    if (satir % 6 == 0)
                    {
                        satir = 0; sutun += 6;
                    }
                }
                //********************** Damper ****************************

                //Not
                int row = 41; // Başlangıç satırı
                foreach (string line in Datas.Not)
                {
                    worksheet.Cell(row, 7).Value = line;
                    row++; // her satırdan sonra bir alt hücreye geç
                }
                //Not

                //filtre
                satir = 1;
                sutun = 0;
                int filterCount = 0;
                for (int i = 0; i < 18; i++)
                {
                    if (Datas.Filtreler[i, 0] == "1")
                    {
                        filterCount++;
                        if (filterCount < 9) 
                        {
                            worksheet.Cell(41 + satir - 1, 1 + sutun).Value = Datas.Filtreler[i, 1] + ": Checked";
                            satir += 1;
                        }
                    }
                    if (satir % 5 == 0)
                    {
                        satir = 1;
                        sutun += 3;
                    }
                }
                //filtre

                //tarih
                worksheet.Cell(9, 10).Value = ": " + DateTime.Now.ToString();
                worksheet.Cell(55, 10).Value = ": " + DateTime.Now.ToString();
                //tarih

                SensorChecked(worksheet, Datas.FreshAirSensorValues, 56, 4, "-");
                SensorChecked(worksheet, Datas.SupplyAirSensorValues, 57, 4, "-");
                SensorChecked(worksheet, Datas.ReturnAirSensorValues, 58, 4, "-");
                SensorChecked(worksheet, Datas.ExhaustAirSensorValues, 59, 4, "-");
                SensorChecked(worksheet, Datas.AfterCoilAirSensorValues, 60, 4, "-");
                SensorChecked(worksheet, Datas.MixAirSensorValues, 61, 4, "-");
                SensorChecked(worksheet, Datas.RoomTempSensor1Values, 56, 10, "-");
                SensorChecked(worksheet, Datas.RoomTempSensor2Values, 57, 10, "-");
                SensorChecked(worksheet, Datas.ReturnCO2SensorValues, 58, 10, "-");
                SensorChecked(worksheet, Datas.WaterTempSensorValues, 59, 10, "-");
                SensorChecked(worksheet, Datas.ReturnCO2AirSensorValues,  60, 10, "-");
                ElectricalHeaterWrite(worksheet, Datas.ElectricalHeater, 64, 3);
                int detailRow = 64;
                int detailCol = 1;
                if (Datas.ElectricalHeater==1)
                {
                    detailRow = 70;
                }
               
                detailRow = SensorValueWrite2(worksheet, Datas.FreshAirSensorValues, "FreshAirSensor", detailRow, detailCol, bosluk);
           
                detailRow = SensorValueWrite2(worksheet, Datas.SupplyAirSensorValues, "SupplyAirSensor", detailRow, detailCol, bosluk);
               
                detailRow = SensorValueWrite2(worksheet, Datas.ReturnAirSensorValues, "ReturnAirSensor", detailRow, detailCol, bosluk);
       
                detailRow = SensorValueWrite2(worksheet, Datas.ExhaustAirSensorValues, "ExhaustAirSensor", detailRow, detailCol, bosluk);
              
                detailRow = SensorValueWrite2(worksheet, Datas.AfterCoilAirSensorValues, "AfterCoilAirSensor", detailRow, detailCol, bosluk);
               
                detailRow = SensorValueWrite2(worksheet, Datas.MixAirSensorValues, "MixCoilAirSensor", detailRow, detailCol, bosluk);
                
                detailRow = SensorValueWrite2(worksheet, Datas.RoomTempSensor1Values, "RoomTempSensor1", detailRow, detailCol, bosluk);
            
                detailRow = SensorValueWrite2(worksheet, Datas.RoomTempSensor2Values, "RoomTempSensor2", detailRow, detailCol, bosluk);
               
                detailRow = SensorValueWrite2(worksheet, Datas.ReturnCO2SensorValues, "ReturnCO2Sensor", detailRow, detailCol, bosluk);
             
                detailRow = SensorValueWrite2(worksheet, Datas.WaterTempSensorValues, "WaterTempSensor", detailRow, detailCol, bosluk);
               
                detailRow = SensorValueWrite2(worksheet, Datas.ReturnCO2AirSensorValues, "ReturnCO2AirSensor", detailRow, detailCol, bosluk);
            }
            catch (IOException)
            {
                MessageBox.Show("Excel dosyası şu anda açık. Lütfen kapatıp tekrar deneyin.",
                                "Dosya Kullanımda", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // İşleme devam etme
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Beklenmedik bir hata oluştu: {ex.Message}",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public static void ExcelSave()
        {
            // Kaydet (mevcut dosyanın üzerine günceller)
            try
            {
                endFilePath=System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), (Datas.ProjectInfo[2] +"_Test_Report_BACnet_Modbus.xlsx"));
                workbook.SaveAs(endFilePath);
                MessageBox.Show("Dosya başarıyla kaydedildi!", "Bilgi", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (IOException)
            {
                MessageBox.Show("Excel dosyası şu anda açık. Lütfen kapatıp tekrar deneyin.",
                                "Dosya Kullanımda", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Beklenmedik bir hata oluştu: {ex.Message}",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            MessageBox.Show($"Excel dosyası güncellendi");

        }
        public static void ElectricalHeaterWrite(IXLWorksheet worksheet, int electricalHeater, int startRow, int startCol)
        {
            if (electricalHeater==1)
            {
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startRow += 1;
                startCol -= 6;

                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startRow += 1;
                startCol -= 6;

                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startRow += 1;
                startCol -= 6;

                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();
                startCol += 2;
                worksheet.Range(startRow, startCol, startRow, startCol + 1).Merge();

                // Başlık (2 sütunu birleştirerek genişletiyoruz)
                startRow -= 3;
                startCol -= 6;
                worksheet.Cell(startRow, startCol).Value = "Electrical Heater";

                // Faz başlıkları (R, S, T) + 1 sütun boşluk ekleyerek genişlet
                worksheet.Cell(startRow, startCol + 2).Value = "R(A)";
                worksheet.Cell(startRow, startCol + 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(startRow, startCol + 4).Value = "S(A)";
                worksheet.Cell(startRow, startCol + 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Cell(startRow, startCol + 6).Value = "T(A)";
                worksheet.Cell(startRow, startCol + 6).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                // Kademe değerlerini yaz (3 kademe, 4 faz)
                for (int i = 0; i < 3; i++) // Kademe satırları
                {
                    worksheet.Cell(startRow + 1 + i, startCol).Value = $"{i + 1}. Kademe";

                    for (int j = 0; j < 3; j++) // R, S, T fazları
                    {
                        int index = i * 3 + j; // Tek boyutlu diziyi 2 boyuta çeviriyoruz
                        worksheet.Cell(startRow + 1 + i, startCol + 2 + j * 2).Value = Datas.ElectricalData[index];
                        worksheet.Cell(startRow + 1 + i, startCol + 2 + j * 2).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                        worksheet.Cell(startRow + 1 + i, startCol + 2 + j * 2).Style.NumberFormat.Format = "0.00";
                    }

                    worksheet.Cell(startRow + 2 + i, startCol + 4).Value = ""; // Extra hücre boş
                }

                // Tablo kenarlıkları (2 sütun genişlettiğimiz için +1 ekledik)
                var range = worksheet.Range(startRow, startCol, startRow + 3, startCol + 7);
                range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
            }
        }
        public static void SensorChecked(IXLWorksheet worksheet, Dictionary<string, string> dict,
         int checkRow, int checkCol, string bosluk)
        {
            if (dict == null) return;

            bool checkedWritten = false; // "Checked" yazıldı mı

            foreach (var kvp in dict)
            {
                string value = kvp.Value;

                // Değer geçerli mi?
                if (!string.IsNullOrEmpty(value) && value != "-" && value != "0")
                {
                    if (!checkedWritten)
                    {
                        worksheet.Cell(checkRow, checkCol).Value ="Checked";
                        checkedWritten = true;
                    }
                    // İsterseniz burada key veya value ile başka işlem de yapabilirsiniz
                }
            }
        }
        public static int SensorValueWrite2(IXLWorksheet ws, Dictionary<string, string> dict,
            string sensorName, int detailStartRow, int detailCol, string bosluk)
        {
            if (dict == null || dict.Values.All(v => v == "-")) return detailStartRow;

            // --- Eğer dict'teki tüm değerler null veya boş ise tabloyu yazma ---
            bool hasValue = dict.Values.Any(v => !string.IsNullOrWhiteSpace(v));
            if (!hasValue) return detailStartRow;

            int startRow = detailStartRow;

            // --- Başlık satırı ---

            string cellAll = "A" + detailStartRow.ToString() + ":H" + detailStartRow.ToString();
            var range = ws.Range(cellAll);
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;

           // ws.Cell(startRow, detailCol).Value = "Sensor";
            int col = detailCol + 2;

          
            foreach (var key in dict.Keys)
            {
                if (dict[key] != "-")
                {

                    ws.Cell(startRow, col).Value = key;

                    ws.Cell(startRow + 1, col).Value = dict[key];

                    ws.Cell(startRow + 1, col).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                    col += 2;
                }

            }

            // --- Sensör adı + değerler ---
            ws.Cell(startRow + 1, detailCol).Value = sensorName;
            col = detailCol + 2;

            return startRow + 2; // +3: 2 satır tablo + 1 satır boşluk
        }
        public static void RegisterDelete(IXLWorksheet worksheet)
        {
            // "ModbusRegister" sayfasını al ya da oluştur


            try
            {



                // fan kontrol registerları ***************************************
                if (Datas.FanTip == 1)
                {
                    ClearRows(worksheet, 57, 110); // EBM papst register
                    ClearRows(worksheet, 111, 146); // EC Zielh-Abegg register


                    if (Datas.SupplyFanNumber == 0)
                    {
                        ClearRows(worksheet, 147, 164); // Supply FC101 Vantilatör 1
                        ClearRows(worksheet, 165, 182); // Supply FC101 Vantilatör 2
                    }
                    else if (Datas.SupplyFanNumber == 1)
                    {
                        ClearRows(worksheet, 165, 182); // Supply FC101 Vantilatör 2
                    }


                    if (Datas.ReturnFanNumber == 0)
                    {
                        ClearRows(worksheet, 183, 200); // Return FC101 Aspiratör 1
                        ClearRows(worksheet, 201, 218); // Return FC101 Aspiratör 2
                    }
                    else if (Datas.ReturnFanNumber == 1)
                    {
                        ClearRows(worksheet, 201, 218); // Return FC101 Aspiratör 2
                    }

                }
                else if (Datas.FanTip == 2)
                {
                    ClearRows(worksheet, 57, 110); // EBM papst register
                    ClearRows(worksheet, 147, 218); // Danfoss Zielh-Abegg register

                    if (Datas.SupplyFanNumber == 0)
                    {
                        ClearRows(worksheet, 111, 116); // EC Zielh-Abegg Supply register 1
                        ClearRows(worksheet, 123, 128); // EC Zielh-Abegg Supply register 2
                        ClearRows(worksheet, 135, 140); // EC Zielh-Abegg Supply register 3
                    }
                    else if (Datas.SupplyFanNumber == 1)
                    {
                        ClearRows(worksheet, 123, 128); // EC Zielh-Abegg Supply register 2
                        ClearRows(worksheet, 135, 140); // EC Zielh-Abegg Supply register 3
                    }
                    else if (Datas.SupplyFanNumber == 2)
                    {
                        ClearRows(worksheet, 135, 140); // EC Zielh-Abegg Supply register 3
                    }


                    if (Datas.ReturnFanNumber == 0)
                    {
                        ClearRows(worksheet, 117, 122); // EC Zielh-Abegg Return register 1
                        ClearRows(worksheet, 129, 134); // EC Zielh-Abegg Return register 2
                        ClearRows(worksheet, 141, 146); // EC Zielh-Abegg Return register 3
                    }
                    else if (Datas.ReturnFanNumber == 1)
                    {
                        ClearRows(worksheet, 129, 134); // EC Zielh-Abegg Return register 2
                        ClearRows(worksheet, 141, 146); // EC Zielh-Abegg Return register 3
                    }
                    else if (Datas.ReturnFanNumber == 2)
                    {
                        ClearRows(worksheet, 141, 146); // EC Zielh-Abegg Return register 3
                    }

                }
                else if (Datas.FanTip == 3)
                {
                    ClearRows(worksheet, 111, 146); // EC Zielh-Abegg register
                    ClearRows(worksheet, 147, 218); // Danfoss Zielh-Abegg register

                    if (Datas.SupplyFanNumber == 0)
                    {
                        ClearRows(worksheet, 57, 65); // EC EBM Supply 1 register
                        ClearRows(worksheet, 66, 74); // EC EBM Supply 2 register
                        ClearRows(worksheet, 75, 83); // EC EBM Supply 3 register
                    }
                    else if (Datas.SupplyFanNumber == 1)
                    {
                        ClearRows(worksheet, 66, 74); // EC EBM Supply 2 register
                        ClearRows(worksheet, 75, 83); // EC EBM Supply 3 register
                    }
                    else if (Datas.SupplyFanNumber == 2)
                    {
                        ClearRows(worksheet, 75, 83); // EC EBM Supply 3 register
                    }


                    if (Datas.ReturnFanNumber == 0)
                    {
                        ClearRows(worksheet, 84, 92); // EC EBM Return 1 register
                        ClearRows(worksheet, 93, 101); // EC EBM Return 2 register
                        ClearRows(worksheet, 102, 110); // EC EBM Return 3 register
                    }
                    else if (Datas.ReturnFanNumber == 1)
                    {
                        ClearRows(worksheet, 93, 101); // EC EBM Return 2 register
                        ClearRows(worksheet, 102, 110); // EC EBM Return 3 register
                    }
                    else if (Datas.ReturnFanNumber == 2)
                    {
                        ClearRows(worksheet, 102, 110); // EC EBM Return 3 register
                    }
                }
                else
                {
                    ClearRows(worksheet, 57, 110); // EBM papst register
                    ClearRows(worksheet, 111, 146); // EC Zielh-Abegg register
                    ClearRows(worksheet, 147, 218); // Danfoss Zielh-Abegg register
                }



                if (Datas.SupplyFanNumber == 0)
                {
                    ClearRows(worksheet, 411, 437); // SAF Register
                }
                if (Datas.ReturnFanNumber == 0)
                {
                    ClearRows(worksheet, 438, 464); // EAF Register
                }


                if (Datas.AirFlowControlOk == false)
                {
                    ClearRows(worksheet, 415, 417);
                    ClearRows(worksheet, 419, 422);
                    worksheet.Row(436).Clear(XLClearOptions.Contents); //SAF Pressure High Alarm
                }
                if (Datas.PressureControlOk == false)
                {

                    ClearRows(worksheet, 426, 430); // Duct pressure register
                    worksheet.Row(434).Clear(XLClearOptions.Contents); //SAF Pressure High Alarm


                }
                // fan kontrol registerları ***************************************


                // Modül kontrol registerları *************************************
                if (Datas.RotorData[0, 0] == 0)
                {
                    ClearRows(worksheet, 399, 405); // Rotor register
                }
                else
                {
                    if (Datas.RotorData[0, 1] == 1)
                    {
                        ClearRows(worksheet, 399, 401); // Rotor oransal register
                    }
                }


                if (Datas.RunAroundData == 0)
                {
                    ClearRows(worksheet, 406, 410); // RunAround register
                }
                if (Datas.ChangeOverData == 0)
                {
                    ClearRows(worksheet, 521, 532); // ChangeOver register
                }
                if (Datas.DXData[0, 0] == 0)
                {
                    ClearRows(worksheet, 219, 232); // DX register
                }
                else
                {
                    switch (Datas.DXData[0, 1])
                    {
                        case 1:
                            ClearRows(worksheet, 221, 229); // DX Kademe 2,3,4,5
                            break;
                        case 2:
                            ClearRows(worksheet, 223, 229); // DX Kademe 3,4,5
                            break;
                        case 3:
                            ClearRows(worksheet, 225, 229); // DX Kademe 4,5
                            break;
                        case 4:
                            ClearRows(worksheet, 227, 229); // DX Kademe 5
                            break;
                        default:
                            ClearRows(worksheet, 219, 229); 
                            break;
                    }
                }


                if (Datas.NemlendiriciData[0, 0] == 0)
                {
                    ClearRows(worksheet, 465, 520); // Nemlendirici register
                }
                else
                {
                    switch (Datas.NemlendiriciData[0, 1])
                    {
                        case 1:
                            ClearRows(worksheet, 470, 505); // Nem Kademe 2,3,4,5,6,7,8
                            break;
                        case 2:
                            ClearRows(worksheet, 475, 505); // Nem Kademe 3,4,5,6,7,8
                            break;
                        case 3:
                            ClearRows(worksheet, 480, 505); // Nem Kademe 4,5,6,7,8
                            break;
                        case 4:
                            ClearRows(worksheet, 485, 505); // Nem Kademe 5,6,7,8
                            break;
                        case 5:
                            ClearRows(worksheet, 490, 505); // Nem Kademe 6,7,8
                            break;
                        case 6:
                            ClearRows(worksheet, 495, 505); // Nem Kademe 7,8
                            break;
                        case 7:
                            ClearRows(worksheet, 500, 505); // Nem Kademe 8
                            break;
                        case 8:
                            break;
                        default:
                            ClearRows(worksheet, 465, 505);
                            break;
                    }
                }


                if (Datas.ValveData[0] == 0) // ısıtma 1 register
                {
                    worksheet.Row(371).Clear(XLClearOptions.Contents);
                    worksheet.Row(375).Clear(XLClearOptions.Contents);
                    worksheet.Row(383).Clear(XLClearOptions.Contents);
                    worksheet.Row(384).Clear(XLClearOptions.Contents);
                }
                if (Datas.ValveData[1] == 0) // soğutma 1 register
                {
                    worksheet.Row(372).Clear(XLClearOptions.Contents);
                    worksheet.Row(374).Clear(XLClearOptions.Contents);
                    worksheet.Row(381).Clear(XLClearOptions.Contents);
                    worksheet.Row(382).Clear(XLClearOptions.Contents);
                }
                if (Datas.ValveData[2] == 0) // soğutma 2 register
                {
                    worksheet.Row(377).Clear(XLClearOptions.Contents);

                }
                if (Datas.ValveData[3] == 0) // ısıtma 2 register
                {
                    worksheet.Row(376).Clear(XLClearOptions.Contents);
                }


                if (Datas.ElectricalHeater == 0)
                {
                    ClearRows(worksheet, 393, 398); // Elektrikli ısıtıcı register
                }



                if (Datas.Components[0] == 0)
                {
                    worksheet.Row(254).Clear(XLClearOptions.Contents); // phase alarm register
                }
                if (Datas.Components[3] == 0)
                {
                    worksheet.Row(253).Clear(XLClearOptions.Contents); // fire alarm register
                }
                if (Datas.Components[4] == 0)
                {
                    worksheet.Row(252).Clear(XLClearOptions.Contents); // emergency register
                }
                if (Datas.Components[5] == 0)
                {
                    worksheet.Row(257).Clear(XLClearOptions.Contents); // frost alarm register
                }
                if (Datas.Components[6] == 0)
                {
                    worksheet.Row(397).Clear(XLClearOptions.Contents); // high temp alarm register
                }



                if (Datas.SupplyFanNumber == 0)
                {
                    worksheet.Row(255).Clear(XLClearOptions.Contents); // van kapı alarm register
                }
                if (Datas.ReturnFanNumber == 0)
                {
                    worksheet.Row(256).Clear(XLClearOptions.Contents); // asp kapı alarm register
                }
                // Modül kontrol registerları *************************************


                // filtre registerları ********************************************
          
                for (int i = 0; i < 18; i++)
                {
                    if (Datas.Filtreler[i, 0] == "0" || Datas.Filtreler[i, 0] == null)
                    {
                        worksheet.Row(i + 234).Clear(XLClearOptions.Contents); // filtre register
                    }
                }
                // filtre registerları ********************************************

                //Room BMS ve Temp Average blokları *****************************
                if (Datas.RoomBMS == 0)
                {
                    ClearRows(worksheet, 53, 56); // Room BMS register
                }
                if (Datas.TempAvgEn == 0)
                {
                    ClearRows(worksheet, 48, 53); // Temp Average register
                }
                //Room BMS ve Temp Average blokları *****************************

                // Bypass Damper registerları ********************************************
                if (Datas.DamperDatas[5, 0] == 0)
                {
                    ClearRows(worksheet, 259, 272); // Bypass damper kontrol register
                }
                else if (Datas.DamperDatas[5, 0] == 1)
                {
                    ClearRows(worksheet, 268, 272); // Bypass damper 2 register
                }
                // Bypass Damper registerları ********************************************

                // Mix Damper registerları ********************************************
                switch (Datas.DamperDatas[4, 0])
                {
                    case 0:
                        ClearRows(worksheet, 345, 365); // Mix damper kontrol register
                        break;
                    case 1:
                        ClearRows(worksheet, 357, 365); // Mix damper 1
                        break;
                    case 2:
                        ClearRows(worksheet, 361, 365); // Mix damper 1, 2
                        break;
                    case 3:
                        ClearRows(worksheet, 363, 365); // Mix damper 1, 2, 3
                        break;
                    default:
                        // Mix damper 1, 2, 3, 4
                        break;
                }
                // Mix Damper registerları ********************************************

                // Supply damper registerları ******************
                switch (Datas.DamperDatas[1, 0])
                {
                    case 0:
                        ClearRows(worksheet, 273, 290); // Supply damper register
                        break;
                    case 1:
                        ClearRows(worksheet, 278, 290);  // Supply damper 1
                        break;
                    case 2:
                        ClearRows(worksheet, 282, 290); // Supply damper 1, 2
                        break;
                    case 3:
                        ClearRows(worksheet, 286, 290); // Supply damper 1, 2, 3
                        break;
                    default:
                        // Supply damper 1, 2, 3, 4
                        break;
                }
                // Supply damper registerları ******************

                // Fresh damper registerları ******************
                switch (Datas.DamperDatas[0, 0])
                {
                    case 0:
                        ClearRows(worksheet, 291, 308); // Fresh damper register
                        break;
                    case 1:
                        ClearRows(worksheet, 296, 308);  // Fresh damper 1
                        break;
                    case 2:
                        ClearRows(worksheet, 300, 308); // Fresh damper 1, 2
                        break;
                    case 3:
                        ClearRows(worksheet, 304, 308); // Fresh damper 1, 2, 3
                        break;
                    default:
                        // Fresh damper 1, 2, 3, 4
                        break;
                }
                // Fresh damper registerları ******************

                // Return damper registerları ******************
                switch (Datas.DamperDatas[2, 0])
                {
                    case 0:
                        ClearRows(worksheet, 309, 326); // Return damper register
                        break;
                    case 1:
                        ClearRows(worksheet, 314, 326);  // Return damper 1
                        break;
                    case 2:
                        ClearRows(worksheet, 318, 326); // Return damper 1, 2
                        break;
                    case 3:
                        ClearRows(worksheet, 322, 326); // Return damper 1, 2, 3
                        break;
                    default:
                        // Return damper 1, 2, 3, 4
                        break;
                }
                // Return damper registerları ******************

                // Exhaust damper registerları ******************
                switch (Datas.DamperDatas[3, 0])
                {
                    case 0:
                        ClearRows(worksheet, 327, 344); // Exhaust damper register
                        break;
                    case 1:
                        ClearRows(worksheet, 332, 344); // Exhaust damper 1
                        break;
                    case 2:
                        ClearRows(worksheet, 336, 344); // Exhaust damper 1, 2
                        break;
                    case 3:
                        ClearRows(worksheet, 340, 344); // Exhaust damper 1, 2, 3
                        break;
                    default:
                        // Exhaust damper 1, 2, 3, 4
                        break;
                }
                // Return damper registerları ******************



                // sensor registerları *****************************************
                if (Datas.FreshAirSensorValues == null || Datas.FreshAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 16, 22); // Fresh Air Sensor Register
                }
                else
                {
                    string[] keyOrder = { 
                    Datas.FreshAirSensorValues.Keys.ElementAt(0),
                    Datas.FreshAirSensorValues.Keys.ElementAt(1),
                    Datas.FreshAirSensorValues.Keys.ElementAt(2)};
                    int[] rowsToClear = { 16, 17, 18, 19, 20 };

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.FreshAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.SupplyAirSensorValues == null || Datas.SupplyAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 2, 8); // Supply Air Sensor Register
                }
                else
                {
                    string[] keyOrder = { 
                    Datas.SupplyAirSensorValues.Keys.ElementAt(0),
                    Datas.SupplyAirSensorValues.Keys.ElementAt(1),
                    Datas.SupplyAirSensorValues.Keys.ElementAt(2)};
                    int[] rowsToClear = { 2, 3, 4, 5, 6 };

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.SupplyAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.ReturnAirSensorValues == null || Datas.ReturnAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 9, 15); // Return Air Sensor Register
                }
                else
                {
                    string[] keyOrder = { 
                    Datas.ReturnAirSensorValues.Keys.ElementAt(0),
                    Datas.ReturnAirSensorValues.Keys.ElementAt(1),
                    Datas.ReturnAirSensorValues.Keys.ElementAt(2)};
                    int[] rowsToClear = { 9, 10, 11, 12, 13 };

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.ReturnAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.ExhaustAirSensorValues == null || Datas.ExhaustAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 28, 32); // Exhaust Air Sensor Register
                }
                else
                {
                    string[] keyOrder = { 
                    Datas.ExhaustAirSensorValues.Keys.ElementAt(0),
                    Datas.ExhaustAirSensorValues.Keys.ElementAt(1),
                    Datas.ExhaustAirSensorValues.Keys.ElementAt(2)};
                    int[] rowsToClear = { 28, 29, 30 };

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.ExhaustAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.AfterCoilAirSensorValues == null || Datas.AfterCoilAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 23, 27); // AfterCoil Air Sensor Register
                }
                else
                {
                    int[] rowsToClear = { 23, 24, 25 };
                    string[] keyOrder = { 
                    Datas.AfterCoilAirSensorValues.Keys.ElementAt(0),
                    Datas.AfterCoilAirSensorValues.Keys.ElementAt(1),
                    Datas.AfterCoilAirSensorValues.Keys.ElementAt(2)};

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.AfterCoilAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.MixAirSensorValues == null || Datas.MixAirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 39, 43); // Mix Air Sensor Register
                }
                else
                {
                    int[] rowsToClear = { 39, 40, 41 };
                    string[] keyOrder = { 
                    Datas.MixAirSensorValues.Keys.ElementAt(0),
                    Datas.MixAirSensorValues.Keys.ElementAt(1),
                    Datas.MixAirSensorValues.Keys.ElementAt(2)};


                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.MixAirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.ReturnCO2AirSensorValues == null || Datas.ReturnCO2AirSensorValues.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 44, 47); // ReturnCO2 Air Sensor Register
                }
                else
                {
                    int[] rowsToClear = { 44, 45 };
                    string[] keyOrder = {
                    Datas.ReturnCO2AirSensorValues.Keys.ElementAt(0),
                    Datas.ReturnCO2AirSensorValues.Keys.ElementAt(1)};

                    for (int i = 0; i < keyOrder.Length && i < rowsToClear.Length; i++)
                    {
                        if (Datas.ReturnCO2AirSensorValues.TryGetValue(keyOrder[i], out var value) &&
                            (string.IsNullOrEmpty(value) || value == "-"))
                        {
                            worksheet.Row(rowsToClear[i]).Clear(XLClearOptions.Contents);
                        }
                    }
                }

                if (Datas.ReturnCO2SensorValues == null || Datas.ReturnCO2SensorValues.Values.All(v => v == "-"))
                {
                    worksheet.Row(33).Clear(XLClearOptions.Contents); // ReturnCO2 Sensor Register
                }

                if (Datas.RoomTempSensor1Values == null || Datas.RoomTempSensor1Values.Values.All(v => v == "-"))
                {
                    ClearRows(worksheet, 34, 36);
                }

                if (Datas.WaterTempSensorValues == null || Datas.WaterTempSensorValues.Values.All(v => v == "-"))
                {
                    worksheet.Row(36).Clear(XLClearOptions.Contents); // Water Sensor Register
                }

                if (Datas.RoomTempSensor2Values == null || Datas.RoomTempSensor2Values.Values.All(v => v == "-"))
                {
                    worksheet.Row(37).Clear(XLClearOptions.Contents); // Water Sensor Register
                }


                // sensor registerları *****************************************



                // ws: IXLWorksheet
                for (int i = worksheet.LastRowUsed().RowNumber(); i >= 1; i--)
                {
                    // Satırda hiç değer veya formül yoksa
                    if (worksheet.Row(i).IsEmpty())
                    {
                        worksheet.Row(i).Delete();
                    }
                }
            }
            catch (IOException)
            {
                MessageBox.Show("Excel dosyası şu anda açık. Lütfen kapatıp tekrar deneyin.",
                                "Dosya Kullanımda", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // İşleme devam etme
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Beklenmedik bir hata oluştu: {ex.Message}",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

        }
        public static void ExcelWrite()
        {
            try
            {
                var worksheetTestReport = workbook.Worksheets.Contains("TestRaporu")
                 ? workbook.Worksheet("TestRaporu")
                 : workbook.Worksheets.Add("TestRaporu");

                var worksheetModbus = workbook.Worksheets.Contains("ModbusRegister")
                 ? workbook.Worksheet("ModbusRegister")
                 : workbook.Worksheets.Add("ModbusRegister");

                var worksheetBACnet = workbook.Worksheets.Contains("BACnetRegister")
                 ? workbook.Worksheet("BACnetRegister")
                 : workbook.Worksheets.Add("BACnetRegister");

                RegisterDelete(worksheetModbus);
                RegisterDelete(worksheetBACnet);
                TestReportWrite(worksheetTestReport);
            }
            catch (IOException)
            {
                MessageBox.Show("Excel dosyası şu anda açık. Lütfen kapatıp tekrar deneyin.",
                                "Dosya Kullanımda", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // İşleme devam etme
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Beklenmedik bir hata oluştu: {ex.Message}",
                                "Hata", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }
        public static void ClearRows(IXLWorksheet ws, int startRow, int endRow)
        {
            for (int i = startRow; i < endRow; i++)
            {
                ws.Row(i).Clear(XLClearOptions.Contents);
            }
        }

        public static void TopLineDraw(IXLWorksheet ws,int ilkHucre, int sonHucre)
        {
            string cellAll = "A" + ilkHucre.ToString() + ":H" + sonHucre.ToString();
            var range = ws.Range(cellAll);
            range.Style.Border.TopBorder = XLBorderStyleValues.Thin;
            MessageBox.Show(cellAll);
        }
    }
}