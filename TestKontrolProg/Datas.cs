using System;
using System.Collections.Generic;
using System.Data.SqlTypes;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;

namespace TestKontrolProg
{
    internal class Datas
    {

        public static byte[] imgBytes = Convert.FromBase64String(Properties.Settings.Default.DefaultImageBase64);
        //File.ReadAllBytes(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,"null.png"));

        public static string UserName = "";

        //klasör ve dosya isimleri
        public static string pdfPath = "";
        public static string newPdfPath = "";
        //klasör ve dosya isimleri


        //Santral Bilgileri
        public static string[] ProjectInfo = {"","",""};
        //Santral Bilgiler


        //fan kontrol variable
        public static int FanTip = 0;
        public static int SupplyFanNumber = 0;
        public static int ReturnFanNumber = 0;
        public static int SupplyAirFlow = 0;
        public static int ReturnAirFlow = 0;
        public static bool AirFlowControlOk = false;
        public static bool PressureControlOk = false;
        //fan kontrol variable


        //Filtre Kontrol
        public static string[,] Filtreler = new string[18,2];
        //Filtre Kontrol


        //Sensor Kontrol
        public static Dictionary<string, string> FreshAirSensorValues =
           new Dictionary<string, string>
           {
               { "Sıcaklık (C°)", "-" }, 
               { "Nem(%)", "-" },  
               { "CO2(ppm)", "-" }          
           };
        public static Dictionary<string, string> SupplyAirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" },
               { "CO2(ppm)", "-" }
          };
        public static Dictionary<string, string> ReturnAirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" },
               { "CO2(ppm)", "-" }
          };
        public static Dictionary<string, string> ExhaustAirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" },
               { "CO2(ppm)", "-" } 
          };
        public static Dictionary<string, string> AfterCoilAirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" },
               { "CO2(ppm)", "-" } 
          };
        public static Dictionary<string, string> MixAirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" },
               { "CO2(ppm)", "-" }
          };
        public static Dictionary<string, string> RoomTempSensor1Values =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "Nem(%)", "-" }
          };
        public static Dictionary<string, string> RoomTempSensor2Values =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" }
          };
        public static Dictionary<string, string> ReturnCO2SensorValues =
          new Dictionary<string, string>
          {
               { "CO2(ppm)", "-" }
          };
        public static Dictionary<string, string>WaterTempSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" }
          };
        public static Dictionary<string, string> ReturnCO2AirSensorValues =
          new Dictionary<string, string>
          {
               { "Sıcaklık (C°)", "-" },
               { "CO2(ppm)", "-" }
          };
        //Sensor Kontrol


        //Modul Kontrol
        public static int[,] RotorData = new int[1, 2];
        public static int RunAroundData = new int();
        public static int ChangeOverData = 0;
        public static int[,] DXData = new int[1, 2];
        public static int[,] NemlendiriciData = new int[1, 2];
        public static int[] ValveData = new int[4];
        public static int ElectricalHeater = 0;
        public static float[] ElectricalData = new float[9];
        public static int[] Components = new int[7];
        public static int RoomBMS = 0;
        public static int TempAvgEn = 0;
        //Modul Kontrol
        public static string[] Modules1 = { "-", "-", "-", "-", "-", "-", "-", "-", "-", "-"};
        public static string[] Modules2 = { "-", "-", "-", "-", "-", "-", "-", "-", "-" };

        //Sensor var mı 
        public static int[] SensorVar = new int[11];
        //Sensor var mı 


        //Damper Kontrol
        public static int[,] DamperDatas = new int[6, 2];
        //Damper Kontrol


        //Not
        public static List<string> Not = new List<string>();
        //Not
    }
}
