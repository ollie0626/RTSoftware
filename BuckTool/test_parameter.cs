using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuckTool
{

    public struct Hi_Lo
    {
        public double Highlevel;
        public double LowLevel;
    }


    public static class test_parameter
    {
        public static List<double> Iout_table = new List<double>();
        public static List<double> Vin_table = new List<double>();
        public static List<string> temp_table = new List<string>();

        public static List<double> tempList = new List<double>();
        public static List<Hi_Lo> HiLo_table = new List<Hi_Lo>();

        public static List<double> HiLevel = new List<double>();
        public static List<double> LoLevel = new List<double>();

        // High Voltage buck frequency control
        public static bool[] Freq_en = new bool[2];
        public static string[] Freq_des = new string[2];
        public static double vout_ideal;
        
        // bin folder path
        public static string specify_bin;
        public static string waveform_path;
        public static string binFolder;

        public static bool run_stop;
        public static bool chamber_en;

        // load transtion variables
        public static double freq;
        public static double duty;
        public static double tr;
        public static double tf;

        // shutdown current
        public static int interval;
        public static int test_cnt;
        public static int en_ms;

        // chamber parameter
        //public static int item;
        //public static int steadyTime;

    }
}
