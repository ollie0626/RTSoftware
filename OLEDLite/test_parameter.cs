using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OLEDLite
{

    public struct Hi_Lo
    {
        public double Highlevel;
        public double LowLevel;
    }

    public static class test_parameter
    {
        // test condition
        public static string vin_info;
        public static string eload_info;
        public static string swire_info;
        public static string date_info;
        public static string ver_info;

        // interface
        public static byte slave;
        public static string bin_path;
        public static string wave_path;
        public static string special_file;
        public static int swire_cnt;
        public static List<string> ESwireList = new List<string>();
        public static List<string> ASwireList = new List<string>();
        public static bool ESwire_state;
        public static bool ASwire_state;
        public static bool ENVO4_state;
        public static bool i2c_enable;
        public static bool CodeInrush_ESwire;

        // power
        public static List<double> vinList = new List<double>();

        // ELoad
        public static List<double> ioutList = new List<double>();
        public static bool eload_select;
        public static int eload_ch_select;
        public static int eload_iin_select;

        // FuncGen
        public static List<double> HiLevel = new List<double>();
        public static List<double> LoLevel = new List<double>();
        public static List<Hi_Lo> HiLo_table = new List<Hi_Lo>();
        public static double Freq;
        public static double duty;
        public static double tr;
        public static double tf;

        // eload enable
        public static bool[] eload_en;
        public static double[] eload_iout;


        // chamber parameter
        public static bool chamber_en;
        public static List<double> tempList = new List<double>();
        public static int steadyTime;
        public static bool run_stop;

        // burst period
        public static double burst_period = 1/(1.45 * Math.Pow(10, 6));


        // code inrush
        public static double vol_max;
        public static double vol_min;
        public static double ontime_scale_ms;
        public static int code_max;
        public static int code_min;
        public static byte addr;

        // current limit
        public static double cv_setting;
        public static double cv_wait;
        public static double cv_step;

        // Lx trigger
        public static bool buck;
        public static bool boost;
        public static bool inverting;
        public static bool[] LX_item = new bool[3]; // freq, sr, jitter

    }
}
