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
        // interface
        public static byte slave;
        public static string bin_path;
        public static string wave_path;
        public static string special_file;
        public static List<string> swireList = new List<string>();
        public static bool i2c_enable;
        public static bool swire_20;

        // power
        public static List<double> vinList = new List<double>();

        // ELoad
        public static List<double> ioutList = new List<double>();
        public static bool eload_select;

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

    }
}
