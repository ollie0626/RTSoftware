using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SDCTool
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

        //public static string swire_info;
        public static string date_info;
        public static string ver_info;

        public static byte slave;
        public static string bin_path;

        public static List<double> vinList = new List<double>();
        // loading include AC and DC setting
        public static List<double> ioutList = new List<double>();

        public static bool chamber_en;
        public static List<double> tempList = new List<double>();
        public static int steadyTime;
        public static bool run_stop;

    }
}
