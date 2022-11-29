using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FFT_Single
{
    public static class test_parameter
    {
        public static string vin_info;
        public static string eload_info;
        public static string swire_info;
        public static string date_info;
        public static string ver_info;

        // user input
        //------------------------------------------
        public static bool brightness_sel;
        public static double peak_level1;
        public static double peak_level2;
        public static double freq;
        public static double duty;
        public static int i2c_code;
        public static List<double> vinList = new List<double>();
        public static List<double> ioutList = new List<double>();
        public static string bin_path;
        public static string wave_path;


        public static int channel;
        public static bool run_stop;
    }
}
