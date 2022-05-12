using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IN528ATE_tool
{
    static public class test_parameter
    {
        public static List<double> VinList = new List<double>();
        public static List<double> IoutList = new List<double>();
        public static string binFolder;
        public static byte slave;
        public static byte specify_id;
        public static string specify_bin;
        public static string waveform_path;
        public static double time_scale_ms;
        public static bool all_en;

        public static bool run_stop;
        public static bool chamber_en;

        /* inrush code coditons */
        public static byte addr;
        public static byte max;
        public static byte min;
        public static double vol_max;
        public static double vol_min;

        /* relay variable */
        public static int[] relay_gpio1 = new int[8];
        public static int[] relay_gpio2 = new int[8];

        /* trigger select */
        public static bool trigger_vin_en;
    }


}
