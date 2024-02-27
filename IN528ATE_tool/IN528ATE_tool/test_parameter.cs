using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IN528ATE_tool
{
    static public class test_parameter
    {
        public static double ch2_level;


        public static List<double> VinList = new List<double>();
        public static List<double> IoutList = new List<double>();
        public static string binFolder;
        public static byte slave;
        public static byte specify_id;
        public static string specify_bin;
        public static string waveform_path;
        public static double ontime_scale_ms;
        public static double offtime_scale_ms;
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
        public static bool trigger_en;
        public static double trigger_level;
        public static double measure_level;

        /* MTP program */
        public static byte mtp_slave;
        public static byte mtp_addr;
        public static byte mtp_data;
        public static bool mtp_enable;

        /* Current Limit */
        public static double cv_setting;
        public static double cv_step;
        public static double cv_wait;

        /* add SST & DT */
        public static double lovol;
        public static double midvol;
        public static double hivol;

        public static double lovout;
        public static double midvout;
        public static double hivout;

        /* add for swire */
        public static bool swire_en;
        public static List<string> swireList = new List<string>();
        public static List<double> voutList = new List<double>();
        public static bool swire_20;

        public static bool bw_en;
        public static int sst_sel;
        public static bool dt_rising_en;

        public static bool ripple_time_manual;
    }


}
