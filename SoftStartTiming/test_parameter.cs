using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SoftStartTiming
{
    static public class test_parameter
    {
        public static List<double> VinList = new List<double>();

        public static string vin_conditions;
        public static string tool_ver;
        public static string bin_file_cnt;
        public static int bin1_cnt;
        public static int bin2_cnt;
        public static int bin3_cnt;
        
        // 0: gpio trigger
        // 1: i2c trigger
        // 2: vin trigger
        public static int trigger_event;
        // for power on and soft-start time 
        // wake up delay test
        public static double sleep_en;
        public static byte slave;
        public static bool sleep_mode;
        public static string[] bin_path = new string[3];
        public static bool[] scope_en = new bool[3];
        public static bool[] bin_en = new bool[3];
        public static string waveform_path;
        public static double ontime_scale_ms;
        public static double offtime_scale_ms;
        public static double offset_time;
        public static bool delay_us_en;

        public static double judge_percent;


        // gpio select
        public static int gpio_pin;
        public static string power_mode;


        public static bool run_stop;
        public static bool chamber_en;

        /* trigger select */
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

        /* add for swire */
        public static bool swire_en;
        public static List<string> swireList = new List<string>();
        public static List<double> voutList = new List<double>();
        public static bool swire_20;
    }


}
