using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoftStartTiming
{
    static public class test_parameter
    {
        public static List<double> VinList = new List<double>();
        public static List<double> IoutList = new List<double>();

        public static double ch1_level;
        public static double ch2_level;
        public static double ch3_level;
        public static double ch4_level;

        public static bool eload_cr;
        public static string vin_conditions;
        public static string iout_conditions;
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

        public static int item_idx;
        public static string[] bin_path = new string[3];
        public static string[] power_off_bin_path = new string[3];

        public static bool[] scope_en = new bool[3];
        public static bool[] bin_en = new bool[3];

        public static string waveform_path;
        public static string waveform_name;

        public static double ontime_scale_ms;
        public static double offtime_scale_ms;
        public static double offset_time;
        public static bool delay_us_en;
        public static double judge_percent;

        public static double LX_Level;
        public static double ILX_Level;

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

        public static int meas_cnt;


        public static byte Rail_addr;
        public static byte Rail_en;
        public static byte Rail_dis;
        public static DataGridView i2c_setting;

        /* Cross talk */
        // test conditions
        public static int ch_num;
        public static string[] rail_name;
        public static byte[] freq_addr;
        public static byte[] vout_addr;

        public static byte[] en_addr;
        public static byte[] en_data;
        public static byte[] disen_data;

        public static byte[] vid_addr;
        public static byte[] hi_code;
        public static byte[] lo_code;

        public static Dictionary<int, List<double>> ccm_eload = new Dictionary<int, List<double>>();
        public static Dictionary<int, List<double>> full_load = new Dictionary<int, List<double>>();
        public static Dictionary<int, List<byte>> freq_data = new Dictionary<int, List<byte>>();
        public static Dictionary<int, List<byte>> vout_data = new Dictionary<int, List<byte>>();
        public static Dictionary<int, List<string>> freq_des = new Dictionary<int, List<string>>();
        public static Dictionary<int, List<string>> vout_des = new Dictionary<int, List<string>>();
        public static Dictionary<int, List<double>> lt_l1 = new Dictionary<int, List<double>>();
        public static Dictionary<int, List<double>> lt_l2 = new Dictionary<int, List<double>>();

        public static List<string> scope_chx = new List<string>();
        public static List<string> scope_lx = new List<string>();
        public static List<int> eload_chx = new List<int>();

        // ins active load
        //public static double[] full_load;
        public static byte[] cross_select;
        public static bool[] cross_en;
        public static int cross_mode;
        public static double tolerance;
        public static int accumulate;
        public static bool freq_en;
        public static bool vout_en;

        // modify for support 8 channel eload and 4 channel scope
        public static List<OutputInfo> outputs = new List<OutputInfo>();

        public static VIDIO_parameter vidio = new VIDIO_parameter();
        public static VIDI2C_parameter vidi2c = new VIDI2C_parameter();
        public static LoadTransient_parameter loadtransient = new LoadTransient_parameter();
        public static LTLab_parameter lt_lab = new LTLab_parameter();


    }
}
