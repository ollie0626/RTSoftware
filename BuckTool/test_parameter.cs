using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BuckTool
{
    public static class test_parameter
    {
        public static List<double> Iout_table = new List<double>();
        public static List<double> Vin_table = new List<double>();
        public static List<string> temp_table = new List<string>();

        // High Voltage buck frequency control
        public static bool[] Freq_en = new bool[2];
        
        // bin folder path
        public static string specify_bin;
        public static string waveform_path;

        public static bool run_stop;
        public static bool chamber_en;

    }
}
