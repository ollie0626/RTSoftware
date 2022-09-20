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
        public static string binFolder;
        public static List<string> swireList = new List<string>();

        // ELoad
        public static List<double> ioutList = new List<double>();
        
        // FuncGen
        public static List<Hi_Lo> HiLo_table = new List<Hi_Lo>();
        public static double Freq;
        public static double duty;
        public static double tr;
        public static double tf;



    }
}
