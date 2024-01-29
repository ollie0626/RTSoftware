using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InsLibDotNet;


namespace SoftStartTiming
{

    public static class InsControl
    {
        public static AgilentOSC _scope;
        public static PowerModule _power;
        public static PowerModule _power2;
        public static EloadModule _eload;
        public static MultiChannelModule _34970A;
        public static ChamberModule _chamber;
        public static FuncGenModule _funcgen;

        public static TekTronix7Serise _tek_scope;
        public static bool _tek_scope_en;


        public static  OscilloscopesModule _oscilloscope;

    }
}
