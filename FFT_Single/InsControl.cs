﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InsLibDotNet;


namespace FFT_Single
{

    public static class InsControl
    {
        public static AgilentOSC _scope;
        public static PowerModule _power;
        public static EloadModule _eload;
        public static MultiChannelModule _34970A;
        public static ChamberModule _chamber;
        public static DMMModule _dmm1;
        public static DMMModule _dmm2;
        public static FuncGenModule _funcgen;
    }
}