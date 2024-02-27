using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using NLua;

namespace InsLibDotNet
{
    public class LuaBase
    {
        public Lua L = new Lua();
    }

    public class LuaRegister : LuaBase
    {

        //static AgilentOSC _scope = new AgilentOSC();

        public void LuaInsLib_Init()
        {
            /* Agilent Function register on Lua script */
            //AgilentOSC_Reg();
            Console.WriteLine("luanet test");
        }

        //private void AgilentOSC_Reg()
        //{
        //    List<string> Func_List = new List<string>();

        //    Func_List.Add("ConnectOscilloscope");
        //    Func_List.Add("AgilentOSC_RST");
        //    Func_List.Add("Root_STOP");
        //    Func_List.Add("Root_RUN");
        //    Func_List.Add("Root_Clear");

        //    Func_List.Add("DoCommand");
        //    Func_List.Add("DoQueryNumber");
        //    Func_List.Add("DoQueryString");

        //    Func_List.Add("AutoTrigger");
        //    Func_List.Add("NormalTrigger");
        //    Func_List.Add("SingleTrigger");
        //    Func_List.Add("TimeBasePosition");              // input position
        //    Func_List.Add("TimeScale");                     // input scale
        //    Func_List.Add("TimeScaleMs");                   // input timeScale_ms
        //    Func_List.Add("TimeScaleUs");                   // input timeScale_us
        //    Func_List.Add("SaveWaveform");                  // input Path & FileName

        //    Func_List.Add("SystemPresetDefault");
        //    Func_List.Add("SetTrigModeEdge");               // input Negative
        //    Func_List.Add("SweepModeTrig");
        //    Func_List.Add("SweepModeAuto");
        //    Func_List.Add("MeasureStatisticsMean");
        //    Func_List.Add("MeasureStatisticsMin");
        //    Func_List.Add("MeasureStatisticsMax");
        //    Func_List.Add("MeasureStatisticsCurrent");
        //    Func_List.Add("GetMeasureStatistics");

        //    Func_List.Add("CH1_10_90Range");
        //    Func_List.Add("CH2_10_90Range");
        //    Func_List.Add("CH3_10_90Range");
        //    Func_List.Add("CH4_10_90Range");

        //    Func_List.Add("CH1_50ohm");
        //    Func_List.Add("CH2_50ohm");
        //    Func_List.Add("CH3_50ohm");
        //    Func_List.Add("CH4_50ohm");

        //    Func_List.Add("CH1_1Mohm");
        //    Func_List.Add("CH2_1Mohm");
        //    Func_List.Add("CH3_1Mohm");
        //    Func_List.Add("CH4_1Mohm");

        //    Func_List.Add("CH1_On");
        //    Func_List.Add("CH2_On");
        //    Func_List.Add("CH3_On");
        //    Func_List.Add("CH4_On");

        //    Func_List.Add("CH1_Off");
        //    Func_List.Add("CH2_Off");
        //    Func_List.Add("CH3_Off");
        //    Func_List.Add("CH4_Off");

        //    Func_List.Add("CH1_Offset");                    // input offset
        //    Func_List.Add("CH2_Offset");                    // input offset
        //    Func_List.Add("CH3_Offset");                    // input offset
        //    Func_List.Add("CH4_Offset");                    // input offset

        //    Func_List.Add("Trigger_CH1");
        //    Func_List.Add("Trigger_CH2");
        //    Func_List.Add("Trigger_CH3");
        //    Func_List.Add("Trigger_CH4");

        //    Func_List.Add("CH1_DCoupling");
        //    Func_List.Add("CH2_DCoupling");
        //    Func_List.Add("CH3_DCoupling");
        //    Func_List.Add("CH4_DCoupling");

        //    Func_List.Add("CH1_ACoupling");
        //    Func_List.Add("CH2_ACoupling");
        //    Func_List.Add("CH3_ACoupling");
        //    Func_List.Add("CH4_ACoupling");

        //    Func_List.Add("TriggerLevel_CH1");
        //    Func_List.Add("TriggerLevel_CH2");
        //    Func_List.Add("TriggerLevel_CH3");
        //    Func_List.Add("TriggerLevel_CH4");

        //    Func_List.Add("CH1_BWLimitOn");
        //    Func_List.Add("CH2_BWLimitOn");
        //    Func_List.Add("CH3_BWLimitOn");
        //    Func_List.Add("CH4_BWLimitOn");

        //    Func_List.Add("CH1_BWLimitOff");
        //    Func_List.Add("CH2_BWLimitOff");
        //    Func_List.Add("CH3_BWLimitOff");
        //    Func_List.Add("CH4_BWLimitOff");

        //    Func_List.Add("CH1_Level");                    // input level
        //    Func_List.Add("CH2_Level");                    // input level
        //    Func_List.Add("CH3_Level");                    // input level
        //    Func_List.Add("CH4_Level");                    // input level


        //    /* Measure Function List ---------------------------------------------------------------- */
        //    Func_List.Add("MeasDelta");                    // input src1, src2
        //    Func_List.Add("Measure_Clear");
        //    Func_List.Add("Meas_Result");

        //    // Rise
        //    Func_List.Add("Meas_CH1Rise");
        //    Func_List.Add("Meas_CH2Rise");
        //    Func_List.Add("Meas_CH3Rise");
        //    Func_List.Add("Meas_CH4Rise");
        //    // Fall
        //    Func_List.Add("Meas_CH1Fall");
        //    Func_List.Add("Meas_CH2Fall");
        //    Func_List.Add("Meas_CH3Fall");
        //    Func_List.Add("Meas_CH4Fall");
        //    // Top
        //    Func_List.Add("Meas_CH1Top");
        //    Func_List.Add("Meas_CH2Top");
        //    Func_List.Add("Meas_CH3Top");
        //    Func_List.Add("Meas_CH4Top");
        //    // Base
        //    Func_List.Add("Meas_CH1Base");
        //    Func_List.Add("Meas_CH2Base");
        //    Func_List.Add("Meas_CH3Base");
        //    Func_List.Add("Meas_CH4Base");
        //    // Freq
        //    Func_List.Add("Meas_CH1Freq");
        //    Func_List.Add("Meas_CH2Freq");
        //    Func_List.Add("Meas_CH3Freq");
        //    Func_List.Add("Meas_CH4Freq");
        //    // Period
        //    Func_List.Add("Meas_CH1Period");
        //    Func_List.Add("Meas_CH2Period");
        //    Func_List.Add("Meas_CH3Period");
        //    Func_List.Add("Meas_CH4Period");
        //    // Max
        //    Func_List.Add("Meas_CH1MAX");
        //    Func_List.Add("Meas_CH2MAX");
        //    Func_List.Add("Meas_CH3MAX");
        //    Func_List.Add("Meas_CH4MAX");
        //    // Min
        //    Func_List.Add("Meas_CH1MIN");
        //    Func_List.Add("Meas_CH2MIN");
        //    Func_List.Add("Meas_CH3MIN");
        //    Func_List.Add("Meas_CH4MIN");
        //    // XDelta 
        //    Func_List.Add("Meas_CH1XDelta");
        //    Func_List.Add("Meas_CH2XDelta");
        //    Func_List.Add("Meas_CH3XDelta");
        //    Func_List.Add("Meas_CH4XDelta");
        //    // VPP
        //    Func_List.Add("Meas_CH1VPP");
        //    Func_List.Add("Meas_CH2VPP");
        //    Func_List.Add("Meas_CH3VPP");
        //    Func_List.Add("Meas_CH4VPP");
        //    // Avg
        //    Func_List.Add("Meas_CH1AVE");
        //    Func_List.Add("Meas_CH2AVE");
        //    Func_List.Add("Meas_CH3AVE");
        //    Func_List.Add("Meas_CH4AVE");
        //    // Pwidth
        //    Func_List.Add("Meas_CH1PWidth");
        //    Func_List.Add("Meas_CH2PWidth");
        //    Func_List.Add("Meas_CH3PWidth");
        //    Func_List.Add("Meas_CH4PWidth");
        //    // Nwidth
        //    Func_List.Add("Meas_CH1NWidth");
        //    Func_List.Add("Meas_CH2NWidth");
        //    Func_List.Add("Meas_CH3NWidth");
        //    Func_List.Add("Meas_CH4NWidth");


        //    foreach (string func in Func_List)
        //    {
        //        L.RegisterFunction("lua_" + func, _scope, _scope.GetType().GetMethod(func));
        //    }
        //}


    }







}
