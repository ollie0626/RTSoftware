using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;


namespace OLEDLite
{

    public interface ITask
    {
        void ATETask();
    }

    public class TaskRun : ITask
    {
        public double temp = 25;
        virtual public void ATETask()
        { }
    }

    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
    };

    public class ATE_TDMA : TaskRun
    {
        //Excel.Application _app;
        //Excel.Worksheet _sheet;
        //Excel.Workbook _book;
        //Excel.Range _range;
        RTBBControl RTDev = new RTBBControl();

        private void OSCInint()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();

            InsControl._scope.TimeScale((1 / (test_parameter.Freq * 1000)) / 10);
            
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();

            InsControl._scope.DoCommand(":MEASure:VAVG CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMIN CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VBASE CHANnel1");
            InsControl._scope.DoCommand(":MEASure:VTOP CHANnel1");
        }

        private void OSCRest()
        {
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);
            MyLib.WaveformCheck();
        }


        // VinH : Hi target
        // VinL : Lo target
        private void ViResize(double VinH, double VinL)
        {
            double margin = 0.02;
            double Offset = 0.005;
            double meas_VinHi, meas_VinLo;
            int VinH_cnt = 0, VinL_cnt = 0;

            InsControl._scope.Root_RUN();
            InsControl._scope.Root_Clear();
            MyLib.Delay1s(1);
            InsControl._scope.Root_RUN();
            InsControl._scope.CH1_Level((VinH - VinL) / 3);
            InsControl._scope.CH1_Offset(VinH);
            InsControl._scope.DoCommand(":MEASure:STATistics MEAN");
            string[] res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
            InsControl._scope.Trigger_CH1();
            InsControl._scope.TriggerLevel_CH1((VinH + VinL) / 2);
            InsControl._scope.SetTriggerMode("TIMeout");
            InsControl._scope.SetTimeoutCondition(true);
            InsControl._scope.SetTimeoutSource(1);
            InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.5) * Math.Pow(10,9));

            // measure real VinHi, VinLo
            meas_VinHi = Convert.ToDouble(res[0]);
            meas_VinLo = Convert.ToDouble(res[1]);
            double VinH_in = VinH, VinL_in = VinL;

            
            while (meas_VinHi < (VinH - margin) || meas_VinHi > (VinH + margin))
            {
                InsControl._scope.Root_Clear();
                InsControl._scope.Root_RUN();
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                meas_VinHi = Convert.ToDouble(res[0]);

                if (meas_VinHi < (VinH - margin))
                {
                    VinH_in += Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL);
                }

                if(meas_VinHi > (VinH + margin))
                {
                    VinH_in -= Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL);
                }
                // out of loop
                VinH_cnt++;
                if (VinH_cnt > 40) break;
            }


            while (meas_VinLo < (VinL - margin) || meas_VinLo > (VinL + margin))
            {
                InsControl._scope.Root_Clear();
                InsControl._scope.Root_RUN();
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                meas_VinLo = Convert.ToDouble(res[1]);

                if (meas_VinLo < (VinL - margin))
                {
                    VinL_in += Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL_in);
                }

                if (meas_VinLo > (VinL + margin))
                {
                    VinL_in -= Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL_in);
                }
                // out of loop
                VinL_cnt++;
                if (VinL_cnt > 40) break;
            }
        }

        private void VoResize()
        {
            double Vo_offset = InsControl._scope.Meas_CH2AVG();

            InsControl._scope.CH2_Offset(Vo_offset);
            InsControl._scope.CH2_Level(0.1); // 100mV
            MyLib.WaveformCheck();

            for(int i = 0; i < 3; i++)
            {
                Vo_offset = InsControl._scope.Meas_CH2AVG();
                InsControl._scope.CH2_Offset(Vo_offset);
            }

            double abs_vo = Math.Abs(Vo_offset);
            if (abs_vo < 5) InsControl._scope.CH2_Level(0.01);
            else if (5 < abs_vo && abs_vo < 10) InsControl._scope.CH2_Level(0.02);
            else InsControl._scope.CH2_Level(0.1);
        }

        public override void ATETask()
        {
            RTDev.BoadInit();
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;

            InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[0].Highlevel + 0.5);
            MyLib.FuncGen_Fixedparameter(test_parameter.Freq * 1000,
                                         test_parameter.duty,
                                         test_parameter.tr,
                                         test_parameter.tf);
            OSCInint();
            for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++) // functino gen vin 
            {
                for(int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++) // interface
                {
                    OSCRest();
                    InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[func_idx].Highlevel + 0.5);
                    MyLib.FuncGen_loopparameter(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);

                    if(test_parameter.i2c_enable)
                    {
                        // i2c interface
                        RTDev.I2C_WriteBin(test_parameter.slave, 0x00, binList[interface_idx]);
                    }
                    else
                    {
                        // swire
                        int[] pulse_tmp = test_parameter.swireList[interface_idx].Split(',').Select(int.Parse).ToArray();
                        for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTDev.SwirePulse(pulse_tmp[pulse_idx]);
                    }

                    MyLib.EloadFixChannel();
                    ViResize(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);
                    VoResize();
                    // save waveform full case
                    InsControl._scope.SaveWaveform(test_parameter.wave_path, "full");


                    // measure part
                    double zoomout_peak = InsControl._scope.Meas_CH2MAX();
                    double zoomout_neg_peak = InsControl._scope.Meas_CH2MIN();
                    double on_time = (1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100);
                    double off_time = (1 / (test_parameter.Freq * 1000)) * ((100 - test_parameter.duty) / 100);

                    InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.8) * Math.Pow(10, 9));
                    InsControl._scope.TimeScale(on_time / 20);
                    InsControl._scope.TimeBasePosition(on_time + ((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.8));
                    double hi_peak = InsControl._scope.Meas_CH2MAX();
                    double hi_neg_peak = InsControl._scope.Meas_CH2MIN();

                    //InsControl._scope.SetTrigModeEdge(true)
                    InsControl._scope.SetTimeoutCondition(false);
                    InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * ((100 - test_parameter.duty) / 100) * 0.8) * Math.Pow(10, 9));
                    InsControl._scope.TimeScale(off_time / 20);
                    InsControl._scope.TimeBasePosition(off_time + ((1 / (test_parameter.Freq * 1000) * ((100 - test_parameter.duty) / 100)) * 0.8));
                    double lo_peak = InsControl._scope.Meas_CH2MAX();
                    double lo_neg_peak = InsControl._scope.Meas_CH2MIN();


                    // save waveform rising case
                    double rising_trigger_time = on_time * 0.2 * Math.Pow(10, 9);
                    double rising_scale = on_time / 20;
                    InsControl._scope.SetTimeoutCondition(true);
                    InsControl._scope.SetTimeoutTime(rising_trigger_time);
                    InsControl._scope.TimeScale(rising_scale);
                    InsControl._scope.TimeBasePosition(rising_scale * -3.5);
                    InsControl._scope.SaveWaveform(test_parameter.wave_path, "rising");

                    // save waveform falling case 
                    double falling_trigger_time = off_time * 0.2 * Math.Pow(10, 9);
                    double falling_scale = off_time / 20;
                    InsControl._scope.SetTimeoutCondition(false);
                    InsControl._scope.SetTimeoutTime(falling_trigger_time);
                    InsControl._scope.TimeScale(falling_scale);
                    InsControl._scope.TimeBasePosition(falling_scale * -3.5);
                    InsControl._scope.SaveWaveform(test_parameter.wave_path, "falling");


                    // power off
                    //InsControl._scope.SetTrigModeEdge(false);
                    InsControl._scope.TimeBasePosition(0);
                    InsControl._scope.TimeScale((1 / (test_parameter.Freq * 1000)) / 10);
                    InsControl._funcgen.CH1_Off();
                }
            }
        }



    }
}
