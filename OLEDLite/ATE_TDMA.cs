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
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);

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

            InsControl._scope.DoCommand(":MEASure:STATistics MEAN");
            string[] res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');

            // measure real VinHi, VinLo
            meas_VinHi = Convert.ToDouble(res[0]);
            meas_VinLo = Convert.ToDouble(res[2]);

            
            while (meas_VinHi < (VinH - margin) || meas_VinHi > (VinH + margin))
            {
                InsControl._scope.Root_Clear();
                InsControl._scope.Root_RUN();
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                meas_VinHi = Convert.ToDouble(res[0]);

                if (meas_VinHi < (VinH - margin))
                {
                    VinH += Offset;
                    MyLib.FuncGen_loopparameter(VinH, VinL);
                }

                if(meas_VinHi < (VinH + margin))
                {
                    VinH -= Offset;
                    MyLib.FuncGen_loopparameter(VinH, VinL);
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
                meas_VinLo = Convert.ToDouble(res[2]);

                if (meas_VinLo < (VinL - margin))
                {
                    VinL += Offset;
                    MyLib.FuncGen_loopparameter(VinH, VinL);
                }

                if (meas_VinLo < (VinL + margin))
                {
                    VinL -= Offset;
                    MyLib.FuncGen_loopparameter(VinH, VinL);
                }
                // out of loop
                VinL_cnt++;
                if (VinL_cnt > 40) break;
            }
        }

        private void VoResize()
        {
            double Vo_offset = InsControl._scope.Meas_CH2AVE();

            InsControl._scope.CH2_Offset(Vo_offset);
            InsControl._scope.CH2_Level(0.1); // 100mV
            MyLib.WaveformCheck();
            InsControl._scope.CH2_Offset(Vo_offset);
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

            OSCInint();
            for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++) // functino gen vin 
            {
                for(int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? test_parameter.swireList.Count : bin_cnt); interface_idx++) // interface
                {
                    OSCRest();
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

                    // power off
                    InsControl._funcgen.CH1_Off();
                }
            }
        }



    }
}
