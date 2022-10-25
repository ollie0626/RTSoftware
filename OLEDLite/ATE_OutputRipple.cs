using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace OLEDLite
{
    public class ATE_OutputRipple : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        //Excel.Range _range;
        string eLoadInfo = "";
        string SwireInfo = "";
        RTBBControl RTDev = new RTBBControl();

        public delegate void FinishNotification();

        FinishNotification delegate_mess;


        public ATE_OutputRipple()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Output ripple test finished!!!", "OLED-Lit Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }


        private void OSCInint()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();

            InsControl._scope.CH1_On(); // Lx
            InsControl._scope.CH2_On(); // output

            InsControl._scope.CH2_ACoupling();
            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();

            InsControl._scope.CH1_Level(6);
            InsControl._scope.CH2_Level(1);

        }

        private void OSCReset()
        {
            // Ch1 measure Lx
            InsControl._scope.DoCommand(":MEASure:PERiod CHANnel1");
            InsControl._scope.DoCommand(":MEASure:STATistics MAX");
            string[] res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
            double period_max = Convert.ToDouble(res[0]);

            InsControl._scope.TimeScaleUs(100);
            double unit = Math.Pow(10, -6);
            //double period = InsControl._scope.Meas_CH1Period();
            double time_scale = (period_max * 3) / 10;
            if (period_max > 9.99 * Math.Pow(10, 10)) time_scale = 100;

            while (period_max > 9.99 * Math.Pow(10, 10))
            {
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                period_max = Convert.ToDouble(res[0]);
                //period_max = InsControl._scope.Meas_CH1Period();
                InsControl._scope.TimeScale(time_scale * unit);
                time_scale--;
                if (time_scale < 10) break;
            }

            InsControl._scope.NormalTrigger();
            InsControl._scope.Root_Clear();
            InsControl._scope.Root_Single();
            for (int i = 0; i < 10; i++)
            {
                InsControl._scope.Root_Single();
                MyLib.Delay1ms(150);
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                MyLib.Delay1ms(50);
                period_max = Convert.ToDouble(res[0]);
                time_scale = (period_max * 4) / 10;
                InsControl._scope.TimeScale(time_scale);
                MyLib.Delay1ms(150);
            }
            InsControl._scope.Root_RUN();
        }

        public override void ATETask()
        {
            // board and timer initial
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();

            // variable declare
            int idx = 0;
            bool ccm_enable = false;    
            int bin_cnt = 1;
            int wave_idx = 0;
            int row = 11;
            string res = "";
            string SwireInfo = "";

            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length == 0 ? 1 : binList.Length;
            double[] ori_vinTable = new double[test_parameter.vinList.Count];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, test_parameter.vinList.Count);
            InsControl._power.AutoSelPowerOn(test_parameter.vinList[0]);

            OSCInint();
            for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
            {
                ccm_enable = false;
                for (int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++)
                {

                    SwireInfo = test_parameter.i2c_enable ? binList[interface_idx] : "Swire=" + test_parameter.swireList[interface_idx];

                    for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;

                        if (test_parameter.i2c_enable)
                        {
                            res = Path.GetFileNameWithoutExtension(binList[interface_idx]);
                        }
                        else
                        {
                            res = SwireInfo;
                        }


                        if ((interface_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        string file_name = string.Format("{0}_Temp={1}C_Vin={2:0.0#}V_{3:0.0#}A_{4}",
                                                            idx,
                                                            temp,
                                                            test_parameter.vinList[vin_idx],
                                                            test_parameter.ioutList[iout_idx],
                                                            res // i2c or swire interface
                                                        );


                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        MyLib.EloadFixChannel();
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);

                        if (test_parameter.i2c_enable)
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


                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("34970 沒有連結 !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }

                        if (!ccm_enable) InsControl._scope.TimeScaleUs(50);
                        double threshold = 9.99 * Math.Pow(10, 20);
                        double burst_period = InsControl._scope.Meas_BurstPeriod(1, test_parameter.burst_period);
                        if (burst_period < threshold)
                        {
                            // set trigger
                            InsControl._scope.TimeScaleUs(50);
                            InsControl._scope.Root_Clear();
                            InsControl._scope.AutoTrigger();
                            // adjust ch1 level
                            InsControl._scope.CH1_Level(2.5);
                            double trigger_level = InsControl._scope.Meas_CH1VPP();
                            double vol_min = InsControl._scope.Measure_Ch_Min(1);
                            if (vol_min < -2) InsControl._scope.TriggerLevel_CH1(0);
                            else InsControl._scope.CH1_Level(trigger_level / 3);
                            // pulse skip mode
                            for (int i = 0; i < 3; i++)
                            {
                                InsControl._scope.Root_Single();
                                MyLib.Delay1ms(500);
                                burst_period = InsControl._scope.Meas_BurstPeriod(1, test_parameter.burst_period);
                               
                                InsControl._scope.TimeScale(burst_period);
                                MyLib.Delay1ms(250);
                            }
                            InsControl._scope.Root_RUN();
                        }
                        else if(!ccm_enable)
                        {
                            // set trigger
                            InsControl._scope.TimeScaleUs(50);
                            InsControl._scope.Root_Clear();
                            InsControl._scope.AutoTrigger();
                            // adjust ch1 level
                            InsControl._scope.CH1_Level(2.5);
                            double trigger_level = InsControl._scope.Meas_CH1VPP();
                            double vol_min = InsControl._scope.Measure_Ch_Min(1);
                            if (vol_min < -2) InsControl._scope.TriggerLevel_CH1(0);
                            else InsControl._scope.CH1_Level(trigger_level / 3);
                            // time scale calculate in CCM mode
                            OSCReset();
                            ccm_enable = true;
                        }


                        MyLib.Delay1ms(250);
                        MyLib.Channel_LevelSetting(1);
                        MyLib.Channel_LevelSetting(2);
                        MyLib.Delay1ms(250);
                        // scope open rgb color function
                        //InsControl._scope.DoCommand(":DISPlay:PERSistence 5");
                        MyLib.Delay1s(5);
                        double max, min, vpp, vin, vout, iin, iout;
                        // save waveform
                        InsControl._scope.Root_STOP();

                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);
                        // measure data
                        max = InsControl._scope.Meas_CH1MAX() * 1000;
                        min = InsControl._scope.Meas_CH1MIN() * 1000;
                        vpp = InsControl._scope.Meas_CH1VPP() * 1000;
                        vin = InsControl._34970A.Get_100Vol(1);
                        vout = InsControl._34970A.Get_100Vol(2);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();


                        MyLib.Delay1ms(500);
                        InsControl._eload.CH1_Loading(0);
                        InsControl._eload.AllChannel_LoadOff();
                        if (Math.Abs(vout) < 0.15)
                        {
                            InsControl._power.AutoPowerOff();
                            System.Threading.Thread.Sleep(500);
                            InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                            InsControl._scope.CH1_Level(1);
                            System.Threading.Thread.Sleep(250);
                        }
                        InsControl._scope.Root_RUN();
                        row++; idx++;
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();


            delegate_mess.Invoke();

        }
    }
}
