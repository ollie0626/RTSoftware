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
        string VinInfo = "";
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

            InsControl._scope.CH1_Level(2);
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
                MyLib.Delay1ms(300);
            }
            InsControl._scope.Root_RUN();
        }

        private void ExcelInit()
        {
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Name = "Output ripple";
            // for iout
        }

        public override void ATETask()
        {
            // board and timer initial
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();
            RTDev.BoadInit();

            // variable declare
            int idx = 0;
            bool ccm_enable = false;
            int bin_cnt = 1;
            int wave_idx = 0;
            int row = 11;
            string res = "";
            //string SwireInfo = "";

            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length == 0 ? 1 : binList.Length;
            double[] ori_vinTable = new double[test_parameter.vinList.Count];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, test_parameter.vinList.Count);
            InsControl._power.AutoSelPowerOn(test_parameter.vinList[0]);

            OSCInint();
            for (int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++)
            {

#if Report
                row = 11;
                ExcelInit();
#endif
                
                for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
                {
#if Report
                    _sheet.Cells[row, XLS_Table.P] = "VIN(V)";
                    _sheet.Cells[row, XLS_Table.Q] = "IIN(mA)";
                    _sheet.Cells[row, XLS_Table.R] = "VO(V)";
                    _sheet.Cells[row, XLS_Table.S] = "IO(mA)";
                    _sheet.Cells[row, XLS_Table.T] = "Reverse";
                    _sheet.Cells[row, XLS_Table.U] = "Fluctuation(mV)";
                    _sheet.Cells[row, XLS_Table.V] = "Vpp(mV)";
                    _sheet.Cells[row + 1, XLS_Table.P] = "VIN=" + test_parameter.vinList[vin_idx] + "V";
                    _sheet.Cells[1, 1] = "Vin:";
                    _sheet.Cells[2, 1] = "Iout:";
                    _sheet.Cells[3, 1] = "setting conditions:";
                    _sheet.Cells[4, 1] = "Note:";
                    _sheet.Cells[5, 1] = "Date:";

                    VinInfo = "Vin=";
                    VinInfo += test_parameter.vinList[0] + "V ~ "
                            + test_parameter.vinList[test_parameter.vinList.Count - 1] + "V\r\n";

                    eLoadInfo += test_parameter.ioutList[0] + "mA ~" 
                            + test_parameter.ioutList[test_parameter.ioutList.Count - 1] + "mA\r\n";
                    SwireInfo = test_parameter.i2c_enable ? binList[interface_idx] : "Swire=" + test_parameter.swireList[interface_idx];
                    _sheet.Cells[1, 2] = VinInfo;
                    _sheet.Cells[2, 2] = eLoadInfo;
                    _sheet.Cells[3, 2] = (test_parameter.i2c_enable) ? Path.GetFileNameWithoutExtension(binList[interface_idx]) : SwireInfo;
                    _sheet.Cells[4, 2] = (test_parameter.i2c_enable) ? "" : test_parameter.swire_20 ? "ASwire=1, ESwire=0" : "ASwire=0, ESwire=1";
                    _sheet.Cells[5, 2] = DateTime.Now.ToString("yyyyMMdd");
#endif
                    ccm_enable = false;
                    
                    row+=2;
                    start_pos.Add(row);
                    for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    {
                        InsControl._scope.Measure_Clear();
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

                        double Vtop = InsControl._scope.Measure_Top(1);
                        double Vbase = InsControl._scope.Measure_Base(1);

                        if (Vtop < 0) InsControl._scope.TriggerLevel_CH1(0);
                        else InsControl._scope.TriggerLevel_CH1(Vtop * 0.4);

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
                            else InsControl._scope.TriggerLevel_CH1(trigger_level / 3);
                            // pulse skip mode
                            for (int i = 0; i < 5; i++)
                            {
                                InsControl._scope.Root_Single();
                                MyLib.Delay1ms(250);
                                burst_period = InsControl._scope.Meas_BurstPeriod(1, test_parameter.burst_period);

                                if(burst_period < (2.5 * Math.Pow(10, -6))) InsControl._scope.TimeScaleUs(10);
                                else InsControl._scope.TimeScale(burst_period);
                                MyLib.Delay1ms(250);
                            }
                            InsControl._scope.Root_RUN();

                            
                        }
                        else if (!ccm_enable)
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
                        InsControl._scope.DoCommand(":MEASure:PERiod CHANnel1");
                        InsControl._scope.DoCommand(":MEASure:VPP CHANnel2");
                        InsControl._scope.DoCommand(":MEASure:VMIN CHANnel2");
                        InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2");
                        MyLib.Delay1s(1);
                        for (int k = 0; k < 20; k++)
                        {
                            InsControl._scope.Root_Single();
                            MyLib.Delay1ms(150);
                        }

                        double max, min, vpp, vin, vout, iin, iout;
                        double fluctulation;
                        // save waveform
                        InsControl._scope.Root_STOP();

                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);
                        // measure data
                        string[] meas_res;
                        InsControl._scope.DoCommand(":MEASure:STATistics MAX");
                        meas_res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                        max = Convert.ToDouble(meas_res[0]) * 1000;
                        InsControl._scope.DoCommand(":MEASure:STATistics MIN");
                        meas_res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                        min = Convert.ToDouble(meas_res[0]) * 1000;
                        fluctulation = max - min;
                        vpp = InsControl._scope.Meas_CH2VPP() * 1000;
                        vin = InsControl._34970A.Get_100Vol(1);
                        vout = InsControl._34970A.Get_100Vol(2);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();

#if Report
                        _sheet.Cells[row, XLS_Table.P] = string.Format("{0:0.###}", vin);
                        _sheet.Cells[row, XLS_Table.Q] = string.Format("{0:0.###}", iin * 1000);
                        _sheet.Cells[row, XLS_Table.R] = string.Format("{0:0.###}", vout);
                        _sheet.Cells[row, XLS_Table.S] = string.Format("{0:0.###}", iout * 1000);
                        _sheet.Cells[row, XLS_Table.T] = "";
                        _sheet.Cells[row, XLS_Table.U] = string.Format("{0:0.###}", fluctulation);
                        _sheet.Cells[row, XLS_Table.V] = string.Format("{0:0.###}", vpp);
                        if (ccm_enable) _sheet.Cells[row, XLS_Table.W] = "DCM/CCM";
                        else _sheet.Cells[row, XLS_Table.W] = "PSM";
#endif
                        MyLib.Delay1ms(500);
                        InsControl._eload.CH1_Loading(0);
                        InsControl._eload.AllChannel_LoadOff();
                        if (Math.Abs(vout) < 0.15)
                        {
                            InsControl._power.AutoPowerOff();
                            System.Threading.Thread.Sleep(500);
                            InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                            //InsControl._scope.CH1_Level(1);
                            System.Threading.Thread.Sleep(250);
                        }
                        InsControl._scope.Root_RUN();
                        row++; idx++;
                    } // eload loop
                    stop_pos.Add(row - 1);
                } // vin loop

#if Report
                TimeSpan timeSpan = stopWatch.Elapsed;
                AddCruve(start_pos, stop_pos);
                string conditions = eLoadInfo == "" ? "" : eLoadInfo.Replace("\r\n", "") + "_";
                MyLib.SaveExcelReport(test_parameter.wave_path, "Temp=" + temp + "C_Ripple&Flu_" + conditions + SwireInfo.Replace("\r\n", "") + "_" + DateTime.Now.ToString("yyyyMMdd"), _book);
                _book.Close(false);
                _book = null;
                _app.Quit();
                _app = null;
                GC.Collect();
#endif
            } // interface loop

        Stop:
            stopWatch.Stop();
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            delegate_mess.Invoke();
        }


        private void AddCruve(List<int> start_pos, List<int> stop_pos)
        {
            Excel.Chart fluc_chart, vpp_chart;
            Excel.Range range;
            Excel.SeriesCollection fluc_collect, vpp_collect;
            Excel.Series fluc_series, vpp_series;
            Excel.Range XRange, YRange;

            range = _sheet.Range["A12", "G29"];
            fluc_chart = MyLib.CreateChart(_sheet, range, "Fluctuation @" + SwireInfo.Replace("\r\n", ""), "Load (mA)", "Fluctuation(mV)", true);
            fluc_chart.ChartTitle.Font.Size = 14;
            fluc_collect = fluc_chart.SeriesCollection();

            range = _sheet.Range["A32", "G49"];
            vpp_chart = MyLib.CreateChart(_sheet, range, "Vpp @" + SwireInfo.Replace("\r\n", ""), "Load (mA)", "Vpp (mV)", true);
            vpp_chart.ChartTitle.Font.Size = 14;
            vpp_collect = vpp_chart.SeriesCollection();

            for (int line = 0; line < start_pos.Count; line++)
            {
                fluc_series = fluc_collect.NewSeries();
                XRange = _sheet.Range["S" + start_pos[line], "S" + stop_pos[line]];
                YRange = _sheet.Range["U" + start_pos[line], "U" + stop_pos[line]];
                fluc_series.XValues = XRange;
                fluc_series.Values = YRange;
                fluc_series.Name = _sheet.Cells[start_pos[line] - 1, XLS_Table.P].Value.ToString();

                vpp_series = vpp_collect.NewSeries();
                YRange = _sheet.Range["V" + start_pos[line], "V" + stop_pos[line]];
                vpp_series.XValues = XRange;
                vpp_series.Values = YRange;
                vpp_series.Name = _sheet.Cells[start_pos[line] - 1, XLS_Table.P].Value.ToString();
            }
        }
    }
}
