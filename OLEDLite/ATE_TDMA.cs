﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;


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
        
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        //Excel.Range _range;
        string eLoadInfo = "";
        string SwireInfo = "";
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
            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandCenter\"");
        }

        private void OSCRest()
        {
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);
            MyLib.WaveformCheck();
            InsControl._scope.DoCommand(":MEASure:RISEtime CHANnel1"); // rising
            InsControl._scope.DoCommand(":MEASure:FALLtime CHANnel1"); // falling
            InsControl._scope.DoCommand(":MEASure:VAVG CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMIN CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VBASE CHANnel1");
            InsControl._scope.DoCommand(":MEASure:VTOP CHANnel1");
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
            InsControl._scope.CH1_Offset(VinH * 0.95);
            InsControl._scope.DoCommand(":MEASure:STATistics MEAN");
            MyLib.Delay1s(1);
            InsControl._scope.TimeScale((1 / (test_parameter.Freq * 1000)) / 10);
            string[] res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
            InsControl._scope.Trigger_CH1();
            InsControl._scope.TriggerLevel_CH1((VinH + VinL) / 2);
            InsControl._scope.SetTriggerMode("TIMeout");
            InsControl._scope.SetTimeoutCondition(true);
            InsControl._scope.SetTimeoutSource(1);
            InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.5) * Math.Pow(10, 9));

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

                if (meas_VinHi > (VinH + margin))
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

            for (int i = 0; i < 3; i++)
            {
                Vo_offset = InsControl._scope.Meas_CH2AVG();
                InsControl._scope.CH2_Offset(Vo_offset);
            }

            double abs_vo = Math.Abs(Vo_offset);
            if (abs_vo < 5) InsControl._scope.CH2_Level(0.01);
            else if (5 < abs_vo && abs_vo < 10) InsControl._scope.CH2_Level(0.02);
            else InsControl._scope.CH2_Level(0.1);
        }

        private void ExcelInit()
        {
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Name = "TDMA";

            _sheet.Cells[10, XLS_Table.U] = "Ideal";
            _sheet.Cells[11, XLS_Table.U] = "Iout (mA)";
            _sheet.Cells[11, XLS_Table.V] = "Overshoot (mV)";
            _sheet.Cells[11, XLS_Table.W] = "Undershoot (mV)"; 
            _sheet.Cells[11, XLS_Table.X] = "VPP (mV)";
        }

        public override void ATETask()
        {

            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();

            RTDev.BoadInit();
            int bin_cnt = 1;
            int wave_idx = 0;
            int row = 11;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length == 0 ? 1 : binList.Length;
            InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[0].Highlevel + 0.5);
            MyLib.FuncGen_Fixedparameter(test_parameter.Freq * 1000,
                                         test_parameter.duty,
                                         test_parameter.tr,
                                         test_parameter.tf);

            OSCInint();

            for (int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swire_cnt); interface_idx++) // interface
            {
#if Report
                Stopwatch stopWatch = new Stopwatch();
                stopWatch.Start();
                row = 11;
                ExcelInit();
#endif
                for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++) // functino gen vin 
                {
#if Report
                    _sheet.Cells[row - 1, XLS_Table.K] = "VIN=" + test_parameter.HiLo_table[func_idx].LowLevel.ToString().Replace(".", "")
                                            + test_parameter.HiLo_table[func_idx].Highlevel.ToString().Replace(".", "");
                    _sheet.Cells[row, XLS_Table.K] = "No.";
                    _sheet.Cells[row, XLS_Table.L] = "Temp(C)";
                    _sheet.Cells[row, XLS_Table.M] = "Vin_L(V)";
                    _sheet.Cells[row, XLS_Table.N] = "Vin_H(V)";
                    _sheet.Cells[row, XLS_Table.O] = "Iout(mA)";
                    _sheet.Cells[row, XLS_Table.P] = "Overshoot(mV)";
                    _sheet.Cells[row, XLS_Table.Q] = "Undershoot(mV)";
                    _sheet.Cells[row, XLS_Table.R] = "VPP(mV)";
                    row++;
                    _sheet.Cells[row, XLS_Table.K] = "ESwire:" + test_parameter.ESwireList[interface_idx] + "_ASwire:" + test_parameter.ASwireList[interface_idx];
                    row++;

                    _sheet.Cells[1, XLS_Table.A] = "Vin:";
                    _sheet.Cells[2, XLS_Table.A] = "Iout:";
                    _sheet.Cells[3, XLS_Table.A] = "Date:";
                    _sheet.Cells[4, XLS_Table.A] = "Note:";
                    _sheet.Cells[5, XLS_Table.A] = "Version";
                    _sheet.Cells[6, XLS_Table.A] = "Temperature";
                    _sheet.Cells[7, XLS_Table.A] = "test time";

                    string res = "";
                    for (int i = 0; i < test_parameter.HiLo_table.Count; i++)
                        res += test_parameter.HiLo_table[i].Highlevel + "->" + test_parameter.HiLo_table[i].LowLevel + ", ";
                    _sheet.Cells[1, XLS_Table.B] = res;
                    _sheet.Cells[2, XLS_Table.B] = test_parameter.eload_info;
                    _sheet.Cells[3, XLS_Table.B] = test_parameter.date_info;
                    _sheet.Cells[5, XLS_Table.B] = test_parameter.ver_info;
                    _sheet.Cells[6, XLS_Table.B] = temp;

                    int idx = 0;
                    eLoadInfo = "";
                    for (int i = 0; i < test_parameter.eload_en.Length; i++)
                    {
                        if (test_parameter.eload_en[i])
                        {
                            _sheet.Cells[row, XLS_Table.S + idx++] = "ELoad" + (i + 1).ToString() + "(mA)";
                            if (eLoadInfo == "")
                            {
                                eLoadInfo = "Wi Load" + (i + 1).ToString() + "=" + test_parameter.eload_iout[i] * 1000 + "mA";
                            }
                            else
                            {
                                eLoadInfo += "Wi Load" + (i + 1).ToString() + "=" + test_parameter.eload_iout[i] * 1000 + "mA";
                            }
                        }
                    }
#endif
                    row++;
                    start_pos.Add(row);
                    for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    {
                        //InsControl._scope.Measure_Clear();
                        if (test_parameter.run_stop == true) goto Stop;


                        if(test_parameter.i2c_enable)
                        {
                            res = Path.GetFileNameWithoutExtension(binList[interface_idx]);
                        }
                        else
                        {
                            res = SwireInfo;
                        }
                        
                        string file_name = string.Format("{0}_Temp={1}_{2}_Vin={3:0.##}V_{4:0.##}V_iout={5:0.##}mA",
                                                        wave_idx, temp, res,
                                                        test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel,
                                                        test_parameter.ioutList[iout_idx] * 1000);
                        if ((func_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();


                        OSCRest();
                        InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[func_idx].Highlevel + 0.5);
                        MyLib.FuncGen_loopparameter(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);

                        if (test_parameter.i2c_enable)
                        {
                            // i2c interface
                            RTDev.I2C_WriteBin(test_parameter.slave, 0x00, binList[interface_idx]);
                        }
                        else
                        {
                            // ic setting
                            int[] pulse_tmp;
                            bool[] Enable_state_table = new bool[] { test_parameter.ESwire_state, test_parameter.ASwire_state, test_parameter.ENVO4_state };
                            int[] Enable_num_table = new int[] { RTBBControl.ESwire, RTBBControl.ASwire, RTBBControl.ENVO4 };
                            pulse_tmp = test_parameter.ESwireList[interface_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(true, pulse_tmp[pulse_idx]);

                            pulse_tmp = test_parameter.ASwireList[interface_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(false, pulse_tmp[pulse_idx]);

                            for (int i = 0; i < Enable_state_table.Length; i++) RTBBControl.Swire_Control(Enable_num_table[i], Enable_state_table[i]);
                        }

                        MyLib.EloadFixChannel();
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);

                        ViResize(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);
                        VoResize();
                        MyLib.Delay1ms(300);
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);
                        InsControl._scope.DoCommand(":MEASure:STATistics MEAN");
                        string[] HiLo_res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');

                        // measure part
                        double zoomout_peak = InsControl._scope.Meas_CH2MAX();
                        double zoomout_neg_peak = InsControl._scope.Meas_CH2MIN();
                        double vpp = InsControl._scope.Meas_CH2VPP();
                        double on_time = (1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100);
                        double off_time = (1 / (test_parameter.Freq * 1000)) * ((100 - test_parameter.duty) / 100);

                        InsControl._scope.TimeScale(on_time / 20);
                        MyLib.Delay1ms(250);
                        double hi_peak = InsControl._scope.Meas_CH2MAX();
                        double hi_neg_peak = InsControl._scope.Meas_CH2MIN();


                        InsControl._scope.TimeScale(on_time / 20);
                        InsControl._scope.TimeBasePosition(on_time);
                        MyLib.Delay1ms(250);
                        double lo_peak = InsControl._scope.Meas_CH2MAX();
                        double lo_neg_peak = InsControl._scope.Meas_CH2MIN();


                        InsControl._scope.TimeScale(on_time / 7);
                        InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.2) * Math.Pow(10, 9));
                        InsControl._scope.TimeBasePosition(0);
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_Rising");

                        InsControl._scope.SetTimeoutCondition(false);
                        InsControl._scope.TimeScale(off_time / 7);
                        InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * ((100 - test_parameter.duty) / 100) * 0.2) * Math.Pow(10, 9));
                        InsControl._scope.TimeBasePosition(0);
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_Falling");

                        // power off
                        InsControl._funcgen.CH1_Off();
                        InsControl._scope.TimeBasePosition(0);

                        // report
                        double[] overshoot_list = new double[] { hi_peak, lo_peak };
                        double[] undershoot_list = new double[] { hi_neg_peak, lo_neg_peak };

#if Report
                        _sheet.Cells[row, XLS_Table.K] = wave_idx;
                        _sheet.Cells[row, XLS_Table.L] = temp;
                        _sheet.Cells[row, XLS_Table.M] = Convert.ToDouble(HiLo_res[1]);
                        _sheet.Cells[row, XLS_Table.N] = Convert.ToDouble(HiLo_res[0]);
                        _sheet.Cells[row, XLS_Table.O] = test_parameter.ioutList[iout_idx] * 1000;
                        _sheet.Cells[row, XLS_Table.P] = Math.Abs(zoomout_peak - overshoot_list.Max()) * 1000;
                        _sheet.Cells[row, XLS_Table.Q] = Math.Abs(zoomout_neg_peak - undershoot_list.Min()) * 1000;
                        _sheet.Cells[row, XLS_Table.R] = vpp * 1000;
#endif
                        row++;
                        wave_idx++;
                    } // Eload loop
                    stop_pos.Add(row - 1);
                    row += 2;
                } // Func loop
#if Report
                stopWatch.Stop();
                TimeSpan timeSpan = stopWatch.Elapsed;
                string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
                _sheet.Cells[7, XLS_Table.B] = time;
                AddCruve(start_pos, stop_pos);
                string conditions = eLoadInfo == "" ? "" : eLoadInfo + "_";
                MyLib.SaveExcelReport(test_parameter.wave_path, "Temp=" + temp + "_TDMA Data Collection_" + conditions + SwireInfo  + "_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
                _book.Close(false);
                _book = null;
                _app.Quit();
                _app = null;
                GC.Collect();
#endif
            } // interface loop

        Stop:
            if(!test_parameter.chamber_en)
                System.Windows.Forms.MessageBox.Show("Test finished!!!", "OLED Lite", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void AddCruve(List<int> start_pos, List<int> stop_pos)
        {
#if Report
            Excel.Chart chart, chart_lor, chart_vpp;
            Excel.Range range;
            Excel.SeriesCollection collection, collection_lor, collection_vpp;
            Excel.Series series, series_lor, series_vpp;
            Excel.Range XRange, YRange;
            range = _sheet.Range["A16", "I32"];
            chart = MyLib.CreateChart(_sheet, range, "TDMA Data Collection @" + SwireInfo , "Load (mA) " + eLoadInfo, "Overshoot(mV)");

            // for LOR
            range = _sheet.Range["A38", "I54"];
            chart_lor = MyLib.CreateChart(_sheet, range, "TDMA Data Collection @" + SwireInfo, "Load (mA) " + eLoadInfo, "Undershoot(mV)");

            range = _sheet.Range["A60", "I76"];
            chart_vpp = MyLib.CreateChart(_sheet, range, "TDMA Data Collection @" + SwireInfo, "Load (mA) " + eLoadInfo, "VPP(mV)");

            chart.ChartTitle.Font.Size = 14;
            chart_lor.ChartTitle.Font.Size = 14;
            chart_vpp.ChartTitle.Font.Size = 14;

            collection = chart.SeriesCollection();
            collection_lor = chart_lor.SeriesCollection();
            collection_vpp = chart_vpp.SeriesCollection();

            for (int line = 0; line < start_pos.Count; line++)
            {
                series = collection.NewSeries();

                XRange = _sheet.Range["O" + start_pos[line].ToString(), "O" + stop_pos[line].ToString()];
                YRange = _sheet.Range["P" + start_pos[line].ToString(), "P" + stop_pos[line].ToString()];
                series.XValues = XRange;
                series.Values = YRange;
                series.Name = _sheet.Cells[start_pos[line] - 2, XLS_Table.K].Value.ToString();

                series_lor = collection_lor.NewSeries();
                YRange = _sheet.Range["Q" + start_pos[line].ToString(), "Q" + stop_pos[line].ToString()];
                series_lor.XValues = XRange;
                series_lor.Values = YRange;
                series_lor.Name = _sheet.Cells[start_pos[line] - 2, XLS_Table.K].Value.ToString();

                series_vpp = collection_vpp.NewSeries();
                YRange = _sheet.Range["R" + start_pos[line].ToString(), "R" + stop_pos[line].ToString()];
                series_vpp.XValues = XRange;
                series_vpp.Values = YRange;
                series_vpp.Name = _sheet.Cells[start_pos[line] - 2, XLS_Table.K].Value.ToString();
            }

            series = collection.NewSeries();
            series.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
            XRange = _sheet.Range["U" + start_pos[0].ToString(), "U" + stop_pos[0].ToString()];
            YRange = _sheet.Range["V" + start_pos[0].ToString(), "V" + stop_pos[0].ToString()];
            series.XValues = XRange;
            series.Values = YRange;
            series.Name = _sheet.Cells[start_pos[0] - 2, XLS_Table.U].Value.ToString();

            series_lor = collection_lor.NewSeries();
            series_lor.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
            YRange = _sheet.Range["W" + start_pos[0].ToString(), "W" + stop_pos[0].ToString()];
            series_lor.XValues = XRange;
            series_lor.Values = YRange;
            series_lor.Name = _sheet.Cells[start_pos[0] - 2, XLS_Table.U].Value.ToString();

            series_vpp = collection_vpp.NewSeries();
            series_vpp.ChartType = Excel.XlChartType.xlXYScatterLinesNoMarkers;
            YRange = _sheet.Range["X" + start_pos[0].ToString(), "X" + stop_pos[0].ToString()];
            series_vpp.XValues = XRange;
            series_vpp.Values = YRange;
            series_vpp.Name = _sheet.Cells[start_pos[0] - 2, XLS_Table.K].Value.ToString();
#endif
        }
    }
}
