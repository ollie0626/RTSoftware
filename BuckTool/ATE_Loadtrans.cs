using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;

namespace BuckTool
{
    public class ATE_Loadtrans: TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        private void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_Off();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_On();

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();
            InsControl._scope.CH1_ACoupling();
            MyLib.WaveformCheck();

            InsControl._scope.Trigger_CH4();
            InsControl._scope.TriggerLevel_CH4(0.2);
            InsControl._scope.CH1_Level(1);
            InsControl._scope.CH4_Level(0.5);
            InsControl._scope.CH1_Offset(-2);
            InsControl._scope.CH4_Offset(1.5);
            InsControl._scope.TimeScale(1 / test_parameter.freq);
            MyLib.WaveformCheck();

            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
        }

        public override void ATETask()
        {
            int freq_cnt = (test_parameter.Freq_en[0] ? 1 : 0) + (test_parameter.Freq_en[1] ? 1 : 0);
            double period = 1 / test_parameter.freq;


            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int row = 22;
            //int bin_cnt = 1;
            MyLib Mylib = new MyLib();
            //string[] binList = new string[1];
            //binList = Mylib.ListBinFile(test_parameter.binFolder);
            //bin_cnt = binList.Length;
            //double[] vinList = new double[test_parameter.Vin_table.Count];
            //Array.Copy(vinList, test_parameter.Vin_table.ToArray(), vinList.Length);

            double[] vinList = test_parameter.Vin_table.ToArray();
#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "LoadTrans", 0, temp);
#endif

            OSCInit();
            MyLib.FuncGen_Fixedparameter(
                                        test_parameter.freq,
                                        test_parameter.duty,
                                        test_parameter.tr,
                                        test_parameter.tf);

            double time_scale = ((1 / test_parameter.freq) * (test_parameter.duty / 100)) / 5;
            InsControl._scope.TimeScale(time_scale);
            InsControl._scope.TimeBasePosition(time_scale * 2.5);


            for (int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();
                for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                {
#if Report
                    printTitle(row); row++;
#endif
                    InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);
                    for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++)
                    {
                        string file_name = string.Format("{0}_Vin={1}V_Freq={2}_Hi={3}V_Lo={4}V",
                                                        row - 22,
                                                        test_parameter.Vin_table[vin_idx],
                                                        (freq_idx == 0 && test_parameter.Freq_en[0]) ? test_parameter.Freq_des[0] : test_parameter.Freq_des[1],
                                                        test_parameter.HiLo_table[func_idx].Highlevel,
                                                        test_parameter.HiLo_table[func_idx].LowLevel);

                        double current_level, trigger_level;
                        double vpp, vmax, vmin, rise, fall, rise_time, fall_time;
                        double imax, imin, overshoot, undershoot;
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((func_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                        InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);

                        MyLib.FuncGen_loopparameter(
                            test_parameter.HiLo_table[func_idx].Highlevel,
                            test_parameter.HiLo_table[func_idx].LowLevel);

                        time_scale = ((1 / test_parameter.freq) * (test_parameter.duty / 100)) / 5;
                        InsControl._scope.TimeScale(time_scale);
                        InsControl._scope.TimeBasePosition(time_scale * 2.5);

                        InsControl._scope.AutoTrigger();
                        current_level = (test_parameter.HiLo_table[func_idx].Highlevel + test_parameter.HiLo_table[func_idx].LowLevel) / 4;
                        trigger_level = test_parameter.HiLo_table[func_idx].Highlevel * 0.6 + test_parameter.HiLo_table[func_idx].LowLevel * 0.4;
                        InsControl._scope.CH4_Level(current_level);
                        InsControl._scope.CH4_Offset(current_level * 3);
                        MyLib.WaveformCheck();
                        InsControl._scope.TriggerLevel_CH4(trigger_level);
                        MyLib.WaveformCheck();
                        InsControl._scope.SetTrigModeEdge(false);
                        InsControl._scope.Root_STOP();
                        InsControl._scope.NormalTrigger();
                        InsControl._scope.DoCommand(":MEASure:CLEar");
                        InsControl._scope.DoCommand(":MEASURE:VPP CHANnel1");
                        InsControl._scope.DoCommand(":MEASURE:VMAX CHANnel1");
                        InsControl._scope.DoCommand(":MEASURE:VMIN CHANnel1");
                        InsControl._scope.DoCommand(":MEASURE:VMAX CHANnel4");
                        InsControl._scope.DoCommand(":MEASURE:VMIN CHANnel4");
                        //MyLib.ProcessCheck();
                        InsControl._scope.CH1_Level(0.3);
                        InsControl._scope.Root_RUN();
                        //MyLib.WaveformCheck();
                        ChannelResize();
                        //MyLib.WaveformCheck();
                        InsControl._scope.Root_STOP();
                        MyLib.Delay1s(1);
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);

                        vpp = InsControl._scope.Meas_CH1VPP();
                        vmax = InsControl._scope.Meas_CH1MAX();
                        vmin = InsControl._scope.Meas_CH1MIN();
                        rise = InsControl._scope.doQueryNumber(":MEASure:SLEWrate? CHANnel4,RISing") / 1000;
                        fall = InsControl._scope.doQueryNumber(":MEASure:SLEWrate? CHANnel4,Falling") / 1000;
                        rise_time = InsControl._scope.Meas_CH4Rise() * Math.Pow(10, 6);
                        fall_time = InsControl._scope.Meas_CH4Fall() * Math.Pow(10, 6);
                        imax = InsControl._scope.Meas_CH4MAX();
                        imin = InsControl._scope.Meas_CH4MIN();
#if Report
                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.Vin_table[vin_idx];
                        if (freq_cnt == 1)
                        {
                            if (test_parameter.Freq_en[0])
                                _sheet.Cells[row, XLS_Table.E] = test_parameter.Freq_des[0];
                            else
                                _sheet.Cells[row, XLS_Table.E] = test_parameter.Freq_des[1];
                        }
                        else
                        {
                            _sheet.Cells[row, XLS_Table.E] = test_parameter.Freq_des[freq_idx];
                        }
                        _sheet.Cells[row, XLS_Table.E] = test_parameter.freq;
                        _sheet.Cells[row, XLS_Table.F] = test_parameter.HiLo_table[func_idx].Highlevel;
                        _sheet.Cells[row, XLS_Table.G] = test_parameter.HiLo_table[func_idx].LowLevel;
                        _sheet.Cells[row, XLS_Table.H] = vpp * 1000;
                        _sheet.Cells[row, XLS_Table.I] = vmin * 1000;
                        _sheet.Cells[row, XLS_Table.J] = vmax * 1000;
                        _sheet.Cells[row, XLS_Table.K] = rise;
                        _sheet.Cells[row, XLS_Table.L] = rise_time;
                        _sheet.Cells[row, XLS_Table.M] = fall;
                        _sheet.Cells[row, XLS_Table.N] = fall_time;

                        _sheet.Cells[row, XLS_Table.O] = imax * 1000;
                        _sheet.Cells[row, XLS_Table.P] = imin * 1000;
                        _sheet.Cells[row, XLS_Table.Q] = Math.Abs(vmax / test_parameter.vout_ideal) * 100;
                        _sheet.Cells[row, XLS_Table.R] = Math.Abs(vmin / test_parameter.vout_ideal) * 100;


#endif

                        for (int i = 0; i < 2; i++)
                        {
                            InsControl._scope.TimeBasePosition(0);
                            switch (i)
                            {
                                case 0: // rise
                                    InsControl._scope.Root_RUN();
                                    InsControl._scope.SetTrigModeEdge(false);
                                    InsControl._scope.TimeScaleUs(test_parameter.tr * 3);
                                    InsControl._scope.TimeBasePositionUs(test_parameter.tr * 9);
                                    //MyLib.WaveformCheck();
                                    InsControl._scope.Root_STOP();
                                    MyLib.Delay1s(1);
                                    InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_Rise");
                                    break;
                                case 1: // fall
                                    InsControl._scope.Root_RUN();
                                    InsControl._scope.SetTrigModeEdge(true);
                                    InsControl._scope.TimeScaleUs(test_parameter.tf * 3);
                                    InsControl._scope.TimeBasePositionUs(test_parameter.tf * 9);
                                    //MyLib.WaveformCheck();
                                    InsControl._scope.Root_STOP();
                                    MyLib.Delay1s(1);
                                    InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_Fall");
                                    break;
                            }
                        }

                        InsControl._scope.Root_RUN();
                        InsControl._scope.AutoTrigger();
                        MyLib.WaveformCheck();
                        InsControl._power.AutoPowerOff();
                        row++;
                    } // iout loop
                } // vin loop
            } // freq loop



            Stop:
            stopWatch.Stop();
#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_LoadTrans" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        } // ATETask

        private void ChannelResize()
        {
            InsControl._scope.CH1_Level(0.3);
            double max = InsControl._scope.Meas_CH1VPP();
            for(int i = 0; i < 3; i++)
            {
                InsControl._scope.CH1_Level(max / 3);
                max = InsControl._scope.Meas_CH1VPP();
                MyLib.ProcessCheck();
            }

            //max = InsControl._scope.Meas_CH1VPP();
            //for (int i = 0; i < 3; i++)
            //{
            //    InsControl._scope.CH1_Level(max / 3);
            //    max = InsControl._scope.Meas_CH1MAX();
            //    MyLib.ProcessCheck();
            //}
        }

        private void printTitle(int row)
        {
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Freq(MHz)";
            _sheet.Cells[row, XLS_Table.E] = "Freq(KHz)";
            _sheet.Cells[row, XLS_Table.F] = "Hi Level(V)";
            _sheet.Cells[row, XLS_Table.G] = "Lo Level(V)";
            _sheet.Cells[row, XLS_Table.H] = "Vpp(mV)";
            _sheet.Cells[row, XLS_Table.I] = "Vmin(mV)";
            _sheet.Cells[row, XLS_Table.J] = "Vmax(mV)";
            _sheet.Cells[row, XLS_Table.K] = "Rise SR(A/ms)";
            _sheet.Cells[row, XLS_Table.L] = "Rise Time(us)";
            _sheet.Cells[row, XLS_Table.M] = "Fall SR(A/ms)";
            _sheet.Cells[row, XLS_Table.N] = "Fall Time(us)";
            _sheet.Cells[row, XLS_Table.O] = "Imax(mA)";
            _sheet.Cells[row, XLS_Table.P] = "Imin(mA)";
            _sheet.Cells[row, XLS_Table.Q] = "Overshoot(%)";
            _sheet.Cells[row, XLS_Table.R] = "UnderShoot(%)";

            _range = _sheet.Range["A" + row, "G" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["H" + row, "R" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
        }

    }
}
