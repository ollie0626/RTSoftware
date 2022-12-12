using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;
using System.IO;

namespace BuckTool
{
    public class ATE_Lx : TaskRun
    {

        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        private const string str_Freqfolder = "\\Freq";
        private const string str_SRfolder = "\\SR";
        private const string str_Jitterfolder = "\\Jitter";

        private void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_Off();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_Off();
            InsControl._scope.CH1_Level(6);
            InsControl._scope.CH1_Offset(6 * 1.5);
            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
        }


        private void WaveformPathCheck()
        {
            if (!Directory.Exists(test_parameter.waveform_path + "\\Lx"))
            {
                Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx");
                Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_Freqfolder);
                Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_SRfolder);
                Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_Jitterfolder);
            }
            else
            {
                if (!Directory.Exists(test_parameter.waveform_path + "\\Lx" + str_Freqfolder)) Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_Freqfolder);
                if (!Directory.Exists(test_parameter.waveform_path + "\\Lx" + str_SRfolder)) Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_SRfolder);
                if (!Directory.Exists(test_parameter.waveform_path + "\\Lx" + str_Jitterfolder)) Directory.CreateDirectory(test_parameter.waveform_path + "\\Lx" + str_Jitterfolder);
            }
            test_parameter.waveform_path = test_parameter.waveform_path + "\\Lx";
        }

        public override void ATETask()
        {

            WaveformPathCheck();
            int freq_cnt = (test_parameter.Freq_en[0] ? 1 : 0) + (test_parameter.Freq_en[1] ? 1 : 0);
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
            Mylib.testCondition(_sheet, "Lx", 0, temp);
#endif
            OSCInit();

            for (int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                for(int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                {
#if Report
                    printTitle(row); row++;
#endif
                    for (int iout_idx = 0; iout_idx < test_parameter.Iout_table.Count; iout_idx++)
                    {
                        string file_neme = "";
                        string file_name = string.Format("{0}_Vin={1}_Iout={2}_Freq={3}",
                                                        row - 22,
                                                        test_parameter.Vin_table[vin_idx],
                                                        test_parameter.Iout_table[iout_idx],
                                                        (freq_idx == 0 && test_parameter.Freq_en[0]) ? test_parameter.Freq_des[0] : test_parameter.Freq_des[1]);
                        double iout = test_parameter.Iout_table[iout_idx];
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((iout_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        MyLib.Switch_ELoadLevel(iout);

                        InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);
                        MyLib.Delay1ms(200);
                        InsControl._eload.CH1_Loading(test_parameter.Iout_table[iout_idx]);

                        double vin, iin, vout;
                        vin = InsControl._power.GetVoltage();
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();
                        vout = InsControl._eload.GetVol();

#if Report
                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = vin;
                        _sheet.Cells[row, XLS_Table.D] = iin;
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
                        _sheet.Cells[row, XLS_Table.F] = iout;
                        _sheet.Cells[row, XLS_Table.G] = vout;
#endif

                        for (int item = 0; item < 3; item++)
                        {
                            switch (item)
                            {
                                case 0: // freq task
                                    FreqTask(row); row++;
                                    InsControl._scope.SaveWaveform(test_parameter.waveform_path + str_Freqfolder, file_neme + "_freq");
                                    InsControl._scope.Root_RUN();
                                    break;
                                case 1: // jitter task
                                    JitterTask(row); row++;
                                    InsControl._scope.SaveWaveform(test_parameter.waveform_path + str_Jitterfolder, file_neme + "_jitter");
                                    InsControl._scope.DoCommand(":HISTogram:MODE OFF");
                                    InsControl._scope.DoCommand(":DISPlay:CGRade OFF");
                                    InsControl._scope.Root_RUN();
                                    break;
                                case 2: // slew rate task
                                    SlewRateTask(row, file_neme); row++;
                                    InsControl._scope.Root_RUN();
                                    break;
                            }
                        }

                        InsControl._power.AutoPowerOff();
                        InsControl._eload.AllChannel_LoadOff();
                        InsControl._scope.AutoTrigger();
                        MyLib.WaveformCheck();
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

            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_Lx" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        } // ATETask

        private void SlewRateTask(int row, string file_name)
        {
            double period, trigger, lx_level;
            InsControl._scope.Measure_Clear();
            InsControl._scope.Measure_Freq(1);
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.Bandwidth_Limit_On(1);
            InsControl._scope.Ch_On(1);
            InsControl._scope.Ch_Off(2);
            InsControl._scope.Ch_Off(3);
            InsControl._scope.Ch_Off(4);
            InsControl._scope.TimeScaleUs(20);
            InsControl._scope.TimeBasePosition(0);

            trigger = InsControl._scope.Meas_CH1VPP() / 3;
            lx_level = InsControl._scope.Meas_CH1VPP() / 3;
            InsControl._scope.SetTrigModeEdge(false);
            InsControl._scope.TriggerLevel_CH1(trigger);
            InsControl._scope.CH1_Level(lx_level);

            period = InsControl._scope.Meas_CH1Period();
            period = period / 10; // show 1 cycle
            InsControl._scope.TimeScale(period);
            InsControl._scope.NormalTrigger();
            MyLib.WaveformCheck();

            InsControl._scope.DoCommand(":MEASure:THResholds:METHod ALL,PERCent");
            InsControl._scope.DoCommand(":MEASure:THResholds:GEN:PERCent ALL,80,50,20");
            // Rise
            InsControl._scope.DoCommand(":MEASure:SLEWrate CHANnel1, RISing");
            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.DoCommand(":MARKer:MODE ON");
            InsControl._scope.DoCommand(":MEASURE:RISetime CHANnel1");
            InsControl._scope.SetTrigModeEdge(false); // trigger rise
            double rise_time = InsControl._scope.Measure_SlewRate_Rising(1);
            double rise = InsControl._scope.Measure_Rise(1);
            InsControl._scope.TimeScale(rise * 2);
            MyLib.WaveformCheck();
            InsControl._scope.Root_STOP();
            rise_time = InsControl._scope.Measure_SlewRate_Rising(1);
            rise = InsControl._scope.Measure_Rise(1);
#if Report
            _sheet.Cells[row, XLS_Table.K] = string.Format("{0:0.000}", rise_time * Math.Pow(10, 9));
            _sheet.Cells[row, XLS_Table.L] = string.Format("{0:0.000}", rise * Math.Pow(10, -6));
            InsControl._scope.Bandwidth_Limit_Off(1);
            _sheet.Cells[row, XLS_Table.R] = InsControl._scope.Measure_Ch_Max(1);
            _sheet.Cells[row, XLS_Table.S] = InsControl._scope.Measure_Ch_Min(1);
#endif
            InsControl._scope.TimeScale(rise * 2);
            // Rise save waveform
            InsControl._scope.SaveWaveform(test_parameter.waveform_path + str_SRfolder, file_name + "_Rise");

            // Fall
            InsControl._scope.Bandwidth_Limit_On(1);
            InsControl._scope.Measure_Clear();
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.Root_RUN();
            InsControl._scope.DoCommand(":MEASure:SLEWrate CHANnel1, Falling");
            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.DoCommand(":MARKer:MODE ON");
            InsControl._scope.DoCommand(":MEASURE:FALLtime CHANnel1");
            InsControl._scope.SetTrigModeEdge(true);
            MyLib.WaveformCheck();
            InsControl._scope.Root_STOP();
            double fall = InsControl._scope.Measure_SlewRate_Falling(1);
            double fall_time = InsControl._scope.Measure_Fall_Time(1);
#if Report
            _sheet.Cells[row, XLS_Table.M] = string.Format("{0:0.000}", fall_time * Math.Pow(10, 9));
            _sheet.Cells[row, XLS_Table.N] = string.Format("{0:0.000}", fall * Math.Pow(10, -6));
            InsControl._scope.Bandwidth_Limit_Off(1);
            _sheet.Cells[row, XLS_Table.T] = InsControl._scope.Measure_Ch_Max(1);
            _sheet.Cells[row, XLS_Table.U] = InsControl._scope.Measure_Ch_Min(1);
#endif
            // fall wave waveform
            InsControl._scope.SaveWaveform(test_parameter.waveform_path + str_SRfolder, file_name + "_Fall");
        }

        private void JitterTask(int row)
        {
            double period, trigger, lx_level;
            double Rlimit, Llimit, Vtop, Vbase, histogramLevel;
            InsControl._scope.Measure_Clear();
            InsControl._scope.Measure_Freq(1);
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.Bandwidth_Limit_On(1);
            InsControl._scope.Ch_On(1);
            InsControl._scope.Ch_Off(2);
            InsControl._scope.Ch_Off(3);
            InsControl._scope.Ch_Off(4);
            InsControl._scope.TimeScaleUs(20);
            InsControl._scope.TimeBasePosition(0);

            period = InsControl._scope.Meas_CH1Period();
            period = period / 6.5; // default 10 period
            InsControl._scope.TimeScale(period);
            InsControl._scope.TimeBasePosition(period * 2);
            Rlimit = (period * 5);
            Llimit = (period * 0.2);
            trigger = InsControl._scope.Meas_CH1VPP() / 3;
            lx_level = InsControl._scope.Meas_CH1VPP() / 3;
            Vtop = InsControl._scope.Meas_CH1Top();
            Vbase = InsControl._scope.Meas_CH1Base();
            histogramLevel = Vtop * 0.5 + Vbase * 0.5;
            InsControl._scope.SetTrigModeEdge(true);
            InsControl._scope.TriggerLevel_CH1(trigger);
            InsControl._scope.CH1_Level(lx_level);
            InsControl._scope.NormalTrigger();
            InsControl._scope.DoCommand(":HISTogram:MODE OFF");
            InsControl._scope.DoCommand(":DISPlay:CGRade 1");
            InsControl._scope.DoCommand(":HISTogram:SCALe:SIZE 2");
            InsControl._scope.DoCommand(":HISTogram:MODE WAVeform");
            InsControl._scope.DoCommand(":HISTogram:WINDow:SOURce CHANnel1");
            InsControl._scope.DoCommand(":HISTogram:WINDow:LLIMit " + Llimit);
            InsControl._scope.DoCommand(":HISTogram:WINDow:RLIMit " + Rlimit);
            InsControl._scope.DoCommand(":HISTogram:WINDow:TLIMit " + (histogramLevel * 1.05));
            InsControl._scope.DoCommand(":HISTogram:WINDow:BLIMit " + (histogramLevel * 0.95));
            MyLib.WaveformCheck();
            MyLib.Delay1s(6);
            InsControl._scope.Root_STOP();

            double MeaPKPK = InsControl._scope.doQueryNumber(":MEASure:HISTogram:PP?") * Math.Pow(10, 9);
            double MeaMean = InsControl._scope.doQueryNumber(":MEASure:HISTogram:PP?");
            double MeaStdDev = InsControl._scope.doQueryNumber(":MEASure:HISTogram:STDDev?") * Math.Pow(10, 9);
            double Freq = InsControl._scope.Measure_Freq(1);

#if Report
            _sheet.Cells[row, XLS_Table.O] = MeaPKPK;
            _sheet.Cells[row, XLS_Table.P] = MeaStdDev;
            _sheet.Cells[row, XLS_Table.Q] = MeaPKPK * Freq * 100 * Math.Pow(10, -9);
#endif
        }

        private void FreqTask(int row)
        {
            double period, trigger, lx_level;
            double freq_mean, freq_max, freq_min;
            InsControl._scope.Measure_Clear();
            InsControl._scope.Measure_Freq(1);
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.Bandwidth_Limit_On(1);
            InsControl._scope.Ch_On(1);
            InsControl._scope.Ch_Off(2);
            InsControl._scope.Ch_Off(3);
            InsControl._scope.Ch_Off(4);
            InsControl._scope.TimeScaleUs(20);
            InsControl._scope.TimeBasePosition(0);

            period = InsControl._scope.Meas_CH1Period();
            period = period / 2; // show 5 cycle
            InsControl._scope.TimeScale(period);
            trigger = InsControl._scope.Meas_CH1VPP() / 3;
            lx_level = InsControl._scope.Meas_CH1VPP() / 3;
            InsControl._scope.SetTrigModeEdge(false);
            InsControl._scope.TriggerLevel_CH1(trigger);
            InsControl._scope.CH1_Level(lx_level);
            InsControl._scope.NormalTrigger();
            MyLib.WaveformCheck();
            InsControl._scope.Root_STOP();

            InsControl._scope.DoCommand(":MEASure:SENDvalid 1");
            InsControl._scope.DoCommand(":MEASURE:STATISTICS MAX ");
            InsControl._scope.DoCommand(":MEASure:Freq CHANnel1");
            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.DoCommand(":MARKer:MODE ON");
            ////:MARKer:MEASurement:MEASurement {MEASurement<N>}
            string[] res = InsControl._scope.GetMeasureStatistics().Split(',');

            InsControl._scope.DoCommand(":MEASURE:STATISTICS MEAN");
            freq_mean = InsControl._scope.Meas_Result();
            InsControl._scope.DoCommand(":MEASURE:STATISTICS MAX");
            freq_max = InsControl._scope.Meas_Result();
            InsControl._scope.DoCommand(":MEASURE:STATISTICS MIN");
            freq_min = InsControl._scope.Meas_Result();

#if Report
            _sheet.Cells[row, XLS_Table.H] = string.Format("{0:##.000}", freq_mean);
            _sheet.Cells[row, XLS_Table.I] = string.Format("{0:##.000}", freq_max);
            _sheet.Cells[row, XLS_Table.J] = string.Format("{0:##.000}", freq_min);
#endif

            MyLib.Delay1ms(200);
        }

        private void printTitle(int row)
        {
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Freq(MHz)";
            _sheet.Cells[row, XLS_Table.F] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.G] = "Vout";

            _sheet.Cells[row, XLS_Table.H] = "Freq(KHz)";
            _sheet.Cells[row, XLS_Table.I] = "Freq Max(KHz)";
            _sheet.Cells[row, XLS_Table.J] = "Freq Min(KHz)";
            _sheet.Cells[row, XLS_Table.K] = "Rise Time(ns)";
            _sheet.Cells[row, XLS_Table.L] = "Rise SR(V/us)";
            _sheet.Cells[row, XLS_Table.M] = "Fall Time(ns)";
            _sheet.Cells[row, XLS_Table.N] = "Fall SR(V/us)";
            _sheet.Cells[row, XLS_Table.O] = "Jitter(ns)";
            _sheet.Cells[row, XLS_Table.P] = "Std Dev(ns)";
            _sheet.Cells[row, XLS_Table.Q] = "Jitter(%)";
            _sheet.Cells[row, XLS_Table.R] = "Rise Max(V)";
            _sheet.Cells[row, XLS_Table.S] = "Rise min(V)";
            _sheet.Cells[row, XLS_Table.T] = "Fall Max(V)";
            _sheet.Cells[row, XLS_Table.U] = "Fall min(V)";

            _range = _sheet.Range["A" + row, "G" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["H" + row, "U" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
        }

    }
}
