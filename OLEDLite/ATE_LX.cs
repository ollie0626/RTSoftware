using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;

namespace OLEDLite
{
    public class ATE_LX : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        RTBBControl RTDev = new RTBBControl();

        private void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();
            InsControl._scope.CH1_On();
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH1_Offset(5 * 1.5);
            InsControl._scope.CH1_BWLimitOn();
        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();

            int row = 11;
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            int vin_cnt = test_parameter.vinList.Count;
            int iout_cnt = test_parameter.ioutList.Count;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);
            RTDev.BoadInit();

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;

            _sheet.Cells[1, XLS_Table.A] = "Vin";
            _sheet.Cells[2, XLS_Table.A] = "Iout";
            _sheet.Cells[3, XLS_Table.A] = "Date";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _sheet.Cells[5, XLS_Table.A] = "Version";
            _sheet.Cells[6, XLS_Table.A] = "Temperature";
            _sheet.Cells[7, XLS_Table.A] = "test time";

            _sheet.Cells[1, XLS_Table.B] = test_parameter.vin_info;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.eload_info;
            _sheet.Cells[3, XLS_Table.B] = test_parameter.date_info;
            _sheet.Cells[5, XLS_Table.B] = test_parameter.ver_info;
            _sheet.Cells[6, XLS_Table.B] = temp;
#endif
            InsControl._power.AutoPowerOff();
            OSCInit();
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
#if Report
                _sheet.Cells[row, XLS_Table.A] = "Temp(C)";
                _sheet.Cells[row, XLS_Table.B] = "超連結";
                _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
                _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
                _sheet.Cells[row, XLS_Table.E] = "Swire";
                _sheet.Cells[row, XLS_Table.F] = "Iload(mA)";
                _sheet.Cells[row, XLS_Table.G] = "Vout(V)";

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
                _range = _sheet.Range["A" + row, "Q" + row];
                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                _sheet.Range["A" + row, "G" + row].Interior.Color = Color.FromArgb(0xff, 0xff, 0x66);
                _sheet.Range["H" + row, "J" + row].Interior.Color = Color.FromArgb(0, 204, 0);
                _sheet.Range["K" + row, "N" + row].Interior.Color = Color.FromArgb(0xB2, 0xFF, 0x66);
                _sheet.Range["O" + row, "Q" + row].Interior.Color = Color.FromArgb(0, 204, 0);
                row++;
#endif

                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        System.Threading.Thread.Sleep(500);
                        InsControl._eload.DoCommand(InsControl._eload.CH1);
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);
                        MyLib.EloadFixChannel();

                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("Please connect DAQ !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }
                        if (binList[0] != "" && test_parameter.i2c_enable) RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                        else
                        {
                            // ic setting
                            int[] pulse_tmp;
                            bool[] Enable_state_table = new bool[] { test_parameter.ESwire_state, test_parameter.ASwire_state, test_parameter.ENVO4_state };
                            int[] Enable_num_table = new int[] { RTBBControl.ESwire, RTBBControl.ASwire, RTBBControl.ENVO4 };
                            pulse_tmp = test_parameter.ESwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(true, pulse_tmp[pulse_idx]);

                            pulse_tmp = test_parameter.ASwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(false, pulse_tmp[pulse_idx]);
                            for (int i = 0; i < Enable_state_table.Length; i++) RTBBControl.Swire_Control(Enable_num_table[i], Enable_state_table[i]);
                        }

                        // Get input condtion
                        double vin, iin, vout, iout;
                        vin = InsControl._power.GetVoltage();
                        iin = InsControl._power.GetCurrent() * 1000;
                        vout = InsControl._34970A.Get_100Vol(2);
                        iout = InsControl._eload.GetIout() * 1000;
                        _sheet.Cells[row, XLS_Table.A] = temp;
                        _sheet.Cells[row, XLS_Table.B] = "LINK";
                        _sheet.Cells[row, XLS_Table.C] = string.Format("{0:##.000}", vin);
                        _sheet.Cells[row, XLS_Table.D] = string.Format("{0:##.000}", iin);
                        _sheet.Cells[row, XLS_Table.E] = "ESwire=" + test_parameter.ESwireList[bin_idx] + ", ASwire=" + test_parameter.ASwireList[bin_idx]; ;
                        _sheet.Cells[row, XLS_Table.F] = string.Format("{0:##.000}", iout);
                        _sheet.Cells[row, XLS_Table.G] = string.Format("{0:##.000}", vout);

                        for (int i = 0; i < 3; i++)
                        {
                            if (test_parameter.LX_item[i])
                            {
                                switch (i)
                                {
                                    case 0:
                                        // Freq print data
                                        FreqTask();
                                        double freq_mean, freq_max, freq_min;
                                        InsControl._scope.DoCommand(":MEASURE:STATISTICS MEAN");
                                        freq_mean = InsControl._scope.Meas_Result() / 1000;

                                        InsControl._scope.DoCommand(":MEASURE:STATISTICS MAX");
                                        freq_max = InsControl._scope.Meas_Result() / 1000;

                                        InsControl._scope.DoCommand(":MEASURE:STATISTICS MIN");
                                        freq_min = InsControl._scope.Meas_Result() / 1000;
                                        _sheet.Cells[row, XLS_Table.H] = string.Format("{0:##.000}", freq_mean);
                                        _sheet.Cells[row, XLS_Table.I] = string.Format("{0:##.000}", freq_max);
                                        _sheet.Cells[row, XLS_Table.J] = string.Format("{0:##.000}", freq_min);
                                        //TODO: Save waveform
                                        InsControl._scope.Root_RUN();
                                        break;
                                    case 1:
                                        // Rising task
                                        RiseTask();
                                        double rise = InsControl._scope.doQueryNumber(":MEASure:SLEWrate? CHANnel1,RISing");
                                        double rise_time = InsControl._scope.Meas_CH1Rise();
                                        //TODO: Save waveform
                                        InsControl._scope.Root_RUN();

                                        // Falling task
                                        FallTask();
                                        double fall = InsControl._scope.doQueryNumber(":MEASure:SLEWrate? CHANnel1,Falling");
                                        double fall_time = InsControl._scope.Meas_CH1Fall();
                                        _sheet.Cells[row, XLS_Table.K] = string.Format("{0:##.000}", rise_time * Math.Pow(10, 9));
                                        _sheet.Cells[row, XLS_Table.L] = string.Format("{0:##.000}", rise * Math.Pow(10, -6));
                                        _sheet.Cells[row, XLS_Table.M] = string.Format("{0:##.000}", fall_time * Math.Pow(10, 9));
                                        _sheet.Cells[row, XLS_Table.N] = string.Format("{0:##.000}", fall * Math.Pow(10, -6));

                                        //TODO: Save waveform
                                        InsControl._scope.Root_RUN();
                                        break;
                                    case 2:
                                        JitterTask();
                                        double MeaPKPK = InsControl._scope.doQueryNumber(":MEASure:HISTogram:PP?") * Math.Pow(10, 9);
                                        //double MeaMean = InsControl._scope.doQueryNumber(":MEASure:HISTogram:PP?");
                                        double MeaStdDev = InsControl._scope.doQueryNumber(":MEASure:HISTogram:STDDev?") * Math.Pow(10, 9);
                                        InsControl._scope.DoCommand(":HISTogram:MODE OFF");
                                        InsControl._scope.DoCommand("DISPlay:CGRade OFF");

                                        _sheet.Cells[row, XLS_Table.O] = MeaPKPK;
                                        _sheet.Cells[row, XLS_Table.P] = MeaStdDev;
                                        //_sheet.Cells[row, XLS_Table.Q] = "=O" + row + "*" + freq_mean + "*100 * 10 ^-9";
                                        //TODO: Save waveform
                                        InsControl._scope.Root_RUN();
                                        InsControl._scope.DoCommand(":HISTogram:MODE OFF");
                                        InsControl._scope.DoCommand("DISPlay:CGRade OFF");
                                        break;
                                }
                            }
                        }

                        row++;
                    } // Eload loop
                } // interface loop
            } // power loop

        Stop:
            stopWatch.Stop();
        }

        private void ChannelResize()
        {
            if (test_parameter.buck || test_parameter.boost)
            {
                double error = 9.999 * Math.Pow(10, 20);
                for (int i = 0; i < 4; i++)
                {
                    double Vpp = InsControl._scope.Meas_CH1VPP();
                    InsControl._scope.CH1_Level(Vpp / 5);
                    MyLib.Delay1ms(20);

                    if (Vpp > error)
                    {
                        i = 0;
                        InsControl._scope.CH1_Level(5);
                        MyLib.Delay1ms(20);
                    }
                }
            }
            else if (test_parameter.inverting)
            {
                for (int i = 0; i <= 1; i++)
                {
                    double avg = 0;
                    double Vmax = Math.Abs(InsControl._scope.Meas_CH1MAX());
                    double Vmin = Math.Abs(InsControl._scope.Meas_CH1MIN());
                    if (Vmax > Vmin)
                        avg = Vmax;
                    else if (Vmax < Vmin)
                        avg = Vmin;

                    InsControl._scope.CH1_Level(avg / 3);
                }
            }

        }

        private void ChannelTrigger()
        {
            double Vtop, Vbase;
            if (test_parameter.buck || test_parameter.inverting)
            {
                InsControl._scope.SetTrigModeEdge(false);
                Vtop = InsControl._scope.Meas_CH1Top();
                Vbase = InsControl._scope.Meas_CH1Base();
                double trigger_level = 0.65 * Vtop + 0.35 * Vbase;
                InsControl._scope.TriggerLevel_CH1(trigger_level);
            }
            else if (test_parameter.boost)
            {
                InsControl._scope.SetTrigModeEdge(true);
                Vtop = InsControl._scope.Meas_CH1Top();
                Vbase = InsControl._scope.Meas_CH1Base();
                double trigger_level = 0.45 * Vtop + 0.65 * Vbase;
                InsControl._scope.TriggerLevel_CH1(trigger_level);
            }
        }

        private void FreqTask()
        {
            InsControl._scope.Measure_Clear();
            InsControl._scope.DoCommand(":MEASURE:FREQ CHANnel1");
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.TimeScaleUs(5);
            InsControl._scope.TimeBasePosition(0);

            double period = InsControl._scope.Meas_CH1Period();
            double time_scale = period * 2; // show 5 cycle
            InsControl._scope.TimeScale(time_scale);
            InsControl._scope.CH1_Level(5);
            ChannelResize();
            ChannelTrigger();
            InsControl._scope.Root_STOP();
        }

        private void RiseTask()
        {
            InsControl._scope.Measure_Clear();
            //InsControl._scope.DoCommand(":MEASURE:FREQ CHANnel1");
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.TimeScaleUs(5);
            InsControl._scope.TimeBasePosition(0);
            double period = InsControl._scope.Meas_CH1Period();
            double time_scale = period * 2; // show 5 cycle
            InsControl._scope.TimeScale(time_scale);
            InsControl._scope.CH1_Level(5);
            ChannelResize();
            ChannelTrigger();
            InsControl._scope.SetTrigModeEdge(false);
            InsControl._scope.DoCommand(":MEASure:SLEWrate CHANnel1, RISing");
            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.DoCommand(":MARKer:MODE ON");
            InsControl._scope.SlewRate20_80Range();
            InsControl._scope.DoCommand(":MEASURE:RISetime CHANnel1");
            MyLib.Delay1ms(200);
            double XDelta = InsControl._scope.Meas_CH1XDelta();
            double XDelta_standard = Math.Pow(10, -9);
            time_scale = XDelta * 2;
            InsControl._scope.TimeScale(time_scale);
            InsControl._scope.TimeBasePosition(time_scale);
            InsControl._scope.Root_STOP();
        }

        private void FallTask()
        {
            InsControl._scope.Measure_Clear();
            InsControl._scope.SetTrigModeEdge(true);
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.DoCommand(":MEASure:SLEWrate CHANnel1,Falling");
            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.DoCommand(":MEASURE:FALLtime CHANnel1");
            InsControl._scope.DoCommand(":MARKer:MODE ON");
            MyLib.Delay1ms(200);
            double XDelta = InsControl._scope.Meas_CH1XDelta();
            double time_scale = XDelta * 2;
            InsControl._scope.TimeScale(time_scale * 2);
            InsControl._scope.TimeBasePosition(time_scale);
            InsControl._scope.Root_STOP();
        }

        private void JitterTask()
        {
            InsControl._scope.Measure_Clear();
            InsControl._scope.DoCommand(":MEASURE:FREQ CHANnel1");
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.TimeScaleUs(5);
            InsControl._scope.TimeBasePosition(0);

            double period = InsControl._scope.Meas_CH1Period();
            double time_scale = period * 1.5; // show 1.5 cycle
            InsControl._scope.TimeScale(time_scale);
            InsControl._scope.TimeBasePosition(time_scale * 3);
            InsControl._scope.CH1_Level(5);
            ChannelResize();
            ChannelTrigger();

            double Rlimit = (time_scale * 6.4);
            double Llimit = (time_scale * 0.2);
            double Vtop = InsControl._scope.Meas_CH1Top();
            double Vbase = InsControl._scope.Meas_CH1Base();
            double histogramLevel = Vtop * 0.5 + Vbase * 0.5;
            InsControl._scope.DoCommand(":HISTogram:MODE OFF");
            InsControl._scope.DoCommand(":DISPlay:CGRade 1");
            InsControl._scope.DoCommand(":HISTogram:SCALe:SIZE 2");
            InsControl._scope.DoCommand(":HISTogram:MODE WAVeform");
            InsControl._scope.DoCommand(":HISTogram:WINDow:SOURce CHANnel1");
            InsControl._scope.DoCommand(":HISTogram:WINDow:LLIMit " + Llimit);
            InsControl._scope.DoCommand(":HISTogram:WINDow:RLIMit " + Rlimit);
            InsControl._scope.DoCommand(":HISTogram:WINDow:TLIMit " + (histogramLevel * 1.05));
            InsControl._scope.DoCommand(":HISTogram:WINDow:BLIMit " + (histogramLevel * 0.95));
            MyLib.Delay1ms(6000);
            InsControl._scope.Root_STOP();
        }

    }
}
