﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace SoftStartTiming
{

    public class ATE_SoftStartTiming : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        //Excel.Chart _chart;

        public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        //TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_SoftStartTiming()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void OSCInit()
        {
            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
            InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
            InsControl._scope.Root_RUN();
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            InsControl._scope.AutoTrigger();
            InsControl._scope.Trigger_CH1();
            MyLib.WaveformCheck();
            InsControl._scope.CHx_On(1);
            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                if (test_parameter.scope_en[i])
                {
                    InsControl._scope.CHx_On(i + 2);
                    InsControl._scope.CHx_Level(i + 2, test_parameter.VinList[0] * 3);
                    InsControl._scope.CHx_Offset(i + 2, 0);
                }
            }

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();
            InsControl._scope.CH3_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();

            InsControl._scope.TriggerLevel_CH1(1);

            //:MEASure:THResholds:GENeral:METHod\sALL,PERCent
            //:MEASure:THResholds:GENeral:PERCent\sALL,100,50,0

            InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:METHod ALL,PERCent");
            InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:PERCent ALL,100,50,1");

            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:METHod ALL,PERCent");
            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent ALL,100,50,1");
            InsControl._scope.Root_RUN();
            MyLib.Delay1ms(200 + (int)((test_parameter.ontime_scale_ms * 10) * 1.2));
            MyLib.WaveformCheck();

            InsControl._scope.DoCommand(":MARKer:MODE MANual");
            InsControl._scope.DoCommand(":MARKer3:ENABle OFF");
            InsControl._scope.DoCommand(":MARKer4:ENABle OFF");
            InsControl._scope.DoCommand(":MARKer3:TYPE XMANual");
            InsControl._scope.DoCommand(":MARKer4:TYPE XMANual");
            int marker_idx = 0;

            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                if (test_parameter.scope_en[i])
                {
                    string cmd;
                    cmd = string.Format(":MARKer{0}:ENABle ON", ++marker_idx);
                    InsControl._scope.DoCommand(cmd);
                    cmd = string.Format(":MARKer{0}:SOURce CHANnel1", marker_idx);
                    InsControl._scope.DoCommand(cmd);
                    cmd = string.Format(":MARKer{0}:TYPE XMANual", marker_idx);
                    InsControl._scope.DoCommand(cmd);

                    cmd = string.Format(":MARKer{0}:ENABle ON", ++marker_idx);
                    InsControl._scope.DoCommand(cmd);
                    cmd = string.Format(":MARKer{0}:TYPE XMANual", marker_idx);
                    InsControl._scope.DoCommand(cmd);
                    cmd = string.Format(":MARKer{0}:SOURce CHANnel1", marker_idx);
                    InsControl._scope.DoCommand(cmd);

                    cmd = string.Format(":MARKer{0}:DELTa MARKer{1}, ON", marker_idx, marker_idx - 1);
                    InsControl._scope.DoCommand(cmd);

                }
            }

            // measure current delta-time.
            InsControl._scope.DoCommand(":MEASure:STATistics CURRent");
        }


        private void Scope_Channel_Resize(int idx, string path)
        {
            InsControl._scope.AutoTrigger();
            InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
            MyLib.Delay1ms(800);
            MyLib.Delay1ms(800);

            double time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");

            InsControl._scope.TimeScaleUs(1);

            RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, path); // test conditions
            MyLib.Delay1ms(800);

            switch (test_parameter.trigger_event)
            {
                case 0: // gpio
                    InsControl._scope.TriggerLevel_CH1(1); // gui trigger level
                    InsControl._scope.CHx_Level(1, 3.3 / 2.5);
                    InsControl._scope.CHx_Offset(1, 3.3 / 2.5);
                    if (test_parameter.sleep_mode)
                        RTDev.GpEn_Enable();
                    else
                        RTDev.GpEn_Disable();
                    break;
                case 1: // i2c trigger
                    double vout = InsControl._scope.Meas_CH1MAX();
                    InsControl._scope.TriggerLevel_CH1(vout * 0.35);
                    break;
                case 2: // vin trigger
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
                    InsControl._scope.TriggerLevel_CH1(test_parameter.VinList[idx] * 0.35);
                    break;
            }
            MyLib.Delay1s(1);

            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                if (test_parameter.scope_en[i])
                {
                    InsControl._scope.CHx_Level(i + 2, test_parameter.VinList[0] * 3);
                    MyLib.Delay1ms(300);
                }
            }
            MyLib.Delay1s(1);

            for (int i = 0; i < 2; i++)
            {
                for (int ch_idx = 0; ch_idx < test_parameter.scope_en.Length; ch_idx++)
                {
                    if (test_parameter.scope_en[ch_idx])
                    {
                        double vmax = InsControl._scope.Measure_Ch_Max(ch_idx + 2);
                        InsControl._scope.CHx_Level(ch_idx + 2, vmax / 2.5);
                        InsControl._scope.CHx_Offset(ch_idx + 2, (vmax / 2.5) * (ch_idx + 1) * 0.5);
                        MyLib.Delay1ms(800);
                    }
                }
            }

            PowerOffEvent();
            InsControl._scope.TimeScale(time_scale);
            MyLib.Delay1ms(250);
        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            RTDev.GpioInit();

            int vin_cnt = test_parameter.VinList.Count;
            int row = 8;
            int wave_row = 8;
            int wave_pos = 0;
            string[] binList;
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if true
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;

            _sheet.Cells[1, XLS_Table.A] = "Item";
            _sheet.Cells[2, XLS_Table.A] = "Test Conditions";
            _sheet.Cells[3, XLS_Table.A] = "Result";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _range = _sheet.Range["A1", "A4"];
            _range.Font.Bold = true;
            _range.Interior.Color = Color.FromArgb(255, 178, 102);
            _range = _sheet.Range["A2"];
            _range.RowHeight = 100;
            _range = _sheet.Range["B1"];
            _range.ColumnWidth = 60;
            _range = _sheet.Range["A1", "B4"];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // print test conditions


            _sheet.Cells[row, XLS_Table.D] = "No.";
            _sheet.Cells[row, XLS_Table.E] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.F] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.G] = "Bin file";
            _range = _sheet.Range["D" + row, "G" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _sheet.Cells[row, XLS_Table.H] = "DT1 (ms)";
            _sheet.Cells[row, XLS_Table.I] = "SST1 (us)";
            _sheet.Cells[row, XLS_Table.J] = "DT2 (ms)";
            _sheet.Cells[row, XLS_Table.K] = "SST2 (us)";
            _sheet.Cells[row, XLS_Table.L] = "DT3 (ms)";
            _sheet.Cells[row, XLS_Table.M] = "SST3 (us)";

            _range = _sheet.Range["H" + row, "M" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif
            InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1s(1);
            int cnt = 0;
            for (int select_idx = 0; select_idx < test_parameter.bin_en.Length; select_idx++)
            {
                if (test_parameter.bin_en[select_idx])
                {
                    binList = MyLib.ListBinFile(test_parameter.bin_path[select_idx]);
                    bin_cnt = binList.Length;
                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                    {
                        // repeat i2c setting time scale need to reset to deafult
                        InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                        InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);

                        for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                        {
                            int retry_cnt = 0;

                            if (test_parameter.run_stop == true) goto Stop;
                            if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                            /* test initial setting */
                            //InsControl._scope.DoCommand(":MARKer:MODE OFF");
                            string file_name;
                            string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);

                            test_parameter.sleep_mode = (res.IndexOf("sleep_en") == -1) ? false : true;

                            InsControl._scope.Measure_Clear();
                            MyLib.Delay1s(1);

                            // Call Measure display to waveform
                            for (int i = 0; i < test_parameter.scope_en.Length; i++)
                            {
                                if (test_parameter.scope_en[i])
                                    InsControl._scope.DoCommand(":MEASure:RISetime CHANnel" + (i + 2).ToString());
                            }
                            // Call Measure display to waveform
                            for (int i = 0; i < test_parameter.scope_en.Length; i++)
                            {
                                if (test_parameter.scope_en[i])
                                {
                                    // measure Delta function configure
                                    // bool isRising1, int start, int level1, bool isRising2, int stop, int level2
                                    if (!test_parameter.sleep_mode)
                                    {
                                        // sleep mode disable: measure point is PWRDIS falling to Rails rising.
                                        InsControl._scope.SetDeltaTime(false, 1, 0, true, 1, 0);
                                        InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1, CHANnel" + (i + 2).ToString());
                                    }
                                    else
                                    {
                                        // sleep mode enable : measure point is Sleep_GPIO rising to Rails rising
                                        InsControl._scope.SetDeltaTime(true, 1, 0, true, 1, 0);
                                        InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1, CHANnel" + (i + 2).ToString());
                                    }
                                }
                                MyLib.Delay1ms(500);
                            }

                            Console.WriteLine(res);
                            file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V",
                                                        row - 22, res, temp,
                                                        test_parameter.VinList[vin_idx]
                                                        );
                        
                            double time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
                            // include test condition
                            Scope_Channel_Resize(vin_idx, binList[bin_idx]);
                            double tempVin = ori_vinTable[vin_idx];
                            MyLib.WaveformCheck();
                        retest:;

                            if (retry_cnt > 3)
                            {
                                _sheet.Cells[row, XLS_Table.F] = "sATE test fail_" + res;
                                InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                                retry_cnt = 0;
                                row++;
                                continue;
                            }

                            InsControl._scope.NormalTrigger();
                            MyLib.Delay1ms(800);
                            switch (test_parameter.trigger_event)
                            {
                                case 0:
                                    // GPIO trigger event
                                    InsControl._scope.Root_Clear();
                                    if (test_parameter.sleep_mode)
                                    {
                                        InsControl._scope.SetTrigModeEdge(false);
                                        MyLib.Delay1ms(800);
                                        RTDev.GpEn_Enable();
                                    }
                                    else
                                    {
                                        InsControl._scope.SetTrigModeEdge(true);
                                        MyLib.Delay1ms(1000);
                                        RTDev.GpEn_Disable();
                                    }
                                    time_scale = time_scale * 1000;
                                    MyLib.Delay1ms((int)((time_scale * 10) * 1.2) + 500);
                                    break;
                                case 1:
                                    // I2C trigger event
                                    break;
                                case 2:
                                    // Power supply trigger event
                                    InsControl._power.AutoPowerOff();
                                    MyLib.Delay1ms((int)((time_scale * 10) * 1.2) + 500);
                                    break;
                            }
                            InsControl._scope.Root_STOP();
                            MyLib.Delay1s(1);

                            time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
                            string reslult_lits = InsControl._scope.doQeury(":MEASure:RESults?");
                            List<double> delay_time = reslult_lits.Split(',').Select(double.Parse).ToList();
                            double time_scale_threshold = (time_scale * 5);

                            int marker_idx = 0;
                            for (int i = 0; i < test_parameter.scope_en.Length; i++)
                            {
                                if (test_parameter.scope_en[i])
                                {
                                    string cmd;
                                    cmd = string.Format(":MARKer{0}:X:POSition {1}", ++marker_idx, 0);
                                    InsControl._scope.DoCommand(cmd);
                                    cmd = string.Format(":MARKer{0}:X:POSition {1}", ++marker_idx, delay_time[i]);
                                    InsControl._scope.DoCommand(cmd);
                                }
                            }

                            double delay_time_res = 0;
                            double sst_res = 0;
                            switch (select_idx)
                            {
                                case 0:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel2");
                                    sst_res = InsControl._scope.Meas_CH2Rise();
                                    break;
                                case 1:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel3");
                                    sst_res = InsControl._scope.Meas_CH3Rise();
                                    break;
                                case 2:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel4");
                                    sst_res = InsControl._scope.Meas_CH4Rise();
                                    break;
                            }

                            if (delay_time_res >= time_scale * 4)
                            {
                                if (delay_time_res > Math.Pow(10, 20))
                                {
                                    retry_cnt++;
                                    InsControl._scope.Root_RUN();
                                    PowerOffEvent();
                                    InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                                    goto retest;
                                }
                                if (delay_time_res > 0)
                                {
                                    double temp = (delay_time_res * 1.2) / 4;
                                    InsControl._scope.TimeScale(temp);
                                    InsControl._scope.TimeBasePosition(temp * 3);
                                }
                                else
                                {
                                    InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                                    InsControl._scope.TimeBasePosition(test_parameter.ontime_scale_ms * 3);
                                }
                                InsControl._scope.Root_RUN();
                                PowerOffEvent();
                                goto retest;
                            }
                            else if (delay_time_res < time_scale)
                            {
                                if(delay_time_res < sst_res)
                                {
                                    InsControl._scope.TimeScale(sst_res);
                                    InsControl._scope.TimeBasePosition(sst_res * 3);
                                }
                                else
                                {
                                    InsControl._scope.TimeScale(delay_time_res / 2);
                                    InsControl._scope.TimeBasePosition((delay_time_res / 2) * 3);
                                }
                                //InsControl._scope.Root_RUN();
                                //PowerOffEvent();
                                //goto retest;
                            }
                            InsControl._scope.SaveWaveform(test_parameter.waveform_path, res);


#if true
                            double vin, dt1, dt2, dt3, sst1, sst2, sst3;
                            vin = InsControl._power.GetVoltage();


                            _sheet.Cells[row, XLS_Table.D] = cnt++;
                            _sheet.Cells[row, XLS_Table.E] = temp;
                            _sheet.Cells[row, XLS_Table.F] = vin;
                            _sheet.Cells[row, XLS_Table.G] = res;

                            //":MEASure:DELTatime CHANnel1,CHANnel2
                            if (test_parameter.scope_en[0])
                            {
                                dt1 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel2");
                                _sheet.Cells[row, XLS_Table.H] = dt1 * Math.Pow(10, 3);

                                sst1 = InsControl._scope.Meas_CH2Rise();
                                _sheet.Cells[row, XLS_Table.I] = sst1 * Math.Pow(10, 6);
                            }

                            if (test_parameter.scope_en[1])
                            {
                                dt2 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel3");
                                _sheet.Cells[row, XLS_Table.J] = dt2 * Math.Pow(10, 3);

                                sst2 = InsControl._scope.Meas_CH3Rise();
                                _sheet.Cells[row, XLS_Table.K] = sst2 * Math.Pow(10, 6);
                            }

                            if (test_parameter.scope_en[2])
                            {
                                dt3 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel4");
                                _sheet.Cells[row, XLS_Table.L] = dt3 * Math.Pow(10, 3);

                                sst3 = InsControl._scope.Meas_CH3Rise();
                                _sheet.Cells[row, XLS_Table.M] = sst3 * Math.Pow(10, 6);
                            }

                            switch(wave_pos)
                            {
                                case 0:
                                    _sheet.Cells[wave_row, XLS_Table.S] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.T] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.U] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.V] = "Bin file";
                                    _range = _sheet.Range["S" + wave_row, "V" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);

                                    _sheet.Cells[wave_row + 1, XLS_Table.S] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.T] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.U] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.V] = "=G" + row;
                                    _range = _sheet.Range["S" + (wave_row + 2).ToString(), "AA" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 1:
                                    _sheet.Cells[wave_row, XLS_Table.AD] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.AE] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.AF] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.AG] = "Bin file";
                                    _range = _sheet.Range["AD" + wave_row, "AG" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);

                                    _sheet.Cells[wave_row + 1, XLS_Table.AD] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AE] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AF] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AG] = "=G" + row;
                                    _range = _sheet.Range["AD" + (wave_row + 2).ToString(), "AL" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 2:
                                    _sheet.Cells[wave_row, XLS_Table.AO] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.AP] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.AQ] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.AR] = "Bin file";
                                    _range = _sheet.Range["AO" + wave_row, "AR" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);

                                    _sheet.Cells[wave_row + 1, XLS_Table.AO] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AP] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AQ] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AR] = "=G" + row;
                                    _range = _sheet.Range["AO" + (wave_row + 2).ToString(), "AW" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 3:
                                    _sheet.Cells[wave_row, XLS_Table.AZ] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.BA] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.BB] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.BC] = "Bin file";
                                    _range = _sheet.Range["AZ" + wave_row, "BC" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);

                                    _sheet.Cells[wave_row + 1, XLS_Table.AZ] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BA] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BB] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BC] = "=G" + row;
                                    _range = _sheet.Range["AZ" + (wave_row + 2).ToString(), "BH" + (wave_row + 16).ToString()];
                                    wave_pos = 0;  wave_row = wave_row + 19;
                                    break;
                            }
                            MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, res);
                            row++;
#endif
                            InsControl._scope.Root_RUN();
                            PowerOffEvent();
                        }
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
#if Report
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;

            //Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_DT_SST_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif

        }


        private void PowerOffEvent()
        {
            switch (test_parameter.trigger_event)
            {
                case 0: // gpio power disable
                    if (test_parameter.sleep_mode)
                        RTDev.GpEn_Disable();
                    else
                        RTDev.GpEn_Enable();
                    break;
                case 1:
                    break;
                case 2: // vin trigger
                    InsControl._power.AutoPowerOff();
                    break;
            }
        }


    }
}




