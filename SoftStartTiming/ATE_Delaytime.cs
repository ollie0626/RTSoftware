

#define Report_en
#define Power_en
#define Eload_en

using System;
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

    public class ATE_DelayTime : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        //Excel.Chart _chart;

        //public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        //TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        const int meas_dt1 = 1;
        const int meas_dt2 = 2;
        const int meas_dt3 = 3;

        const int meas_sst1 = 4;
        const int meas_sst2 = 5;
        const int meas_sst3 = 6;

        const int meas_vmax = 7;

        public ATE_DelayTime()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void GpioOnSelect(int num)
        {
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Enable();
                    break;
                case 1:
                    RTDev.Gp2En_Enable();
                    break;
                case 2:
                    RTDev.Gp3En_Enable();
                    break;
            }
        }

        private void GpioOffSelect(int num)
        {
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Disable();
                    break;
                case 1:
                    RTDev.Gp2En_Disable();
                    break;
                case 2:
                    RTDev.Gp3En_Disable();
                    break;
            }
        }

        private void OSCInit()
        {
            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                InsControl._tek_scope.SetTimeBasePosition(15);
                InsControl._tek_scope.SetRun();
                InsControl._tek_scope.SetTriggerMode(); // auto trigger
                InsControl._tek_scope.SetTriggerSource(1);

                InsControl._tek_scope.CHx_On(1);
                InsControl._tek_scope.CHx_On(2);
                InsControl._tek_scope.CHx_On(3);
                InsControl._tek_scope.CHx_On(4);

                InsControl._tek_scope.CHx_BWlimitOn(1);
                InsControl._tek_scope.CHx_BWlimitOn(2);
                InsControl._tek_scope.CHx_BWlimitOn(3);
                InsControl._tek_scope.CHx_BWlimitOn(4);

                InsControl._tek_scope.CHx_Position(1, 1.5);
                InsControl._tek_scope.CHx_Position(2, 0);
                InsControl._tek_scope.CHx_Position(3, -1);
                InsControl._tek_scope.CHx_Position(4, -3);

                for (int i = 0; i < test_parameter.scope_en.Length; i++)
                {
                    if (test_parameter.scope_en[i])
                    {
                        InsControl._tek_scope.CHx_Level(i + 2, test_parameter.VinList[0]);
                    }
                }

                for (int i = 0; i < test_parameter.scope_en.Length; i++)
                {
                    if (test_parameter.scope_en[i])
                    {
                        // set all meassure
                        switch (i)
                        {
                            case 0: // Channel2
                                if (test_parameter.sleep_mode)
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt1, 1, 2, true, true);
                                else
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt1, 1, 2, false, true);
                                InsControl._tek_scope.SetMeasureSource(2, meas_sst1, "RISe");
                                //InsControl._tek_scope.SetMeasureSource(2, meas_vtop1, "HIGH");
                                //InsControl._tek_scope.SetMeasureSource(2, meas_vbase1, "LOW");

                                break;
                            case 1: // Channel3
                                if (test_parameter.sleep_mode)
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt2, 1, 3, true, true);
                                else
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt2, 1, 3, false, true);
                                InsControl._tek_scope.SetMeasureSource(3, meas_sst2, "RISe");
                                //InsControl._tek_scope.SetMeasureSource(3, meas_vtop2, "HIGH");
                                //InsControl._tek_scope.SetMeasureSource(3, meas_vbase2, "LOW");

                                break;
                            case 2: // Channel4
                                if (test_parameter.sleep_mode)
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt3, 1, 4, true, true);
                                else
                                    InsControl._tek_scope.SetMeasureDelay(meas_dt3, 1, 4, false, true);
                                InsControl._tek_scope.SetMeasureSource(4, meas_sst3, "RISe");
                                //InsControl._tek_scope.SetMeasureSource(4, meas_vtop3, "HIGH");
                                //InsControl._tek_scope.SetMeasureSource(4, meas_vbase3, "LOW");
                                break;
                        }
                    }
                }

                InsControl._tek_scope.DoCommand("MEASUrement:IMMed:REFLevel:METHod PERCent");
                InsControl._tek_scope.DoCommand("MEASUrement:REFLevel:PERCent:HIGH 100");
                InsControl._tek_scope.DoCommand("MEASUrement:REFLevel:PERCent:MID 50");
                InsControl._tek_scope.DoCommand("MEASUrement:REFLevel:PERCent:LOW 1");
                InsControl._tek_scope.DoCommand("MEASUrement:REFLevel:PERCent:MID2 10");
                InsControl._tek_scope.DoCommand("HORizontal:ROLL OFF");
                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            }
            else
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
                InsControl._scope.CHx_Offset(1, 0);
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
        }

        private void TriggerEvent(int idx)
        {
            switch (test_parameter.trigger_event)
            {
                case 0: // gpio trigger

                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTriggerSource(1);
                        InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                        InsControl._tek_scope.CHx_Position(1, 1.5);
                    }
                    else
                    {
                        InsControl._scope.TriggerLevel_CH1(1); // gui trigger level
                        InsControl._scope.CHx_Level(1, 3.3 / 2);
                        InsControl._scope.CHx_Offset(1, 0);
                    }

                    if (test_parameter.sleep_mode)
                        GpioOnSelect(test_parameter.gpio_pin);
                    else
                        GpioOffSelect(test_parameter.gpio_pin);
                    break;
                case 1: // i2c trigger
                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTriggerSource(1);
                        InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                        InsControl._tek_scope.CHx_Position(1, 1.5);
                    }
                    else
                    {
                        InsControl._scope.TriggerLevel_CH1(1); // gui trigger level
                        InsControl._scope.CHx_Level(1, 3.3 / 2);
                        InsControl._scope.CHx_Offset(1, 0);
                    }

                    // rails enable
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });


                    break;
                case 2: // vin trigger
#if Power_en
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
#endif
                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTriggerSource(1);
                        InsControl._tek_scope.SetTriggerLevel(test_parameter.VinList[idx] * 0.35);
                    }
                    else
                    {
                        InsControl._scope.TriggerLevel_CH1(test_parameter.VinList[idx] * 0.35);
                    }
                    break;
            }
        }

        private void LevelEvent()
        {
            InsControl._tek_scope.SetMeasureSource(2, meas_vmax, "MAXimum");
            InsControl._tek_scope.CHx_Level(2, test_parameter.ch2_level);
            InsControl._tek_scope.CHx_Level(3, test_parameter.ch3_level);
            InsControl._tek_scope.CHx_Level(4, test_parameter.ch4_level);
            int re_cnt = 0;
            for (int ch_idx = 0; ch_idx < test_parameter.scope_en.Length; ch_idx++)
            {
                if (test_parameter.scope_en[ch_idx])
                {
                re_scale:;
                    if (re_cnt > 3)
                    {
                        re_cnt = 0;
                        continue;
                    }

                    double vmax = 0;

                    if (InsControl._tek_scope_en)
                    {
                        // tek get max
                        for (int k = 0; k < 3; k++)
                        {
                            vmax = InsControl._tek_scope.CHx_Meas_Mean(ch_idx + 2, meas_vmax);
                            vmax = InsControl._tek_scope.CHx_Meas_Mean(ch_idx + 2, meas_vmax);
                            MyLib.Delay1ms(50);
                            vmax = InsControl._tek_scope.CHx_Meas_Mean(ch_idx + 2, meas_vmax);
                            //Console.WriteLine("VMax = {0}", vmax);

                            if (vmax > 0.3 && vmax < Math.Pow(10, 3))
                                InsControl._tek_scope.CHx_Level(ch_idx + 2, vmax / 3);
                            MyLib.Delay1ms(300);
                        }
                    }
                    else
                    {
                        // agilent get max
                        vmax = InsControl._scope.Measure_Ch_Max(ch_idx + 2);
                        if (vmax > Math.Pow(10, 9))
                        {
                            re_cnt++;
                            InsControl._scope.CHx_Level(ch_idx + 2, test_parameter.VinList[0] * 2.5);
                            InsControl._scope.CHx_Offset(ch_idx + 2, test_parameter.VinList[0] * 2.5 * (ch_idx + 1));
                            MyLib.Delay1ms(500);
                            goto re_scale;
                        }
                        InsControl._scope.CHx_Level(ch_idx + 2, vmax / 2.5);
                        InsControl._scope.CHx_Offset(ch_idx + 2, (vmax / 2.5) * (ch_idx + 1));
                    }
                    MyLib.Delay1ms(300);
                }
            }
        }

        private void Scope_Channel_Resize(int idx, string path)
        {
            double time_scale = 0;
            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetRun();
                InsControl._tek_scope.SetTriggerMode();
                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
                if (time_scale <= 55 * Math.Pow(10, -6) || time_scale > 100 * Math.Pow(10, -3))
                {
                    time_scale = test_parameter.ontime_scale_ms / 1000;
                }
            }
            else
            {
                time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
            }

            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTimeScale((25 * Math.Pow(10, -12)));
                //InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 2E6");
                //InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                //MyLib.Delay1ms(200);
            }
            else
            {
                InsControl._scope.TimeScaleUs(1);
            }

            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetRun();
                InsControl._tek_scope.SetTriggerMode();
            }
            else
            {
                InsControl._scope.Root_RUN();
                InsControl._scope.AutoTrigger();
            }
#if Power_en
            InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
#endif
            MyLib.Delay1ms(1000);

            for (int i = 0; i < 100; i++)
            {
                int ret = RTDev.I2C_WriteBin((byte)(test_parameter.slave), 0x00, path); // test conditions
                Console.WriteLine("I2C Return value {0}", ret);
                if (ret == 0) break;
            }
            TriggerEvent(idx);
            //MyLib.Delay1ms(3000);
            for (int i = 0; i < 100; i++)
            {
                int ret = RTDev.I2C_WriteBin((byte)(test_parameter.slave), 0x00, path); // test conditions
                Console.WriteLine("I2C Return value {0}", ret);
                if (ret == 0) break;
            }


            //MyLib.Delay1s(1);

            if (InsControl._tek_scope_en) MyLib.Delay1s(1);

            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                if (test_parameter.scope_en[i])
                {
                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.CHx_Level(i + 2, test_parameter.VinList[0] / 2);
                        InsControl._tek_scope.CHx_Position(i + 2, (i + 1) * -1);
                        //MyLib.Delay1ms(800);
                    }
                    else
                    {
                        InsControl._scope.CHx_Level(i + 2, test_parameter.VinList[0] * 3);
                        InsControl._scope.CHx_Offset(i + 2, test_parameter.VinList[0] * 3 * (i + 1));
                        //MyLib.Delay1ms(800);
                    }
                }
            }
            MyLib.Delay1ms(900);

            LevelEvent();
            PowerOffEvent();

            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTimeScale(time_scale);
                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
            else
                InsControl._scope.TimeScale(time_scale);


            //MyLib.Delay1ms(250);
        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
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

#if Report_en
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
#endif
            //InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1s(1);
            int cnt = 0;
            for (int select_idx = 0; select_idx < test_parameter.bin_en.Length; select_idx++)
            {
                if (test_parameter.bin_en[select_idx])
                {

#region "Report initial"
#if Report_en
                    _sheet = _book.Worksheets.Add();
                    _sheet.Name = "CH" + (select_idx + 1).ToString();
                    _sheet.Cells.Font.Name = "Calibri";
                    _sheet.Cells.Font.Size = 11;
                    row = 8;
                    wave_row = 8;
                    wave_pos = 0;
                    _sheet.Cells[1, XLS_Table.A] = "Item";
                    _sheet.Cells[2, XLS_Table.A] = "Test Conditions";
                    _sheet.Cells[3, XLS_Table.A] = "Result";
                    _sheet.Cells[4, XLS_Table.A] = "Note";
                    _range = _sheet.Range["A1", "A4"];
                    _range.Font.Bold = true;
                    _range.Interior.Color = Color.FromArgb(255, 178, 102);
                    _range = _sheet.Range["A2"];
                    _range.RowHeight = 150;
                    _range = _sheet.Range["B1"];
                    _range.ColumnWidth = 60;
                    _range = _sheet.Range["A1", "B4"];
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    // print test conditions
                    _sheet.Cells[1, XLS_Table.B] = "Delay time/Slot time";
                    _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions + test_parameter.bin_file_cnt;

                    _sheet.Cells[row, XLS_Table.D] = "No.";
                    _sheet.Cells[row, XLS_Table.E] = "Temp(C)";
                    _sheet.Cells[row, XLS_Table.F] = "Vin(V)";
                    _sheet.Cells[row, XLS_Table.G] = "Bin file";
                    _range = _sheet.Range["D" + row, "G" + row];
                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    // major measure timing
                    _sheet.Cells[row, XLS_Table.H] = test_parameter.delay_us_en ? "DT1 (us)" : "DT1 (ms)";
                    _sheet.Cells[row, XLS_Table.I] = "SST1 (us)";
                    _sheet.Cells[row, XLS_Table.J] = test_parameter.delay_us_en ? "DT2 (us)" : "DT2 (ms)";
                    _sheet.Cells[row, XLS_Table.K] = "SST2 (us)";
                    _sheet.Cells[row, XLS_Table.L] = test_parameter.delay_us_en ? "DT3 (us)" : "DT3 (ms)";
                    _sheet.Cells[row, XLS_Table.M] = "SST3 (us)";

                    // Add new measure
                    _sheet.Cells[row, XLS_Table.N] = "V1 Top (V)";
                    _sheet.Cells[row, XLS_Table.O] = "V2 Top (V)";
                    _sheet.Cells[row, XLS_Table.P] = "V3 Top (V)";
                    _sheet.Cells[row, XLS_Table.Q] = "V1 Base (V)";
                    _sheet.Cells[row, XLS_Table.R] = "V2 Base (V)";
                    _sheet.Cells[row, XLS_Table.S] = "V3 Base (V)";
                    _sheet.Cells[row, XLS_Table.T] = "Max (V)";
                    _sheet.Cells[row, XLS_Table.U] = "Min (V)";
                    _sheet.Cells[row, XLS_Table.V] = "Pass/Fail";

                    _range = _sheet.Range["H" + row, "U" + row];
                    _range.Interior.Color = Color.FromArgb(30, 144, 255);
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                    _range = _sheet.Range["V" + row, "V" + row];
                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    row++;
#endif
#endregion

                    stopWatch.Start();
                    binList = MyLib.ListBinFile(test_parameter.bin_path[select_idx]);
                    bin_cnt = binList.Length;
                    cnt = 0;

                    if (!Directory.Exists(test_parameter.waveform_path + @"/CH" + (select_idx).ToString()))
                    {
                        Directory.CreateDirectory(test_parameter.waveform_path + @"/CH" + (select_idx).ToString());
                    }

                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                    {
                        if (InsControl._tek_scope_en)
                        {
                            InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                            InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                            InsControl._tek_scope.SetTimeBasePosition(15);
                        }
                        else
                        {
                            // repeat i2c setting time scale need to reset to deafult
                            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                            InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
                        }


                        for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                        {
                            int retry_cnt = 0;

                            if (test_parameter.run_stop == true) goto Stop;
                            if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                            /* test initial setting */
                            //InsControl._scope.DoCommand(":MARKer:MODE OFF");
                            string file_name;
                            string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);

                            // gpio high power disable --> falling trigger, test_parameter.sleep_mode = false
                            // gpio high sleep mode --> rising trigger, test_parameter.sleep_mode = true

                            //test_parameter.sleep_mode = (res.IndexOf("sleep_en") == -1) ? false : true;
                            if (!InsControl._tek_scope_en) InsControl._scope.Measure_Clear();
                            MyLib.Delay1ms(500);

                            if (!InsControl._tek_scope_en)
                            {
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
                            }

                            Console.WriteLine(res);
                            file_name = string.Format("{0}_Temp={2}C_vin={3:0.##}V_{1}",
                                                        cnt, res, temp,
                                                        test_parameter.VinList[vin_idx]
                                                        );

                            double time_scale = 0;

                            if (InsControl._tek_scope_en)
                            {
                                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
                            }
                            else
                            {
                                time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
                            }

                        // include test condition
                        retest:;
   
                            Scope_Channel_Resize(vin_idx, binList[bin_idx]);
                            double tempVin = ori_vinTable[vin_idx];
                            if (!InsControl._tek_scope_en) MyLib.WaveformCheck();
                            //retest:;

                            if (retry_cnt > 3)
                            {
                                _sheet.Cells[row, XLS_Table.F] = "sATE test fail_" + res;
                                if (InsControl._tek_scope_en)
                                {
                                    InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                                    InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                                    InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                                }
                                else
                                    InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                                retry_cnt = 0;
                                row++;
                                continue;
                            }

                            if (InsControl._tek_scope_en)
                                InsControl._tek_scope.SetTriggerMode(false);
                            else
                                InsControl._scope.NormalTrigger();
                            MyLib.Delay1ms(500);

                            // power on trigger
                            switch (test_parameter.trigger_event)
                            {
                                case 0:
                                    // GPIO trigger event
                                    if (InsControl._tek_scope_en)
                                        InsControl._tek_scope.SetClear();
                                    else
                                        InsControl._scope.Root_Clear();

                                    //MyLib.Delay1ms(1500);
                                    if (test_parameter.sleep_mode)
                                    {
                                        // sleep mode
                                        if (InsControl._tek_scope_en)
                                            InsControl._tek_scope.SetTriggerRise();
                                        else
                                            InsControl._scope.SetTrigModeEdge(false);

                                        MyLib.Delay1ms(800);
                                        GpioOnSelect(test_parameter.gpio_pin);
                                    }
                                    else
                                    {
                                        // PWRDis
                                        if (InsControl._tek_scope_en)
                                            InsControl._tek_scope.SetTriggerFall();
                                        else
                                            InsControl._scope.SetTrigModeEdge(true);

                                        MyLib.Delay1ms(1000);
                                        GpioOffSelect(test_parameter.gpio_pin);
                                    }


                                    //time_scale = time_scale * 1000;
                                    //MyLib.Delay1ms((int)((time_scale * 10) * 1.2) + 500);

                                    if (InsControl._tek_scope_en) MyLib.Delay1s(1);
                                    break;
                                case 1:
                                    // I2C trigger event 
                                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                                    MyLib.Delay1s(1);
                                    break;
                                case 2:
                                    // Power supply trigger event
                                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                                    MyLib.Delay1ms((int)((time_scale * 10) * 1.2) + 500);
                                    break;
                            }

                            if (InsControl._tek_scope_en)
                                InsControl._tek_scope.SetStop();
                            else
                                InsControl._scope.Root_STOP();


                            if (InsControl._tek_scope_en)
                                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
                            else
                                time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");

                            string reslult_lits = "";
                            List<double> delay_time = new List<double>();
                            double time_scale_threshold = 0;

                            if (!InsControl._tek_scope_en)
                            {
                                // agilent add marker
                                reslult_lits = InsControl._scope.doQeury(":MEASure:RESults?");
                                delay_time = reslult_lits.Split(',').Select(double.Parse).ToList();
                                time_scale_threshold = (time_scale * 5);
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
                            }

                            double delay_time_res = 0;
                            double sst_res = 0;
                            int meas_idx = 1;
                            if (InsControl._tek_scope_en)
                            {
                                InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                                InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH1");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH" + (select_idx + 2).ToString());
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 0");
                                double data = 0;
                                
                                switch(select_idx)
                                {
                                    case 0:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt1);
                                        break;
                                    case 1:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt2);
                                        break;
                                    case 2:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt3);
                                        break;
                                }
                                
                                
                                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + data.ToString());
                                MyLib.Delay1ms(100);
                            }
                            MyLib.Delay1s(1);
                            // measure delay time
                            switch (select_idx)
                            {
                                case 0:
                                    if (InsControl._tek_scope_en)
                                    {
                                        delay_time_res = InsControl._tek_scope.MeasureMean(meas_dt1);
                                        sst_res = InsControl._tek_scope.MeasureMean(meas_sst1);
                                    }
                                    else
                                    {
                                        delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel2");
                                        sst_res = InsControl._scope.Meas_CH2Rise();
                                    }
                                    break;
                                case 1:
                                    if (InsControl._tek_scope_en)
                                    {
                                        delay_time_res = InsControl._tek_scope.MeasureMean(meas_dt2);
                                        sst_res = InsControl._tek_scope.MeasureMean(meas_sst2);
                                    }
                                    else
                                    {
                                        delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel3");
                                        sst_res = InsControl._scope.Meas_CH3Rise();
                                    }
                                    break;
                                case 2:
                                    if (InsControl._tek_scope_en)
                                    {
                                        delay_time_res = InsControl._tek_scope.MeasureMean(meas_dt3);
                                        sst_res = InsControl._tek_scope.MeasureMean(meas_sst3);
                                    }
                                    else
                                    {
                                        delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel4");
                                        sst_res = InsControl._scope.Meas_CH4Rise();
                                    }
                                    break;
                            }

                            sst_res = InsControl._tek_scope.MeasureMean(meas_sst3);
                            sst_res = InsControl._tek_scope.MeasureMean(meas_sst3);

                            double us_unit = Math.Pow(10, -6);
                            double ms_unit = Math.Pow(10, -3);
                            double[] time_table = new double[] { 
                                500 * us_unit, 400 * us_unit, 200 * us_unit, 100 * us_unit, 50 * us_unit, 20 * us_unit, 10 * us_unit, 
                                40 * ms_unit, 20 * ms_unit, 10 * ms_unit, 4 * ms_unit, 2 * ms_unit, 1 * ms_unit
                            };
                            List<double> min_list = new List<double>();
                            double time_temp = (delay_time_res) / 4.5;
                            double time_div = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");

                            if (delay_time_res > Math.Pow(10, 20) || delay_time_res < 0)
                            {
                                InsControl._tek_scope.SetRun();
                                InsControl._tek_scope.SetTriggerMode();
                                PowerOffEvent();
                                InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                                InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                                InsControl._tek_scope.SetTimeBasePosition(15);
                                retry_cnt++;
                                goto retest;
                            } 
                            else if (delay_time_res > time_div * 4)
                            {
                                InsControl._tek_scope.SetRun();
                                InsControl._tek_scope.SetTriggerMode();
                                PowerOffEvent();
                                InsControl._tek_scope.SetTimeScale(time_temp);
                                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                                InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                                InsControl._tek_scope.SetTimeBasePosition(15);

                                if (!(time_div == InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?")))
                                {
                                    retry_cnt++;
                                    goto retest;
                                }
                            }
                            else if (delay_time_res < time_div)
                            {
                                InsControl._tek_scope.SetRun();
                                InsControl._tek_scope.SetTriggerMode();
                                PowerOffEvent();
                                InsControl._tek_scope.SetTimeScale(delay_time_res/2);
                                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                                InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                                InsControl._tek_scope.SetTimeBasePosition(15);
                                retry_cnt++;
                                goto retest;
                            }

                            MyLib.Delay1ms(500);

                            if (InsControl._tek_scope_en)
                            {
                                InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                                InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH1");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH" + (select_idx + 2).ToString());
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                                MyLib.Delay1ms(100);
                                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 0");
                                double data = 0;
                                switch (select_idx)
                                {
                                    case 0:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt1);
                                        break;
                                    case 1:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt2);
                                        break;
                                    case 2:
                                        data = InsControl._tek_scope.MeasureMean(meas_dt3);
                                        break;
                                }
                                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + data.ToString());
                                MyLib.Delay1ms(100);
                            }

                            if (InsControl._tek_scope_en)
                            {
                                InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path + @"\CH" + (select_idx).ToString(), file_name);
                            }
                            else
                            {
                                InsControl._scope.SaveWaveform(test_parameter.waveform_path + @"\CH" + (select_idx).ToString(), file_name);
                            }

#if true
                            double vin = 0, dt1 = 0, dt2 = 0, dt3 = 0, sst1 = 0, sst2 = 0, sst3 = 0;
                            double vmax = 0, vmin = 0;
                            double vtop = 0, vbase = 0;
#if Power_en
                            vin = InsControl._power.GetVoltage();
#endif
#if Report_en
                            _sheet.Cells[row, XLS_Table.D] = cnt++;
                            _sheet.Cells[row, XLS_Table.E] = temp;
                            _sheet.Cells[row, XLS_Table.F] = vin;
                            _sheet.Cells[row, XLS_Table.G] = res;
#endif

                            // Add new measure
                            switch (select_idx)
                            {
                                case 0:
                                    if (InsControl._tek_scope_en)
                                    {
                                        InsControl._tek_scope.SetMeasureSource(2, 8, "MAXimum"); MyLib.Delay1ms(500);
                                        vmax = InsControl._tek_scope.CHx_Meas_MAX(2, 8);
                                        InsControl._tek_scope.SetMeasureSource(2, 8, "MINImum"); MyLib.Delay1ms(500);
                                        vmin = InsControl._tek_scope.CHx_Meas_MIN(2, 8);
                                    }
                                    else
                                    {
                                        vmax = InsControl._scope.Meas_CH2MAX();
                                        vmin = InsControl._scope.Meas_CH2MIN();
                                    }

                                    break;
                                case 1:
                                    if (InsControl._tek_scope_en)
                                    {
                                        InsControl._tek_scope.SetMeasureSource(3, 8, "MAXimum"); MyLib.Delay1ms(500);
                                        vmax = InsControl._tek_scope.CHx_Meas_MAX(3, 8);
                                        InsControl._tek_scope.SetMeasureSource(3, 8, "MINImum"); MyLib.Delay1ms(500);
                                        vmin = InsControl._tek_scope.CHx_Meas_MIN(3, 8);
                                    }
                                    else
                                    {
                                        vmax = InsControl._scope.Meas_CH3MAX();
                                        vmin = InsControl._scope.Meas_CH3MIN();
                                    }

                                    break;
                                case 2:
                                    if (InsControl._tek_scope_en)
                                    {
                                        InsControl._tek_scope.SetMeasureSource(4, 8, "MAXimum"); MyLib.Delay1ms(500);
                                        vmax = InsControl._tek_scope.CHx_Meas_MAX(4, 8);
                                        InsControl._tek_scope.SetMeasureSource(4, 8, "MINImum"); MyLib.Delay1ms(500);
                                        vmin = InsControl._tek_scope.CHx_Meas_MIN(4, 8);
                                    }
                                    else
                                    {
                                        vmax = InsControl._scope.Meas_CH4MAX();
                                        vmin = InsControl._scope.Meas_CH4MIN();
                                    }

                                    break;
                            }
#if Report_en
                            _sheet.Cells[row, XLS_Table.T] = vmax;
                            _sheet.Cells[row, XLS_Table.U] = vmin;
#endif
                            int meas_cnt = 10;

                            //":MEASure:DELTatime CHANnel1,CHANnel2
                            if (test_parameter.scope_en[0])
                            {
                                if (InsControl._tek_scope_en)
                                {

                                    dt1 = InsControl._tek_scope.MeasureMean(meas_dt1) - test_parameter.offset_time;
                                    sst1 = InsControl._tek_scope.MeasureMean(meas_sst1);
                                    InsControl._tek_scope.SetMeasureSource(2, 8, "HIGH"); MyLib.Delay1ms(500);
                                    vtop = InsControl._tek_scope.MeasureMean(8);
                                    InsControl._tek_scope.SetMeasureSource(2, 8, "LOW"); MyLib.Delay1ms(500);
                                    vbase = InsControl._tek_scope.MeasureMean(8);
                                }
                                else
                                {
                                    dt1 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel2") - test_parameter.offset_time;
                                    sst1 = InsControl._scope.Meas_CH2Rise();
                                    vtop = InsControl._scope.Meas_CH2Top();
                                    vbase = InsControl._scope.Meas_CH2Base();
                                }

                                double calculate_dt = (test_parameter.delay_us_en ? dt1 * Math.Pow(10, 6) : dt1 * Math.Pow(10, 3));
#if Report_en
                                _sheet.Cells[row, XLS_Table.H] = calculate_dt;
                                _sheet.Cells[row, XLS_Table.I] = sst1 * Math.Pow(10, 6);
                                _sheet.Cells[row, XLS_Table.N] = vtop;
                                _sheet.Cells[row, XLS_Table.Q] = vbase;
#endif
                            }

                            if (test_parameter.scope_en[1])
                            {
                                if (InsControl._tek_scope_en)
                                {

                                    dt2 = InsControl._tek_scope.MeasureMean(meas_dt2) - test_parameter.offset_time;
                                    sst2 = InsControl._tek_scope.MeasureMean(meas_sst2);
                                    InsControl._tek_scope.SetMeasureSource(3, 8, "HIGH"); MyLib.Delay1ms(500);
                                    vtop = InsControl._tek_scope.MeasureMean(8);
                                    InsControl._tek_scope.SetMeasureSource(3, 8, "LOW"); MyLib.Delay1ms(500);
                                    vbase = InsControl._tek_scope.MeasureMean(8);
                                }
                                else
                                {
                                    dt2 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel3") - test_parameter.offset_time;
                                    sst2 = InsControl._scope.Meas_CH3Rise();
                                    vtop = InsControl._scope.Meas_CH3Top();
                                    vbase = InsControl._scope.Meas_CH3Base();
                                }

                                double calculate_dt = (test_parameter.delay_us_en ? dt2 * Math.Pow(10, 6) : dt2 * Math.Pow(10, 3));
#if Report_en
                                _sheet.Cells[row, XLS_Table.J] = calculate_dt;
                                _sheet.Cells[row, XLS_Table.K] = sst2 * Math.Pow(10, 6);
                                _sheet.Cells[row, XLS_Table.O] = vtop;
                                _sheet.Cells[row, XLS_Table.R] = vbase;
#endif
                            }

                            if (test_parameter.scope_en[2])
                            {

                                if (InsControl._tek_scope_en)
                                {
                                    dt3 = InsControl._tek_scope.MeasureMean(meas_dt3) - test_parameter.offset_time;
                                    sst3 = InsControl._tek_scope.MeasureMean(meas_sst3);
                                    InsControl._tek_scope.SetMeasureSource(4, 8, "HIGH"); MyLib.Delay1ms(500);
                                    vtop = InsControl._tek_scope.MeasureMean(8);
                                    InsControl._tek_scope.SetMeasureSource(4, 8, "LOW"); MyLib.Delay1ms(500);
                                    vbase = InsControl._tek_scope.MeasureMean(8);
                                }
                                else
                                {
                                    dt3 = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,CHANnel4") - test_parameter.offset_time;
                                    sst3 = InsControl._scope.Meas_CH3Rise();
                                    vtop = InsControl._scope.Meas_CH3Top();
                                    vbase = InsControl._scope.Meas_CH3Base();
                                }

                                double calculate_dt = (test_parameter.delay_us_en ? dt3 * Math.Pow(10, 6) : dt3 * Math.Pow(10, 3));
#if Report_en
                                _sheet.Cells[row, XLS_Table.L] = calculate_dt;
                                _sheet.Cells[row, XLS_Table.M] = sst3 * Math.Pow(10, 6);
                                _sheet.Cells[row, XLS_Table.P] = vtop;
                                _sheet.Cells[row, XLS_Table.S] = vbase;
#endif
                            }

                            double criteria = MyLib.GetCriteria_time(res);
                            criteria = (test_parameter.delay_us_en ? criteria * Math.Pow(10, 6) : criteria * Math.Pow(10, 9));
                            double criteria_up = (test_parameter.judge_percent * criteria) + criteria;
                            double criteria_down = criteria - (test_parameter.judge_percent * criteria);
                            Console.WriteLine(criteria);
                            double value = 0;

#if Report_en
                            switch (select_idx)
                            {
                                case 0:
                                    value = Convert.ToDouble(_sheet.Cells[row, XLS_Table.H].Value);
                                    if (value > criteria_up || value < criteria_down)
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Fail";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.Red;
                                    }
                                    else
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Pass";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.LightGreen;
                                    }
                                    break;
                                case 1:
                                    value = Convert.ToDouble(_sheet.Cells[row, XLS_Table.J].Value);
                                    if (value > criteria_up || value < criteria_down)
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Fail";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.Red;
                                    }
                                    else
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Pass";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.LightGreen;
                                    }
                                    break;
                                case 2:
                                    value = Convert.ToDouble(_sheet.Cells[row, XLS_Table.L].Value);
                                    if (value > criteria_up || value < criteria_down)
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Fail";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.Red;
                                    }
                                    else
                                    {
                                        _sheet.Cells[row, XLS_Table.V] = "Pass";
                                        _range = _sheet.Range["V" + row];
                                        _range.Interior.Color = Color.LightGreen;
                                    }
                                    break;
                            }

                            switch (wave_pos)
                            {
                                case 0:
                                    _sheet.Cells[wave_row, XLS_Table.AA] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.AB] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.AC] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.AD] = "Bin file";
                                    _range = _sheet.Range["AA" + wave_row, "AD" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[wave_row + 1, XLS_Table.AA] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AB] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AC] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AD] = "=G" + row;
                                    _range = _sheet.Range["AA" + (wave_row + 2).ToString(), "AG" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 1:
                                    _sheet.Cells[wave_row, XLS_Table.AL] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.AM] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.AN] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.AO] = "Bin file";
                                    _range = _sheet.Range["AL" + wave_row, "AO" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[wave_row + 1, XLS_Table.AL] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AM] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AN] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AO] = "=G" + row;
                                    _range = _sheet.Range["AL" + (wave_row + 2).ToString(), "AR" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 2:
                                    _sheet.Cells[wave_row, XLS_Table.AW] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.AX] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.AY] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.AZ] = "Bin file";
                                    _range = _sheet.Range["AW" + wave_row, "AZ" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[wave_row + 1, XLS_Table.AW] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AX] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AY] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.AZ] = "=G" + row;
                                    _range = _sheet.Range["AW" + (wave_row + 2).ToString(), "BC" + (wave_row + 16).ToString()];
                                    wave_pos++;
                                    break;
                                case 3:
                                    _sheet.Cells[wave_row, XLS_Table.BH] = "No.";
                                    _sheet.Cells[wave_row, XLS_Table.BI] = "Temp(C)";
                                    _sheet.Cells[wave_row, XLS_Table.BJ] = "Vin(V)";
                                    _sheet.Cells[wave_row, XLS_Table.BK] = "Bin file";
                                    _range = _sheet.Range["BH" + wave_row, "BK" + wave_row];
                                    _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[wave_row + 1, XLS_Table.BH] = "=D" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BI] = "=E" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BJ] = "=F" + row;
                                    _sheet.Cells[wave_row + 1, XLS_Table.BK] = "=G" + row;
                                    _range = _sheet.Range["BH" + (wave_row + 2).ToString(), "BN" + (wave_row + 16).ToString()];
                                    wave_pos = 0; wave_row = wave_row + 19;
                                    break;
                            }

                            MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path + @"\CH" + (select_idx).ToString(), file_name);
#endif
                            row++;
#endif
                            if (InsControl._tek_scope_en)
                            {
                                InsControl._tek_scope.SetRun();
                            }
                            else
                            {
                                InsControl._scope.Root_RUN();
                            }

                            PowerOffEvent();
                        }
                    }
                    // record test finish time
#if Report_en
                    stopWatch.Stop();
                    TimeSpan timeSpan = stopWatch.Elapsed;
                    string str_temp = _sheet.Cells[2, XLS_Table.B].Value;
                    string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
                    str_temp += "\r\n" + time;
                    _sheet.Cells[2, 2] = str_temp;
#endif
                }
            }
            Stop:
            stopWatch.Stop();
#if Report_en
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_DT_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
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
                        GpioOffSelect(test_parameter.gpio_pin);
                    else
                        GpioOnSelect(test_parameter.gpio_pin);
                    break;
                case 1:
                    // rails disable
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { 0x00 });
                    break;
                case 2: // vin trigger
#if Power_en
                    InsControl._power.AutoPowerOff();
#endif
                    break;
            }
        }

    }
}





