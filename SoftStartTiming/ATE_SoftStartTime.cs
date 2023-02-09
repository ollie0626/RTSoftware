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

    public class ATE_SoftStartTime : TaskRun
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

        public ATE_SoftStartTime()
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
                InsControl._tek_scope.SetTimeBasePosition(25);
                InsControl._tek_scope.SetRun();
                InsControl._tek_scope.SetTriggerMode();
                InsControl._tek_scope.SetTriggerSource(1);
                InsControl._tek_scope.SetTriggerLevel(1.5);

                InsControl._tek_scope.CHx_On(1);
                InsControl._tek_scope.CHx_On(2);
                InsControl._tek_scope.CHx_On(3);
                InsControl._tek_scope.CHx_On(4);

                InsControl._tek_scope.CHx_Level(1, 1.65);
                InsControl._tek_scope.CHx_Level(2, test_parameter.VinList[0] * 3);
                InsControl._tek_scope.CHx_Level(3, test_parameter.LX_Level);
                InsControl._tek_scope.CHx_Level(4, test_parameter.ILX_Level);

                InsControl._tek_scope.CHx_Position(1, 0);
                InsControl._tek_scope.CHx_Position(2, -1);
                InsControl._tek_scope.CHx_Position(3, -2);
                InsControl._tek_scope.CHx_Position(4, -3);

                InsControl._tek_scope.CHx_BWlimitOn(1);
                InsControl._tek_scope.CHx_BWlimitOn(2);
                InsControl._tek_scope.CHx_BWlimitOn(3);
                InsControl._tek_scope.CHx_BWlimitOn(4);

                MyLib.Delay1ms(500);
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

                InsControl._scope.CHx_On(2);
                InsControl._scope.CHx_Level(2, test_parameter.VinList[0] * 3);
                InsControl._scope.CHx_Offset(2, 0);

                InsControl._scope.CH1_BWLimitOn();
                InsControl._scope.CH2_BWLimitOn();
                InsControl._scope.CH3_BWLimitOn();
                InsControl._scope.CH4_BWLimitOn();

                InsControl._scope.TriggerLevel_CH1(1);
                InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:METHod ALL,PERCent");
                InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:PERCent ALL,100,50,1");

                InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:METHod ALL,PERCent");
                InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent ALL,100,50,1");
                InsControl._scope.Root_RUN();
                MyLib.Delay1ms(1000);
                MyLib.WaveformCheck();
                // measure current delta-time.
                InsControl._scope.DoCommand(":MEASure:STATistics CURRent");
            }
        }

        private void Scope_Channel_Resize(int idx, string path)
        {
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

            //InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
            MyLib.Delay1ms(800);

            double time_scale = 0;
            if (InsControl._tek_scope_en)
            {
                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
            }
            else
            {
                time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
            }

            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTimeScale(Math.Pow(10, -9) * 40);
            }
            else
            {
                InsControl._scope.TimeScaleUs(1);
            }


            RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, path); // test conditions
            MyLib.Delay1ms(800);

            switch (test_parameter.trigger_event)
            {
                case 0: // gpio

                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTriggerLevel(1);
                        InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                        InsControl._tek_scope.CHx_Position(1, 0);
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
                    double vout = InsControl._scope.Meas_CH1MAX();
                    InsControl._scope.TriggerLevel_CH1(vout * 0.35);
                    break;
                case 2: // vin trigger
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
                    if (InsControl._tek_scope_en)
                        InsControl._tek_scope.SetTriggerLevel(test_parameter.VinList[idx] * 0.35);
                    else
                        InsControl._scope.TriggerLevel_CH1(test_parameter.VinList[idx] * 0.35);
                    break;
            }
            MyLib.Delay1s(1);


            if (InsControl._tek_scope_en)
            {
                InsControl._tek_scope.CHx_Level(2, test_parameter.VinList[0] * 3);
                InsControl._tek_scope.CHx_Level(3, test_parameter.LX_Level);
                InsControl._tek_scope.CHx_Level(4, test_parameter.ILX_Level);

            }
            else
            {
                // CH2
                InsControl._scope.CHx_Level(2, test_parameter.VinList[0] * 3);
                InsControl._scope.CHx_Offset(2, test_parameter.VinList[0] * 3);

                // CH3 LX
                InsControl._scope.CHx_Level(3, test_parameter.LX_Level);
                InsControl._scope.CHx_Offset(3, 0);

                // CH4 ILX
                InsControl._scope.CHx_Level(4, test_parameter.ILX_Level);
                InsControl._scope.CHx_Offset(4, 0);
            }

            MyLib.Delay1s(3);

            int re_cnt = 0;
            for (int ch_idx = 0; ch_idx < 3; ch_idx++)
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
                    vmax = InsControl._tek_scope.CHx_Meas_MAX(ch_idx + 2, 1);
                }
                else
                {
                    vmax = InsControl._scope.Measure_Ch_Max(ch_idx + 2);
                }

                if (vmax > Math.Pow(10, 9))
                {
                    re_cnt++;

                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.CHx_Level(ch_idx + 2, test_parameter.VinList[0] * 3);
                    }
                    else
                    {
                        InsControl._scope.CHx_Level(ch_idx + 2, test_parameter.VinList[0] * 3);
                        InsControl._scope.CHx_Offset(ch_idx + 2, test_parameter.VinList[0] * 3 * (ch_idx + 1));
                    }

                    MyLib.Delay1ms(800);
                    goto re_scale;
                }

                if(InsControl._tek_scope_en)
                {
                    InsControl._tek_scope.CHx_Level(ch_idx + 2, vmax / 2.5);
                }
                else
                {
                    InsControl._scope.CHx_Level(ch_idx + 2, vmax / 2.5);
                    InsControl._scope.CHx_Offset(ch_idx + 2, (vmax / 2.5) * (ch_idx + 1));
                }
                MyLib.Delay1ms(800);
            }


            double trigger_level = 0;
            if(InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTriggerSource(2);
                trigger_level = InsControl._tek_scope.CHx_Meas_MAX(2, 1) * 0.1;
                InsControl._tek_scope.SetTriggerLevel(trigger_level);
            }
            else
            {
                InsControl._scope.Trigger_CH2();
                trigger_level = InsControl._scope.Meas_CH2MAX() * 0.3;
                InsControl._scope.TriggerLevel_CH2(trigger_level);
            }



            PowerOffEvent();

            if(InsControl._tek_scope_en)
            {
                InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                InsControl._tek_scope.SetTimeBasePosition(25);
            }
            else
            {
                InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms);
            }


            MyLib.Delay1ms(250);
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

#if true
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
            #region "Report initial"
#if true
            _sheet = _book.Worksheets.Add();
            _sheet.Name = "SST Test";
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
            _sheet.Cells[1, XLS_Table.B] = "Soft-Start time";
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions + test_parameter.bin_file_cnt;


            _sheet.Cells[row, XLS_Table.D] = "No.";
            _sheet.Cells[row, XLS_Table.E] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.F] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.G] = "Bin file";
            _range = _sheet.Range["D" + row, "G" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // major measure timing
            _sheet.Cells[row, XLS_Table.H] = "SST (us)";
            _sheet.Cells[row, XLS_Table.I] = "V1 Max (V)";
            _sheet.Cells[row, XLS_Table.J] = "V1 Min (V)";
            _sheet.Cells[row, XLS_Table.K] = "ILx Max (mA)";
            _sheet.Cells[row, XLS_Table.L] = "ILx Min (mA)";
            _sheet.Cells[row, XLS_Table.M] = "Pass/Fail";

            _range = _sheet.Range["H" + row, "L" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["M" + row, "M" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            row++;
#endif
            #endregion

            stopWatch.Start();
            binList = MyLib.ListBinFile(test_parameter.bin_path[0]);
            bin_cnt = binList.Length;
            cnt = 0;

            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {

                if(InsControl._tek_scope_en)
                {
                    InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                    InsControl._tek_scope.SetTimeBasePosition(25);
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
                    if(!InsControl._tek_scope_en) InsControl._scope.DoCommand(":MARKer:MODE OFF");
                    string file_name;
                    string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                    test_parameter.sleep_mode = (res.IndexOf("sleep_en") == -1) ? false : true;
                    if(!InsControl._tek_scope_en) InsControl._scope.Measure_Clear();
                    MyLib.Delay1s(1);

                    // Call Measure display to waveform
                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.CHx_Meas_Rise(2, 1); // meas1 soft-start time
                        InsControl._tek_scope.CHx_Meas_MAX(2, 2);  // meas2 vout max
                        InsControl._tek_scope.CHx_Meas_MIN(2, 3);  // meas3 vout min
                        InsControl._tek_scope.CHx_Meas_MIN(4, 4);  // meas4 ILx max
                        InsControl._tek_scope.CHx_Meas_MIN(4, 5);  // meas5 ILX min
                    }
                    else
                    {
                        InsControl._scope.DoCommand(":MEASure:VMAX CHANnel4"); // measure5 ILx max
                        InsControl._scope.DoCommand(":MEASure:VMIN CHANnel4"); // measure4 ILx min
                        InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2"); // measure3 Vout max
                        InsControl._scope.DoCommand(":MEASure:VMIN CHANnel2"); // measure2 Vout min
                        InsControl._scope.DoCommand(":MEASure:RISetime CHANnel" + (2).ToString()); // measure1
                        InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement MEAS1");
                    }

                    MyLib.Delay1ms(500);

                    Console.WriteLine(res);
                    file_name = string.Format("{0}_Temp={2}C_vin={3:0.##}V_{1}",
                                                cnt, res, temp,
                                                test_parameter.VinList[vin_idx]
                                                );

                    double time_scale = 0;
                    if(InsControl._tek_scope_en)
                    {
                        time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
                    }
                    else
                    {
                        time_scale = InsControl._scope.doQueryNumber(":TIMebase:SCALe?");
                    }
                    
                    // include test condition
                    Scope_Channel_Resize(vin_idx, binList[bin_idx]);
                    double tempVin = ori_vinTable[vin_idx];

                    MyLib.WaveformCheck();
                    MyLib.Delay1ms(800);

                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTriggerMode(false);
                        InsControl._tek_scope.SetTriggerRise();
                        InsControl._tek_scope.SetClear();
                    }
                    else
                    {
                        InsControl._scope.NormalTrigger();
                        InsControl._scope.SetTrigModeEdge(false);
                        InsControl._scope.Root_Clear();
                    }

                    MyLib.Delay1ms(1500);

                    // power on trigger
                    switch (test_parameter.trigger_event)
                    {
                        case 0:
                            // GPIO trigger event
                            if (test_parameter.sleep_mode)
                            {
                                GpioOnSelect(test_parameter.gpio_pin);
                                MyLib.Delay1ms(1000);
                            }
                            else
                            {
                                GpioOffSelect(test_parameter.gpio_pin);
                                MyLib.Delay1ms(1000);
                            }
                            time_scale = time_scale * 1000;
                            break;
                        case 1:
                            // I2C trigger event
                            break;
                        case 2:
                            // Power supply trigger event
                            InsControl._power.AutoPowerOff();
                            break;
                    }

                    MyLib.Delay1s(1);
                    if (InsControl._tek_scope_en) InsControl._tek_scope.SetStop();
                    else InsControl._scope.Root_STOP();
                    MyLib.Delay1ms(1000);

                    double delay_time = 0;
                    if(InsControl._tek_scope_en) delay_time = InsControl._tek_scope.CHx_Meas_Rise(2, 1);
                    else delay_time = InsControl._scope.Meas_CH2Rise();

                    double temp_time = 0;
                    if (InsControl._tek_scope_en)
                        temp_time = (delay_time / 4);
                    else
                        temp_time = (delay_time) / 3.5;

                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetTimeScale(temp_time);
                    }
                    else
                    {
                        InsControl._scope.TimeScale(temp_time);
                        InsControl._scope.TimeBasePosition(temp_time * 1);
                    }

                    PowerOffEvent();

                    MyLib.Delay1ms(1000);
                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetRun();
                        InsControl._tek_scope.SetClear();
                    }
                    else
                    {
                        InsControl._scope.Root_RUN();
                        InsControl._scope.Root_Clear();
                    }

                    MyLib.Delay1ms(1500);
                    switch (test_parameter.trigger_event)
                    {
                        case 0:
                            // GPIO trigger event
                            if (test_parameter.sleep_mode)
                            {
                                GpioOnSelect(test_parameter.gpio_pin);
                                MyLib.Delay1ms(1000);
                            }
                            else
                            {
                                GpioOffSelect(test_parameter.gpio_pin);
                                MyLib.Delay1ms(1000);
                            }
                            break;
                        case 1:
                            // I2C trigger event
                            break;
                        case 2:
                            // Power supply trigger event
                            InsControl._power.AutoPowerOff();
                            break;
                    }
                    MyLib.Delay1s(1);

                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SetStop();
                    }
                    else
                    {
                        InsControl._scope.Root_STOP();
                    }

#if true
                    double vin = 0;
                    double sst = 0, vmax = 0, vmin = 0, ilx_max = 0, ilx_min = 0;

                    if(InsControl._tek_scope_en)
                    {
                        //vin         = InsControl._power.GetVoltage();
                        sst         = InsControl._tek_scope.CHx_Meas_Rise(2, 1) * Math.Pow(10, 6);
                        vmax        = InsControl._tek_scope.CHx_Meas_MAX(2, 2);  
                        vmin        = InsControl._tek_scope.CHx_Meas_MIN(2, 3);  
                        ilx_max     = InsControl._tek_scope.CHx_Meas_MIN(4, 4);  
                        ilx_min     = InsControl._tek_scope.CHx_Meas_MIN(4, 5);

                        InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                        MyLib.Delay1ms(100);

                        InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 0");
                        MyLib.Delay1ms(100);
                        double data = InsControl._tek_scope.CHx_Meas_Rise(2, 1) * 0.9;
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + data.ToString());
                    }
                    else
                    {
                        double[] data = InsControl._scope.doQeury(":MEASure:RESults?").Split(',').Select(double.Parse).ToArray();
                        vin = InsControl._power.GetVoltage();
                        sst = data[0] * Math.Pow(10, 6);
                        vmax = data[1];
                        vmin = data[2];
                        ilx_max = data[3] * Math.Pow(10, -3);
                        ilx_min = data[4] * Math.Pow(10, -3);
                    }
                    MyLib.Delay1s(1);
                    if (InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name);
                    }
                    else
                    {
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);
                    }
                    _sheet.Cells[row, XLS_Table.D] = cnt++;
                    _sheet.Cells[row, XLS_Table.E] = temp;
                    _sheet.Cells[row, XLS_Table.F] = vin;
                    _sheet.Cells[row, XLS_Table.G] = res;

                    _sheet.Cells[row, XLS_Table.H] = sst;
                    _sheet.Cells[row, XLS_Table.I] = vmax;
                    _sheet.Cells[row, XLS_Table.J] = vmin;
                    _sheet.Cells[row, XLS_Table.K] = ilx_max;
                    _sheet.Cells[row, XLS_Table.L] = ilx_min;

                    double criteria = MyLib.GetCriteria_time(res);
                    criteria = criteria * Math.Pow(10, 6);
                    double criteria_up = (test_parameter.judge_percent * criteria) + criteria;
                    double criteria_down = criteria - (test_parameter.judge_percent * criteria);
                    Console.WriteLine(criteria);


                    if (sst > criteria_up || sst < criteria_down)
                    {
                        _sheet.Cells[row, XLS_Table.M] = "Fail";
                        _range = _sheet.Range["M" + row];
                        _range.Interior.Color = Color.Red;
                    }
                    else
                    {
                        _sheet.Cells[row, XLS_Table.M] = "Pass";
                        _range = _sheet.Range["M" + row];
                        _range.Interior.Color = Color.LightGreen;
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
                            _range = _sheet.Range["AA" + (wave_row + 2).ToString(), "AI" + (wave_row + 16).ToString()];
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
                            _range = _sheet.Range["AL" + (wave_row + 2).ToString(), "AT" + (wave_row + 16).ToString()];
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
                            _range = _sheet.Range["AW" + (wave_row + 2).ToString(), "BE" + (wave_row + 16).ToString()];
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
                            _range = _sheet.Range["BH" + (wave_row + 2).ToString(), "BP" + (wave_row + 16).ToString()];
                            wave_pos = 0; wave_row = wave_row + 19;
                            break;
                    }

                    MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name);
                    row++;
#endif
                    if (InsControl._tek_scope_en) InsControl._tek_scope.SetRun();
                    else InsControl._scope.Root_RUN();

                    PowerOffEvent();
                }
            }
        Stop:
            stopWatch.Stop();
            // record test finish time
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, XLS_Table.B].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            //TimeSpan timeSpan = stopWatch.Elapsed;
#if true
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_SST_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
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
                    break;
                case 2: // vin trigger
                    InsControl._power.AutoPowerOff();
                    break;
            }
        }

    }
}





