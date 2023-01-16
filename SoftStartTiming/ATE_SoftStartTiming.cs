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
            int row = 22;
            string[] binList;
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if false
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Bin file";
            _sheet.Cells[row, XLS_Table.E] = "delay time(ms)";
            _sheet.Cells[row, XLS_Table.F] = "Soft Start(us)";
            _sheet.Cells[row, XLS_Table.G] = "Vmax(V)";
            _range = _sheet.Range["A" + row, "E" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range = _sheet.Range["F" + row, "J" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif
            InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1s(1);
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
                            if (test_parameter.run_stop == true) goto Stop;

                            if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                            /* test initial setting */
                            //InsControl._scope.DoCommand(":MARKer:MODE OFF");
                            string file_name;
                            string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);

                            test_parameter.sleep_mode = (res.IndexOf("sleep_en") == -1) ? false : true;

                            InsControl._scope.Measure_Clear();
                            MyLib.Delay1s(1);

                            //:MEASure: RISetime\sCHAN2
                            for (int i = 0; i < test_parameter.scope_en.Length; i++)
                            {
                                if (test_parameter.scope_en[i])
                                    InsControl._scope.DoCommand(":MEASure:RISetime CHANnel" + (i + 2).ToString());
                            }


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

                            switch (select_idx)
                            {
                                case 0:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel2");
                                    break;
                                case 1:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel3");
                                    break;
                                case 2:
                                    delay_time_res = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel4");
                                    break;
                            }

                            if (delay_time_res >= time_scale * 4)
                            {
                                if(delay_time_res > 0)
                                {
                                    double temp = (delay_time_res * 1.2) / 4;
                                    InsControl._scope.TimeScale(temp);
                                    InsControl._scope.TimeBasePosition(temp * 3);
                                    //InsControl._scope.TimeBasePosition((delay_time_res * 1.2) / 4) * 3);
                                }
                                else
                                {
                                    InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                                    InsControl._scope.TimeBasePosition(test_parameter.ontime_scale_ms * 3);
                                }
                                InsControl._scope.Root_RUN();

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
                                goto retest;
                            }
                            else if (delay_time_res <= time_scale)
                            {
                                InsControl._scope.TimeScale(delay_time_res);
                                InsControl._scope.TimeBasePosition(delay_time_res * 3);
                                InsControl._scope.Root_RUN();

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
                                goto retest;
                            }

                            InsControl._scope.SaveWaveform(@"D:\", res);


                            // need to judge select which channel to re-scale time scale.
                            //if (delay_time_res >= time_scale_threshold)
                            //{
                            //    InsControl._scope.TimeScale((delay_time_res * 1.2));
                            //    InsControl._scope.TimeBasePosition(((delay_time_res * 1.2)) * 3);
                            //}

                            InsControl._scope.Root_RUN();
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
    }
}





