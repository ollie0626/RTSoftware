using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;

namespace IN528ATE_tool
{
    public class ATE_PowerOn : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        //Excel.Chart _chart;

        public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_PowerOn()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void OSCInit()
        {
            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
            InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
            InsControl._scope.DoCommand(":FUNCtion1:VERTical AUTO");
            double level = InsControl._scope.doQueryNumber(":CHANNEL2:SCALE?");
            InsControl._scope.DoCommand(string.Format(":FUNCTION1:ABSolute CHANNEL{0}", 2));
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");

            InsControl._scope.CH2_On();
            InsControl._scope.CH3_On();
            InsControl._scope.CH2_Level(6);
            InsControl._scope.CH3_Level(6);


            //InsControl._scope.DoCommand(":MEASure:THResholds:METHod ALL,PERCent");
            //InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent ALL,100,50,0");
            //InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:PERCent ALL,100,50,0");
            InsControl._scope.TriggerLevel_CH1(test_parameter.trigger_level); // gui trigger level
            InsControl._scope.NormalTrigger();
            RTDev.GpEn_Disable();
            InsControl._scope.Root_RUN();
            MyLib.Delay1s(1);
        }


        private void Scope_Channel_Resize(int idx, string path)
        {
            InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
            MyLib.Delay1ms(250);
            // write default bin file
            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
            MyLib.Delay1ms(100);
            RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, path); // test conditions
            MyLib.Delay1ms(250);
            // program test conditons 
            if (test_parameter.mtp_enable)
            {
                byte[] buf = new byte[] { test_parameter.mtp_data };
                RTDev.I2C_Write((byte)(test_parameter.mtp_slave >> 1), test_parameter.mtp_addr, buf);
            }
            MyLib.Delay1ms(250);
            InsControl._eload.CH1_Loading(0.01);
            InsControl._scope.AutoTrigger();
            // inital channel level setting
            if(test_parameter.trigger_vin_en)
            {
                //double vin = test_parameter.VinList[idx];
                InsControl._scope.CH1_Level(test_parameter.trigger_level);
            }
            else if(test_parameter.trigger_en)
            {
                InsControl._scope.CH1_Level(1);
            }

            InsControl._scope.CH2_Level(6);
            InsControl._scope.CH3_Level(6);
            for (int i = 0; i < 3; i++)
            {
                // Inrush ???
                //InsControl._scope.CH4_Level(1);
                double Vo;
                Vo = Math.Abs(InsControl._scope.Meas_CH2MAX());
                MyLib.Delay1ms(100);
                InsControl._scope.CH2_Level(Vo / 2);
                Vo = Math.Abs(InsControl._scope.Meas_CH3MAX());
                MyLib.Delay1ms(100);
                InsControl._scope.CH3_Level(Vo / 2);
                // Inrush ????
                //Vo = Math.Abs(InsControl._scope.Meas_CH4MAX());
                //InsControl._scope.CH4_Level(Vo / 2);
            }

            InsControl._power.AutoPowerOff();
            MyLib.Delay1ms(100);
            InsControl._eload.AllChannel_LoadOff();
            MyLib.Delay1ms(100);
            InsControl._scope.NormalTrigger();
            MyLib.Delay1ms(300);
        }



        public void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            RTDev.GpioInit();

            int vin_cnt = test_parameter.VinList.Count;
            int iout_cnt = test_parameter.IoutList.Count;
            int row = 22;
            string[] binList;
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;
            binList = Mylib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Delay Time & Soft-start Time", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Bin file";
            _sheet.Cells[row, XLS_Table.F] = "delay time(ms)";
            _sheet.Cells[row, XLS_Table.G] = "Soft Start(us)";
            _sheet.Cells[row, XLS_Table.H] = "Vmax(V)";
            _sheet.Cells[row, XLS_Table.I] = "Inrush(mA)";
            _sheet.Cells[row, XLS_Table.J] = "delay time(ms)_cal";
            _range = _sheet.Range["A" + row, "E" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range = _sheet.Range["F" + row, "J" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
            InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1s(1);
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;

                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        //Scope_Channel_Resize(vin_idx);

                        /* test initial setting */
                        InsControl._scope.DoCommand(":MARKer:MODE OFF");
                        string file_name;
                        string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V_iout={4:0.##}A",
                                                    row - 22, res, temp,
                                                    test_parameter.VinList[vin_idx],
                                                    test_parameter.IoutList[iout_idx]
                                                    );
                        Scope_Channel_Resize(vin_idx, binList[bin_idx]);
                        

                        Mylib.eLoadLevelSwich(InsControl._eload, test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];
                        InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                        InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
                        System.Threading.Thread.Sleep(1000);

                        if (test_parameter.trigger_vin_en)
                        {
                            // vin trigger
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            // rising edge trigger
                            InsControl._scope.SetTrigModeEdge(false);
                            MyLib.Delay1ms(100);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1ms(500);
                            //if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                            //MyLib.Delay1ms(150);
                            
                        }
                        else if (test_parameter.trigger_en)
                        {
                            //Gpio 2.0 trigger
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            InsControl._scope.SetTrigModeEdge(false);
                            InsControl._scope.TriggerLevel_CH1(1.5);
                            MyLib.Delay1ms(100);
                            RTDev.GpEn_Enable();
                            MyLib.Delay1ms(250);
                            //if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                            //MyLib.Delay1ms(150);
                            //RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, binList[bin_idx]);
                            //MyLib.Delay1ms(250);
                            //if (test_parameter.mtp_enable)
                            //{
                            //    byte[] buf = new byte[] { test_parameter.mtp_data };
                            //    RTDev.I2C_Write((byte)(test_parameter.mtp_slave >> 1), test_parameter.mtp_addr, buf);
                            //}
                        }
                        else
                        {
                            // I2c run and GPIO trigger
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            InsControl._scope.Trigger(1);
                            InsControl._scope.TriggerLevel_CH1(1.65);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1ms(250);
                            InsControl._scope.Root_STOP();
                            MyLib.Delay1ms(100);
                            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                            InsControl._scope.NormalTrigger();
                            InsControl._scope.Root_RUN();

                            MyLib.Delay1ms(500);
                            if (binList[0] != "") RTDev.I2C_WriteBinAndGPIO((byte)(test_parameter.slave), 0x00, binList[bin_idx]);
                            MyLib.Delay1ms(250);
                            InsControl._scope.Measure_Clear();
                            MyLib.Delay1ms(800);
                        }
                        InsControl._scope.Root_STOP();

                        double delay_time, ss_time, Vmax, Inrush;

                        if(test_parameter.trigger_vin_en)
                        {
                            // adjust CH1 level to Vout 10mV
                            // measure UVLO to Vout 10mV
                            // measure thresholds method is abs
                            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                            InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
                            InsControl._scope.DoCommand(":MEASure:THResholds:METHod CHANnel1,ABSolute");
                            InsControl._scope.DoCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute CHANnel1,{0},{1},{2}",
                                                        InsControl._scope.Meas_CH1Top(),
                                                        test_parameter.measure_level,
                                                        0));
                            InsControl._scope.DoCommand(":MEASure:THResholds:METHod FUNC1,ABSolute");
                            InsControl._scope.DoCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute FUNC1,{0},{1},{2}",
                                                        InsControl._scope.doQueryNumber(":MEASure:VTOP? FUNC1"),
                                                        InsControl._scope.doQueryNumber(":MEASure:VTOP? FUNC1") * 0.5,
                                                        0.05));
                            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:METHod ALL,PERCent");
                            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent FUNC1,100,50,0");

                            System.Threading.Thread.Sleep(150);
                        }


                        // delay time and sst measure
                        InsControl._scope.Measure_Clear();
                        MyLib.Delay1s(1);
                        // MEAS2
                        InsControl._scope.SetDeltaTime_Rising_to_Rising(1, 1);
                        InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1, FUNC1");


                        // MEAS1
                        InsControl._scope.SetDeltaTime(true, 1, 0, true, 1, 2);
                        InsControl._scope.DoCommand(":MEASure:DELTatime FUNC1, FUNC1");
                        

                        delay_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, FUNC1") * 1000;
                        ss_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? FUNC1, FUNC1") * 1000;
                        Vmax = InsControl._scope.Meas_CH2MAX();
                        Inrush = InsControl._scope.Meas_CH4MAX();

                        InsControl._scope.DoCommand(":MARKer:MODE MANual");
                        InsControl._scope.DoCommand(":MARKer3:ENABle OFF");
                        InsControl._scope.DoCommand(":MARKer4:ENABle OFF");
                        InsControl._scope.DoCommand(":MARKer3:TYPE XMANual");
                        InsControl._scope.DoCommand(":MARKer4:TYPE XMANual");
                        InsControl._scope.DoCommand(":MARKer3:ENABle ON");
                        InsControl._scope.DoCommand(":MARKer4:ENABle ON");
                        InsControl._scope.DoCommand(":MARKer1:DELTa MARKer2, ON");
                        InsControl._scope.DoCommand(":MARKer4:DELTa MARKer3, ON");
                        InsControl._scope.DoCommand(":MARKer3:SOURce CHANnel2");
                        InsControl._scope.DoCommand(":MARKer4:SOURce CHANnel2");
                        InsControl._scope.DoCommand(string.Format(":MARKer1:X:POSition 0"));
                        InsControl._scope.DoCommand(string.Format(":MARKer2:X:POSition {0}", delay_time / 1000));
                        InsControl._scope.DoCommand(string.Format(":MARKer3:X:POSition {0}", delay_time / 1000));
                        InsControl._scope.DoCommand(string.Format(":MARKer4:X:POSition {0}", (delay_time + ss_time) / 1000));
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_ON");
                        InsControl._scope.DoCommand(":MARKer:MODE OFF");

                        InsControl._scope.Measure_Clear();
                        MyLib.Delay1s(1);
                        InsControl._scope.Root_Clear();
                        InsControl._scope.Root_RUN();

                        // gpio control for relay 
                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.IoutList[iout_idx];
                        _sheet.Cells[row, XLS_Table.E] = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        _sheet.Cells[row, XLS_Table.F] = delay_time;
                        _sheet.Cells[row, XLS_Table.G] = ss_time;
                        _sheet.Cells[row, XLS_Table.H] = Vmax;
                        _sheet.Cells[row, XLS_Table.I] = Inrush;

                        InsControl._scope.Measure_Clear();
                        InsControl._scope.TimeScaleMs(test_parameter.offtime_scale_ms);
                        InsControl._scope.TimeBasePositionMs(test_parameter.offtime_scale_ms * 1);
                        System.Threading.Thread.Sleep(1000);
                        InsControl._scope.NormalTrigger();
                        InsControl._scope.SetTrigModeEdge(true);
                        InsControl._scope.Root_RUN();
                        System.Threading.Thread.Sleep(1000);

                        // power off section
                        if (test_parameter.trigger_vin_en)
                        {
                            InsControl._power.AutoPowerOff();
                            System.Threading.Thread.Sleep(250);
                            RTDev.GpEn_Disable();
                        }
                        else if (test_parameter.trigger_en)
                        {
                            RTDev.GpEn_Disable();
                            System.Threading.Thread.Sleep(250);
                            InsControl._power.AutoPowerOff();
                        }
                        System.Threading.Thread.Sleep(800);
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_OFF");
                        MyLib.Delay1s(2);
                        row++;
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;

            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_DT_SST_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
            if (!test_parameter.all_en && !test_parameter.chamber_en) delegate_mess.Invoke();

        }
    }
}
