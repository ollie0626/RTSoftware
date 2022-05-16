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
            InsControl._scope.TimeScaleMs(test_parameter.time_scale_ms);
            InsControl._scope.TimeBasePositionMs(test_parameter.time_scale_ms * 3);



            InsControl._scope.DoCommand(":FUNCtion1:VERTical AUTO");
            // add new measure method
            double level = InsControl._scope.doQueryNumber(":CHANNEL2:SCALE?");
            InsControl._scope.DoCommand(string.Format(":FUNCTION1:ABSolute CHANNEL{0}", 2));
            InsControl._scope.DoCommand(":FUNCtion1:VERTical MANual");
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
            
            //InsControl._scope.DoCommand(string.Format(":FUNCtion1:VERTical:RANGe {0}", level));
            InsControl._scope.DoCommand(string.Format(":FUNCtion1:VERTical:OFFSet {0}", level * 3));
            InsControl._scope.DoCommand(":MEASure:THResholds:METHod ALL,PERCent");
            InsControl._scope.DoCommand(string.Format(":FUNCtion1:VERTical:RANGe {0}", level));
            //InsControl._scope.DoCommand(":MEASure:THResholds:GENeral:PERCent ALL,100,50,0");

            MyLib.Delay1s(1);
            //InsControl._scope.SetDeltaTime_Rising_to_Rising(1, 1);
            //InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1,FUNC1");
            //InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
            InsControl._scope.CH2_Off();
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
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        //for(int relay_idx = 0; relay_idx < 8; relay_idx++)
                        //{
                        // gpio control for relay
                        //RTDev.RelayOn(RTBBControl.in_gpio_table[relay_idx]);

                        if (test_parameter.run_stop == true) goto Stop;

                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();


                        /* test initial setting */
                        //double hi = 0;
                        //double mid = 0;
                        //double low = 0;
                        //double max = 0;
                        //double min = 0;
                        InsControl._scope.DoCommand(":MARKer:MODE OFF");
                        string file_name;
                        string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V_iout={4:0.##}A",
                                                    row - 22, res, temp,
                                                    test_parameter.VinList[vin_idx],
                                                    test_parameter.IoutList[iout_idx]
                                                    );

                        Mylib.eLoadLevelSwich(InsControl._eload, test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];

                        if (test_parameter.trigger_vin_en)
                        {
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            // rising edge trigger
                            InsControl._scope.SetTrigModeEdge(false);
                            InsControl._scope.TriggerLevel_CH1(test_parameter.VinList[vin_idx] / 2);
                            MyLib.Delay1ms(100);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1ms(250);
                            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                            MyLib.Delay1ms(250);
                        }
                        else
                        {
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            InsControl._scope.Trigger(1);
                            InsControl._scope.TriggerLevel_CH1(1.65);

                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1ms(250);
                            InsControl._scope.Root_STOP();
                            MyLib.Delay1ms(100);
                            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);

                            //InsControl._scope.DoCommand(":TRIGger:MODE Timeout");
                            //InsControl._scope.SetTimeoutCondition(true);
                            //InsControl._scope.SetTimeoutSource(1);
                            //InsControl._scope.SetTimeoutTime(1000 * 7); // unit is ns
                            InsControl._scope.NormalTrigger();
                            InsControl._scope.Root_RUN();
                            MyLib.Delay1ms(500);
                            if (binList[0] != "") RTDev.I2C_WriteBinAndGPIO((byte)(test_parameter.slave), 0x00, binList[bin_idx]);
                            MyLib.Delay1ms(250);
                            //max = InsControl._scope.Meas_CH2MAX();
                            //min = InsControl._scope.Meas_CH2MIN();
                            //if (min < -1.2)
                            //{
                            //    // channel vout is negative
                            //    hi = -0.5;
                            //    mid = min * 0.5;
                            //    low = min * 0.95;

                            //    InsControl._scope.Meas_ThresholdMethod(2, false);
                            //    InsControl._scope.Meas_Absolute(2, hi, mid, low);
                            //    InsControl._scope.SetDeltaTime_Rising_to_Falling(1, 1);
                            //}
                            //else
                            //{
                            //    // channel vout is positive
                            //    hi = max * 0.95;
                            //    mid = hi * 0.5;
                            //    low = 0.2;
                            //    InsControl._scope.Meas_ThresholdMethod(2, false);
                            //    InsControl._scope.Meas_Absolute(2, hi, mid, low);
                            //    InsControl._scope.SetDeltaTime_Rising_to_Rising(1, 1);
                            //}
                            InsControl._scope.SetDeltaTime_Rising_to_Rising(1, 1);
                            //InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1,FUNC1");
                            InsControl._scope.Measure_Clear();
                            MyLib.Delay1ms(500);
                            InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1, FUNC1");
                            InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        }

                        InsControl._scope.Root_STOP();
                        double delay_time, ss_time, Vmax, Inrush;
                        //double ss_time1;
                        //delay_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1,FUNC1") * 1000;
                        //delay_time = InsControl._scope.Meas_DeltaTime(1, 2) * 1000;
                        // add new function
                        delay_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, FUNC1") * 1000;
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_DT");


                        // sst part
                        Vmax = InsControl._scope.Meas_CH2MAX();
                        Inrush = InsControl._scope.Meas_CH4MAX();
                        InsControl._scope.Measure_Clear();
                        MyLib.Delay1s(2);
                        InsControl._scope.Meas_Absolute(2, Vmax * 0.95, Vmax / 2, 0.5);
                        InsControl._scope.SetDeltaTime(true, 1, 0, true, 1, 2);
                        InsControl._scope.DoCommand(":MEASure:DELTatime FUNC1, FUNC1");
                        InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        MyLib.Delay1ms(250);
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_SST");

                        //ss_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? FUNC1,FUNC1") * 1000;
                        MyLib.Delay1ms(250);
                        //ss_time = InsControl._scope.Meas_DeltaTime(2, 2) * 1000;
                        ss_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? FUNC1,FUNC1") * 1000;
                        //ss_time1 = tsClass.CalcSSTime(InsControl._scope);

                        MyLib.Delay1s(1);
                        InsControl._scope.Root_Clear();
                        InsControl._scope.Root_RUN();

                        // gpio control for relay 
                        //RTDev.RelayOff(RTBBControl.out_gpio_table[relay_idx]);
                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.IoutList[iout_idx];
                        _sheet.Cells[row, XLS_Table.E] = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        _sheet.Cells[row, XLS_Table.F] = delay_time;
                        _sheet.Cells[row, XLS_Table.G] = ss_time;
                        _sheet.Cells[row, XLS_Table.H] = Vmax;
                        _sheet.Cells[row, XLS_Table.I] = Inrush;
                        //_sheet.Cells[row, XLS_Table.J] = ss_time1;
                        InsControl._power.PowerOff();
                        MyLib.Delay1ms(500);
                        row++;
                        //}
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
