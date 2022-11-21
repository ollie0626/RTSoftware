﻿using System;
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
    public class ATE_CodeInrush : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        MyLib MyLib;
        RTBBControl RTDev = new RTBBControl();

        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_CodeInrush()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Code Inrush test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        public void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            System.Threading.Thread.Sleep(2000);

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();

            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH4_On();
            InsControl._scope.CH4_1Mohm();

            double level_max = Math.Abs(test_parameter.vol_max) > Math.Abs(test_parameter.vol_min) ? Math.Abs(test_parameter.vol_max) : Math.Abs(test_parameter.vol_min);
            double level_min = Math.Abs(test_parameter.vol_max) < Math.Abs(test_parameter.vol_min) ? Math.Abs(test_parameter.vol_max) : Math.Abs(test_parameter.vol_min);
            bool neg_vol = test_parameter.vol_min < 0;
            // -3, -6
            double ch_level = (level_max - level_min) / 4;
            InsControl._scope.CH1_Level(ch_level);
            InsControl._scope.CH4_Level(0.2);

            InsControl._scope.CH4_Offset(0.2 * 3);
            InsControl._scope.CH1_Offset(neg_vol ? (level_min + (ch_level * 3)) * -1 : level_min + (ch_level * 3));

            System.Threading.Thread.Sleep(1000);
            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
            System.Threading.Thread.Sleep(1000);

            //System.Threading.Thread.Sleep(1000);
            double trigger_level = neg_vol ? (level_max * 0.8) * -1 : level_max * 0.8;
            InsControl._scope.TriggerLevel_CH1(trigger_level);
            System.Threading.Thread.Sleep(500);
            //InsControl._scope.Trigger_CH2();
            //InsControl._scope.TriggerLevel_CH2(1.2);

            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:METHod ALL,PERCent");
            InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent ALL,100,50,0");
        }


        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();



            MyLib = new MyLib();
            int row = 11;
            int idx = 0;
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            bool ispos = test_parameter.vol_max > test_parameter.vol_min;
            int vin_cnt = test_parameter.vinList.Count;
            int iout_cnt = test_parameter.ioutList.Count;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);

            RTDev.BoadInit();
#if Report
            //MyLib.ExcelReportInit(_sheet);
            //MyLib.testCondition(_sheet, "Code Inrush", bin_cnt, temp);
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
            _sheet.Cells[row, XLS_Table.E] = test_parameter.i2c_enable ? "Bin File" : "Swire";
            _sheet.Cells[row, XLS_Table.F] = "Iout (mA)";
            _sheet.Cells[row, XLS_Table.G] = "Imax(mA)_min";
            _sheet.Cells[row, XLS_Table.H] = "Vmax(V)_min";
            _sheet.Cells[row, XLS_Table.I] = "Vmin(V)_min";
            _sheet.Cells[row, XLS_Table.J] = "Imax(mA)_max";
            _sheet.Cells[row, XLS_Table.K] = "Vmax(V)_max";
            _sheet.Cells[row, XLS_Table.L] = "Vmin(V)_max";
            _range = _sheet.Range["A" + row, "L" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["A" + row.ToString(), "F" + row.ToString()];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["G" + row.ToString(), "L" + row.ToString()];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif

            OSCInit();
            InsControl._power.AutoPowerOff();

            for (int bin_idx = 0; bin_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); bin_idx++)
            {
                for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        string res = test_parameter.i2c_enable ? Path.GetFileNameWithoutExtension(binList[bin_idx]) : "Swire_" + test_parameter.code_min + "_" + test_parameter.code_max;
                        string file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V_iout={4:0.##}A",
                                                        row - 22, res, temp,
                                                        test_parameter.vinList[vin_idx],
                                                        test_parameter.ioutList[iout_idx]);
                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        System.Threading.Thread.Sleep(500);
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("34970 沒有連結 !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }
                        
                        if (binList[0] != "" && test_parameter.i2c_enable) RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                        else
                        {
                            // ic setting
                            int[] pulse_tmp = test_parameter.swireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTDev.SwirePulse(pulse_tmp[pulse_idx]);
                        }

                        /* test conditonss */
                        byte[] buf_min = new byte[1] { (byte)test_parameter.code_min };
                        byte[] buf_max = new byte[1] { (byte)test_parameter.code_max };


                        double max, min, vin, iin, imax, iout;
                        vin = InsControl._34970A.Get_100Vol(1);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout() * 1000;
#if Report
                        _sheet.Cells[row, XLS_Table.A] = idx;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = vin;
                        _sheet.Cells[row, XLS_Table.D] = iin * 1000;
                        _sheet.Cells[row, XLS_Table.E] = test_parameter.i2c_enable ? Path.GetFileNameWithoutExtension(binList[bin_idx]) : "setting: " + test_parameter.swireList[bin_idx] + "_Channel pulse: " + test_parameter.code_min + "→" + test_parameter.code_max;
                        _sheet.Cells[row, XLS_Table.F] = iout;
#endif
                        /* min to max code */
                        InsControl._scope.Root_RUN();
                        if (ispos) InsControl._scope.SetTrigModeEdge(false);
                        else InsControl._scope.SetTrigModeEdge(true);
                        InsControl._scope.NormalTrigger();

                        if (test_parameter.i2c_enable)
                        {
                            RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.addr, ispos ? buf_min : buf_max);
                            System.Threading.Thread.Sleep(500);
                            RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.addr, ispos ? buf_max : buf_min);
                            System.Threading.Thread.Sleep(2000);
                        }
                        else
                        {
                            RTDev.SwirePulse(ispos ? test_parameter.code_min : test_parameter.code_max);
                            System.Threading.Thread.Sleep(500);
                            RTDev.SwirePulse(ispos ? test_parameter.code_max : test_parameter.code_min);
                            System.Threading.Thread.Sleep(2000);
                        }

                        InsControl._scope.Root_STOP();
                        InsControl._scope.Measure_Clear();
                        InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        InsControl._scope.DoCommand(":MEASure:RISetime CHANnel1");
                        InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement MEASurement1");
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_min");

                        imax = InsControl._scope.Meas_CH4MAX();
                        max = InsControl._scope.Meas_CH1MAX();
                        min = InsControl._scope.Meas_CH1MIN();
#if Report
                        _sheet.Cells[row, XLS_Table.G] = imax * 1000;
                        _sheet.Cells[row, XLS_Table.H] = max;
                        _sheet.Cells[row, XLS_Table.I] = min;
#endif
                        InsControl._scope.Root_Clear();
                        System.Threading.Thread.Sleep(2000);

                        /* max to min code */
                        InsControl._scope.SetTrigModeEdge(true);
                        //if (ispos) InsControl._scope.SetTrigModeEdge(true);
                        //else InsControl._scope.SetTrigModeEdge(false);

                        //InsControl._scope.SetTrigModeEdge(true);
                        InsControl._scope.Root_RUN();
                        //RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.addr, buf_max);
                        System.Threading.Thread.Sleep(500);
                        if (test_parameter.i2c_enable) RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.addr, ispos ? buf_min : buf_max);
                        else RTDev.SwirePulse(ispos ? test_parameter.code_min : test_parameter.code_max);
                        System.Threading.Thread.Sleep(2000);
                        InsControl._scope.Root_STOP();
                        InsControl._scope.Measure_Clear();
                        InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        InsControl._scope.DoCommand(":MEASure:FALLtime CHANnel1");
                        InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement MEASurement1");
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_max");
                        imax = InsControl._scope.Meas_CH4MAX();
                        max = InsControl._scope.Meas_CH1MAX();
                        min = InsControl._scope.Meas_CH1MIN();
#if Report
                        _sheet.Cells[row, XLS_Table.J] = imax * 1000;
                        _sheet.Cells[row, XLS_Table.K] = max;
                        _sheet.Cells[row, XLS_Table.L] = min;
                        for (int i = 1; i < 13; i++) _sheet.Columns[i].AutoFit();
#endif
                        InsControl._scope.Root_Clear();
                        InsControl._power.AutoPowerOff();
                        InsControl._eload.CH1_Loading(0);
                        InsControl._eload.AllChannel_LoadOff();
                        System.Threading.Thread.Sleep(500);
                        row++; idx++;

                    } // iout loop
                } // power loop


            } // vin loop

            InsControl._scope.AutoTrigger();
            InsControl._scope.Root_RUN();

        Stop:
            stopWatch.Stop();
#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;

            MyLib.SaveExcelReport(test_parameter.wave_path, temp + "C_CodeInrush_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
            delegate_mess.Invoke();
        }
    }
}
