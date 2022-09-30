using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using System.Diagnostics;

namespace IN528ATE_tool
{
    public class ATE_OutputRipple : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        Excel.Chart _chart;

        public double temp;
        MyLib myLib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_OutputRipple()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Output ripple test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            System.Threading.Thread.Sleep(2000);
            InsControl._scope.CH1_On();
            InsControl._scope.CH1_Level(0.3);
            InsControl._scope.TimeScaleMs(5);
            InsControl._scope.CH1_ACoupling();

            InsControl._scope.CH2_Off();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_Off();

            InsControl._scope.DoCommand(":MEASure:VPP CHANnel1");
            InsControl._scope.DoCommand(":MEASure:VMAX CHANnel1");
            InsControl._scope.DoCommand(":MEASure:VMIN CHANnel1");
            System.Threading.Thread.Sleep(1000);
        }


        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            int idx = 0;
            int vin_cnt = test_parameter.VinList.Count;
            int iout_cnt = test_parameter.IoutList.Count;
            int row = 22;
            string[] binList = new string[1];
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;

            // chart excel
            Excel.Range X_Range, Y_Range;
            Excel.SeriesCollection colletion;
            Excel.Series line;

            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);
            binList = myLib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Name = "outputRipple";
            myLib.ExcelReportInit(_sheet);
            myLib.testCondition(_sheet, "output ripple", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.F] = "Bin File";
            _range = _sheet.Range["A" + row.ToString(), "F" + row.ToString()];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _sheet.Cells[row, XLS_Table.G] = "Vout(V)";
            _sheet.Cells[row, XLS_Table.H] = "Vpp(mV)";
            _sheet.Cells[row, XLS_Table.I] = "Vmax(mV)";
            _sheet.Cells[row, XLS_Table.J] = "Vmin(mV)";
            _range = _sheet.Range["G" + row.ToString(), "J" + row.ToString()];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);

            _range = _sheet.Range["A" + row, "J" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["O22", "W38"];
            _chart = myLib.CreateChart(_sheet, _range, "Ripple", "index", "ripple(mV)");
            colletion = _chart.SeriesCollection();
            line = colletion.NewSeries();
            _chart.Legend.Delete();
#endif

            row++;
            InsControl._power.AutoPowerOff();
            OSCInit();
            InsControl._scope.DoCommand(":MEASHURE:VPP CHANnel1");
            System.Threading.Thread.Sleep(1000);
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for(int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for(int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        // the file name of waveform
                        string res = binList[bin_idx];
                        string file_name = string.Format("{0}_{1}_Temp={2}C_Vin={3:0.0#}V_{4:0.0#}A",
                                                        row - 22,
                                                        Path.GetFileNameWithoutExtension(res),
                                                        temp,
                                                        test_parameter.VinList[vin_idx],
                                                        test_parameter.IoutList[iout_idx]);
                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
#region "ripple test flow"
                        // power on voltage
                        InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                        System.Threading.Thread.Sleep(250);
                        // first write default code
                        if(test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                        //System.Threading.Thread.Sleep(150);
                        // write test conditiions
                        if (binList[0] != "") RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                        //System.Threading.Thread.Sleep(150);
                        // eload level switch
                        myLib.eLoadLevelSwich(InsControl._eload, test_parameter.IoutList[iout_idx]);
                        // eload sink current
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];
                        if (!myLib.Vincompensation(InsControl._power, InsControl._34970A, ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("34970 沒有連結 !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }

                        MyLib.Delay1ms(50);
                        // p2.0 enable
                        RTDev.GpEn_Enable();
                        MyLib.Delay1ms(50);
                        RTDev.Gp2En_Enable();

                        // adjust ch1 level
                        InsControl._scope.CH1_Level(1);
                        System.Threading.Thread.Sleep(500);
                        myLib.Channel_LevelSetting(InsControl._scope, 1);
                        System.Threading.Thread.Sleep(1000);
                        // scope open rgb color function
                        InsControl._scope.DoCommand(":DISPlay:PERSistence 5");
                        System.Threading.Thread.Sleep(5000);
                        double max, min, vpp, vin, vout, iin, iout;
                        // save waveform
                        InsControl._scope.Root_STOP();
                        // measure data
                        max = InsControl._scope.Meas_CH1MAX() * 1000;
                        min = InsControl._scope.Meas_CH1MIN() * 1000;
                        vpp = InsControl._scope.Meas_CH1VPP() * 1000;
                        vin = InsControl._34970A.Get_100Vol(1);
                        vout = InsControl._34970A.Get_100Vol(2);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);
#if Report
                        _sheet.Cells[row, XLS_Table.A] = idx;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = vin;
                        _sheet.Cells[row, XLS_Table.D] = iin * 1000;
                        _sheet.Cells[row, XLS_Table.E] = iout;
                        _sheet.Cells[row, XLS_Table.F] = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        _sheet.Cells[row, XLS_Table.G] = vout;
                        _sheet.Cells[row, XLS_Table.H] = vpp;
                        _sheet.Cells[row, XLS_Table.I] = max;
                        _sheet.Cells[row, XLS_Table.J] = min;
                        _range = _sheet.Range["A" + row, "J" + row];
                        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        for(int i = 1; i < 11; i++) _sheet.Columns[i].AutoFit();
                        InsControl._scope.Root_RUN();

                        X_Range = _sheet.Range["A23", "A" + row];
                        Y_Range = _sheet.Range["H23", "H" + row];
                        line.XValues = X_Range;
                        line.Values = Y_Range;
#endif

                        if(Math.Abs(vout) < 0.15)
                        {
                            RTDev.Gp2En_Disable();
                            MyLib.Delay1ms(50);
                            RTDev.GpEn_Disable();
                            MyLib.Delay1ms(50);
                            InsControl._power.AutoPowerOff();
                            InsControl._eload.CH1_Loading(0);
                            InsControl._eload.AllChannel_LoadOff();
                            System.Threading.Thread.Sleep(500);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            InsControl._scope.CH1_Level(1);
                            System.Threading.Thread.Sleep(250);
                        }

                        row++; idx++;
#endregion
                    } /* iout loop */
                } /* bin loop */
            } /* power loop */

            InsControl._scope.DoCommand(":DISPlay:PERSistence INFinite");
        Stop:
            stopWatch.Stop();
#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;


            myLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_Ripple_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
            if(!test_parameter.all_en && !test_parameter.chamber_en) delegate_mess.Invoke();
        }
    }
}
