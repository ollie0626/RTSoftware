using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.Drawing;

namespace BuckTool
{
    public class ATE_Line: TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        override public void ATETask()
        {
            int freq_cnt = (test_parameter.Freq_en[0] ? 1 : 0) + (test_parameter.Freq_en[1] ? 1 : 0);
            bool meter1_10A_en = false;
            bool meter2_10A_en = false;
            bool sw400mA = true;
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int row = 22;
            MyLib Mylib = new MyLib();
            //int bin_cnt = 1;
            //string[] binList = new string[1];
            //binList = Mylib.ListBinFile(test_parameter.binFolder);
            //bin_cnt = binList.Length;
            double[] vinList = test_parameter.Vin_table.ToArray();
            //Array.Copy(vinList, test_parameter.Vin_table.ToArray(), vinList.Length);

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Line", 0, temp);
            //printTitle(row); row++;
#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CH1_Loading(0);
            InsControl._eload.CCL_Mode();

            for(int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                InsControl._power.AutoPowerOff();
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();

                for(int iout_idx = 0; iout_idx < test_parameter.Iout_table.Count; iout_idx++)
                {
                    double Iin, level;
                    
                    meter2_10A_en = false;
                    MyLib.Relay_Reset(true);
#if Report
                    printTitle(row); row++;
#endif
                    level = test_parameter.Iout_table[iout_idx];
                    MyLib.Switch_ELoadLevel(level);
                    if (!meter2_10A_en)
                        MyLib.Relay_Process(RTBBControl.GPIO2_1, level, iout_idx, 0, false, sw400mA, ref meter2_10A_en);

                    InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);
                    InsControl._eload.CH1_Loading(level);
                    Iin = InsControl._power.GetCurrent();

                    //if (!meter1_10A_en)
                    //    MyLib.Relay_Process(RTBBControl.GPIO2_0, Iin, 0, true, sw400mA, ref meter1_10A_en);



                    start_pos.Add(row);
                    for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                    {
                        double Vin, Vout, Iout;
                        meter1_10A_en = false;
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((iout_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        double target = test_parameter.Vin_table[vin_idx];
                        InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);

                        //Iin = InsControl._power.GetCurrent();
                        if (Iin > 0.3 && !meter1_10A_en)
                        {
                            InsControl._power.AutoPowerOff();
                            InsControl._eload.AllChannel_LoadOff();
                            InsControl._dmm1.ChangeCurrentLevel(false); // select 10A level
                            RTBBControl.Meter10A(RTBBControl.GPIO2_0);  // switch relay
                            InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);      // re-power on
                            InsControl._eload.CH1_Loading(level);       // turn on loading
                            meter1_10A_en = true;
                        }


                        MyLib.Vincompensation(target, ref vinList[vin_idx]);




                        MyLib.Delay1ms(250);
                        Vin = InsControl._34970A.Get_100Vol(1);
                        Vout = InsControl._34970A.Get_100Vol(2);
                        Iin = meter1_10A_en ? InsControl._dmm1.GetCurrent(3) : InsControl._dmm1.GetCurrent(1);
                        Iout = meter2_10A_en ? InsControl._dmm2.GetCurrent(3) : InsControl._dmm2.GetCurrent(1);

#if Report
                        _sheet.Cells[row, XLS_Table.A] = vin_idx;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = Vin;
                        _sheet.Cells[row, XLS_Table.D] = Iin;
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
                        _sheet.Cells[row, XLS_Table.F] = Vout;
                        _sheet.Cells[row, XLS_Table.G] = Iout;
                        _sheet.Cells[row, XLS_Table.H] = Math.Abs((Vout - test_parameter.vout_ideal) / test_parameter.vout_ideal) * 100;
                        _sheet.Cells[row, XLS_Table.I] = InsControl._34970A.Get_1Vol(3);    // for FB voltage
#endif
                        row++;
                    } // vin loop
                    stop_pos.Add(row - 1);
                } // iout loop
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

            AddCruve(start_pos, stop_pos);
            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_Line" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif

        } // ATETask

        private void AddCruve(List<int> start_pos, List<int> stop_pos)
        {
#if Report
            Excel.Chart chart;
            Excel.Range range;
            Excel.SeriesCollection collection;
            Excel.Series series;
            Excel.Range XRange, YRange;
            range = _sheet.Range["M16", "V32"];
            chart = MyLib.CreateChart(_sheet, range, "Line Regulation", "Vin (V)", "Vout (V)");
            // for LOR
            //range = _sheet.Range["M38", "V54"];

            chart.Legend.Delete();
            collection = chart.SeriesCollection();

            for(int line = 0; line < start_pos.Count; line++)
            {
                series = collection.NewSeries();
                XRange = _sheet.Range["C" + start_pos[line].ToString(), "C" + stop_pos[line].ToString()];
                YRange = _sheet.Range["H" + start_pos[line].ToString(), "H" + stop_pos[line].ToString()];
                series.XValues = XRange;
                series.Values = YRange;
                series.Name = "line" + (line + 1).ToString();
            }
#endif
        }

        private void printTitle(int row)
        {
#if Report
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Freq(MHz)";
            _sheet.Cells[row, XLS_Table.F] = "Vout(V)";
            _sheet.Cells[row, XLS_Table.G] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.H] = "LIR(%)";
            _sheet.Cells[row, XLS_Table.I] = "FB accuracy(V)";

            _range = _sheet.Range["A" + row, "E" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["F" + row, "I" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
#endif
        }
    }
}
