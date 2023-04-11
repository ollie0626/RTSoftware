using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;
using System.Windows.Forms;

namespace OLEDLite
{
    public struct CurveInfo
    {
        public string title;
        public string Xtitle;
        public string Ytitle;
        public string X;
        public string Y;
    };


    public class ATE_Eff : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        //public double temp;
        RTBBControl RTDev = new RTBBControl();

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();
            List<string> Channel_num = new List<string>();

            Channel_num.Add("101"); // VIN
            Channel_num.Add("102"); // ELVDD
            Channel_num.Add("103"); // ELVSS
            Channel_num.Add("104"); // AVDD
            Channel_num.Add("105"); // DVDD
                                    // -----------------------------------
                                    // Eload Channel 1 -> ELVDD + ELVSS
                                    // Eload Channel 2 -> AVDD
                                    // Eload Channel 3 -> DVDD


            int row = 11;
            int bin_cnt = 1;
            string X_axis = "";
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            int vin_cnt = test_parameter.vinList.Count;
            int iout_cnt = test_parameter.ioutList.Count;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);
            RTDev.BoadInit();

            InsControl._power.AutoPowerOff();
            for (int bin_idx = 0;
                bin_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swire_cnt);
                bin_idx++)
            {
                row = 11;
                _app = new Excel.Application();
                _app.Visible = true;
                _book = (Excel.Workbook)_app.Workbooks.Add();
                _sheet = (Excel.Worksheet)_book.ActiveSheet;

                for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
#if Report
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

                    
                    switch (test_parameter.eload_ch_select)
                    {
                        case 0:
                            X_axis = "G";
                            break;
                        case 1:
                            X_axis = "H";
                            break;
                        case 2:
                            X_axis = "I";
                            break;
                    }

                    _sheet.Cells[row, XLS_Table.A] = "VIN (V)";
                    _sheet.Cells[row, XLS_Table.B] = "Iin (mA)";
                    _sheet.Cells[row, XLS_Table.C] = "ELVDD (V)";
                    _sheet.Cells[row, XLS_Table.D] = "ELVSS (V)";
                    _sheet.Cells[row, XLS_Table.E] = "AVDD (V)";
                    _sheet.Cells[row, XLS_Table.F] = "DVDD (V)";
                    _sheet.Cells[row, XLS_Table.G] = "IO12 (mA)";
                    _sheet.Cells[row, XLS_Table.H] = "IO3 (mA)";
                    _sheet.Cells[row, XLS_Table.I] = "IO4 (mA)";
                    _sheet.Cells[row, XLS_Table.J] = "Pin";
                    _sheet.Cells[row, XLS_Table.K] = "Po";
                    _sheet.Cells[row, XLS_Table.L] = "Eff(%)";
                    _range = _sheet.Range["A" + row, "L" + row];
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    row++;
                    _sheet.Cells[row, XLS_Table.A] = string.Format("VIN={0:0.00}V", test_parameter.vinList[vin_idx]);
                    _sheet.Cells[row, XLS_Table.B] = "ESwire=" + test_parameter.ESwireList[bin_idx] + ", ASwire=" + test_parameter.ASwireList[bin_idx];
                    row++;
#endif
                    start_pos.Add(row);
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        System.Threading.Thread.Sleep(500);

                        switch (test_parameter.eload_ch_select)
                        {
                            case 0:
                                InsControl._eload.DoCommand(InsControl._eload.CH1);
                                break;
                            case 1:
                                InsControl._eload.DoCommand(InsControl._eload.CH2);
                                break;
                            case 2:
                                InsControl._eload.DoCommand(InsControl._eload.CH3);
                                break;
                        }
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);

                        switch (test_parameter.eload_ch_select)
                        {
                            case 0:
                                InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);
                                break;
                            case 1:
                                InsControl._eload.CH2_Loading(test_parameter.ioutList[iout_idx]);
                                break;
                            case 2:
                                InsControl._eload.CH3_Loading(test_parameter.ioutList[iout_idx]);
                                break;
                        }
                        MyLib.EloadFixChannel();

                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("Please connect DAQ !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }

                        // ic setting
                        if(iout_idx == 0)
                        {
                            System.Threading.Thread.Sleep(1000);
                            for (int k = 0; k < 3; k++)
                            {
                                int[] pulse_tmp;
                                bool[] Enable_state_table = new bool[] { test_parameter.ESwire_state, test_parameter.ASwire_state, test_parameter.ENVO4_state };
                                int[] Enable_num_table = new int[] { RTBBControl.ESwire, RTBBControl.ASwire, RTBBControl.ENVO4 };
                                pulse_tmp = test_parameter.ESwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                                for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(true, pulse_tmp[pulse_idx]);

                                pulse_tmp = test_parameter.ASwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                                for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(false, pulse_tmp[pulse_idx]);
                                for (int i = 0; i < Enable_state_table.Length; i++) RTBBControl.Swire_Control(Enable_num_table[i], Enable_state_table[i]);
                                MyLib.Delay1ms(50);
                            }

                        }

                        // vin, vo12, vo3, vo4
                        double[] measure_data = InsControl._34970A.QuickMEasureDefine(100, Channel_num);
                        double Iin = 0;

                        switch (test_parameter.eload_iin_select)
                        {
                            case 0: // 10A level
                                Iin = InsControl._dmm1.GetCurrent(3);
                                break;
                            case 1: // power supply 
                                Iin = InsControl._power.GetCurrent();
                                break;
                            case 2: // 400mA level
                                Iin = InsControl._dmm1.GetCurrent(1);
                                break;
                            case 3: // DMM6500 10A
                                Iin = InsControl._dmm1.GetCurrent(4);
                                break;
                            case 4: // DMM6500 400mA
                                Iin = InsControl._dmm1.GetCurrent(5);
                                break;
                        }

                        // Io12, Io3, Io4
                        double[] Iout = InsControl._eload.GetAllChannel_Iout();

                        _sheet.Cells[row, XLS_Table.A] = string.Format("{0:0.000}", measure_data[0]);
                        _sheet.Cells[row, XLS_Table.B] = string.Format("{0:0.000}", Iin * 1000);
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.ESwire_state ? string.Format("{0:0.000}", measure_data[1]) : "0";
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.ESwire_state ? string.Format("{0:0.000}", measure_data[2]) : "0";
                        _sheet.Cells[row, XLS_Table.E] = test_parameter.ASwire_state ? string.Format("{0:0.000}", measure_data[3]) : "0";
                        _sheet.Cells[row, XLS_Table.F] = test_parameter.ENVO4_state ? string.Format("{0:0.000}", measure_data[4]) : "0";

                        _sheet.Cells[row, XLS_Table.G] = test_parameter.ESwire_state ? string.Format("{0:0.000}", Iout[0] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.H] = test_parameter.ASwire_state ? string.Format("{0:0.000}", Iout[1] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.I] = test_parameter.ENVO4_state ? string.Format("{0:0.000}", Iout[2] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.J] = "=ABS(A" + row + "*B" + row + ")"; // pin
                        _sheet.Cells[row, XLS_Table.K] = "=(C" + row + "+ABS(D" + row + "))*G" + row
                                                       + "+E" + row + "*H" + row
                                                       + "+F" + row + "*I" + row;
                        _sheet.Cells[row, XLS_Table.L] = "=(K" + row + "/J" + row + ")*100";
                        row++;
                    } // eload loop

                    InsControl._power.AutoPowerOff();

                    InsControl._eload.AllChannel_LoadOff();
                    InsControl._eload.CH1_ClearSetting();
                    InsControl._eload.CH2_ClearSetting();
                    InsControl._eload.CH3_ClearSetting();
                    
                    stop_pos.Add(row - 1);
                } // power loop
#if Report

                stopWatch.Stop();
                TimeSpan timeSpan = stopWatch.Elapsed;
                string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
                _sheet.Cells[7, XLS_Table.B] = time;
                AddAllCurve(start_pos, stop_pos, bin_idx, X_axis);
                LORData(start_pos, stop_pos);
                MyLib.SaveExcelReport(test_parameter.wave_path, "Temp=" + temp + "_Efficiency @ ESwire=" + test_parameter.ESwireList[bin_idx]
                                      + "ASsire=" + test_parameter.ASwireList[bin_idx] + "_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
                _book.Close(false);
                _book = null;
                _app.Quit();
                _app = null;
                GC.Collect();
                start_pos.Clear();
                stop_pos.Clear();
#endif
            } // interface loop
        Stop:
            if (!test_parameter.chamber_en)
                System.Windows.Forms.MessageBox.Show("Test finished!!!", "OLED Lite", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void AddAllCurve(List<int> start_pos, List<int> stop_pos, int bin_idx, string X_axis)
        {
            Excel.Range range_pos;
            CurveInfo info;
            info.title = "Efficiency @ESwire=" + test_parameter.ESwireList[bin_idx] + "ASsire=" + test_parameter.ASwireList[bin_idx];
            info.Xtitle = "Load (mA)";
            info.Ytitle = "Eff (%)";
            info.X = X_axis;
            info.Y = "L";

            range_pos = _sheet.Range["D2", "J15"];
            AddCurve(
                _sheet, range_pos, info, start_pos, stop_pos
                );

            info.title = "LOR ELVDD @ ESwire=" + test_parameter.ESwireList[bin_idx] + "ASsire=" + test_parameter.ASwireList[bin_idx];
            info.Ytitle = "ELVDD(V)";
            info.Y = "C";
            range_pos = _sheet.Range["D18", "J31"];
            AddCurve(
                _sheet, range_pos, info, start_pos, stop_pos
                );


            info.title = "LOR ELVSS @ ESwire=" + test_parameter.ESwireList[bin_idx] + "ASsire=" + test_parameter.ASwireList[bin_idx];
            info.Ytitle = "ELVSS(V)";
            info.Y = "D";
            range_pos = _sheet.Range["D34", "J47"];
            AddCurve(
            _sheet, range_pos, info, start_pos, stop_pos
            );


            info.title = "LOR AVDD @ ESwire=" + test_parameter.ESwireList[bin_idx] + "ASsire=" + test_parameter.ASwireList[bin_idx];
            info.Ytitle = "AVDD(V)";
            info.Y = "E";
            range_pos = _sheet.Range["D50", "J63"];
            AddCurve(
            _sheet, range_pos, info, start_pos, stop_pos
            );

            info.title = "LOR DVDD @ ESwire=" + test_parameter.ESwireList[bin_idx] + "ASsire=" + test_parameter.ASwireList[bin_idx];
            info.Ytitle = "DVDD(V)";
            info.Y = "F";
            range_pos = _sheet.Range["D66", "J79"];
            AddCurve(
            _sheet, range_pos, info, start_pos, stop_pos
            );
        }

        private void AddCurve(
                              Excel.Worksheet sheet,
                              Excel.Range range_pos,
                              CurveInfo info,
                              List<int> start_pos,
                              List<int> stop_pos)
        {
            Excel.Chart chart;
            Excel.SeriesCollection collection;
            Excel.Series series;
            Excel.Range XRange, YRange;

            chart = MyLib.CreateChart(sheet, range_pos, info.title, info.Xtitle, info.Ytitle, true);
            chart.ChartTitle.Font.Size = 14;
            collection = chart.SeriesCollection();

            for (int line = 0; line < start_pos.Count; line++)
            {
                series = collection.NewSeries();
                XRange = sheet.Range[info.X + start_pos[line], info.X + stop_pos[line]];
                YRange = sheet.Range[info.Y + start_pos[line], info.Y + stop_pos[line]];

                series.Name = "VIN=" + sheet.Range["A" + (start_pos[line] - 1)].Value.ToString();
                series.XValues = XRange;
                series.Values = YRange;
            }
        }

        private void LORData(List<int> start_pos, List<int> stop_pos)
        {
            int row = 12;
            int interval = test_parameter.vinList.Count;
            // ELVDD
            _sheet.Cells[row + interval * 0, XLS_Table.O] = "ELVDD";
            _sheet.Cells[row + interval * 0, XLS_Table.P] = "Max (V)";
            _sheet.Cells[row + interval * 0, XLS_Table.Q] = "Min (V)";
            _sheet.Cells[row + interval * 0, XLS_Table.R] = "Offset (mV)";
            _sheet.Cells[row + interval * 0, XLS_Table.S] = "PLOR (mV)";
            _sheet.Cells[row + interval * 0, XLS_Table.T] = "NLOR (mV)";
            _sheet.Cells[row + interval * 0, XLS_Table.U] = "MaxIO (mA)";
            row++;
            // ELVSS
            _sheet.Cells[row + interval * 1, XLS_Table.O] = "ELVSS";
            _sheet.Cells[row + interval * 1, XLS_Table.P] = "Max (V)";
            _sheet.Cells[row + interval * 1, XLS_Table.Q] = "Min (V)";
            _sheet.Cells[row + interval * 1, XLS_Table.R] = "Offset (mV)";
            _sheet.Cells[row + interval * 1, XLS_Table.S] = "PLOR (mV)";
            _sheet.Cells[row + interval * 1, XLS_Table.T] = "NLOR (mV)";
            _sheet.Cells[row + interval * 1, XLS_Table.U] = "MaxIO (mA)";
            row++;
            // AVDD
            _sheet.Cells[row + interval * 2, XLS_Table.O] = "AVDD";
            _sheet.Cells[row + interval * 2, XLS_Table.P] = "Max (V)";
            _sheet.Cells[row + interval * 2, XLS_Table.Q] = "Min (V)";
            _sheet.Cells[row + interval * 2, XLS_Table.R] = "Offset (mV)";
            _sheet.Cells[row + interval * 2, XLS_Table.S] = "PLOR (mV)";
            _sheet.Cells[row + interval * 2, XLS_Table.T] = "NLOR (mV)";
            _sheet.Cells[row + interval * 2, XLS_Table.U] = "MaxIO (mA)";
            row++;
            // DVDD
            _sheet.Cells[row + interval * 3, XLS_Table.O] = "DVDD";
            _sheet.Cells[row + interval * 3, XLS_Table.P] = "Max (V)";
            _sheet.Cells[row + interval * 3, XLS_Table.Q] = "Min (V)";
            _sheet.Cells[row + interval * 3, XLS_Table.R] = "Offset (mV)";
            _sheet.Cells[row + interval * 3, XLS_Table.S] = "PLOR (mV)";
            _sheet.Cells[row + interval * 3, XLS_Table.T] = "NLOR (mV)";
            _sheet.Cells[row + interval * 3, XLS_Table.U] = "MaxIO (mA)";
            row++;
            row = 12;
            for (int line = 0; line < test_parameter.vinList.Count; line++)
            {
                // ELVDD
                _sheet.Cells[row + 1, XLS_Table.O] = _sheet.Range["A" + start_pos[line]].Value;
                _sheet.Cells[row + 1, XLS_Table.P] = "=MAX(C" + start_pos[line] + ":C" + stop_pos[line] + ")";
                _sheet.Cells[row + 1, XLS_Table.Q] = "=MIN(C" + start_pos[line] + ":C" + stop_pos[line] + ")";
                _sheet.Cells[row + 1, XLS_Table.R] = "=P" + (row + 1) + "-Q" + (row + 1);
                _sheet.Cells[row + 1, XLS_Table.S] = "=(P" + (row + 1) + "-C" + start_pos[line] + ")*1000";
                _sheet.Cells[row + 1, XLS_Table.T] = "=(Q" + (row + 1) + "-C" + start_pos[line] + ")*1000";
                _sheet.Cells[row + 1, XLS_Table.U] = "=MAX(G" + start_pos[line] + ":G" + stop_pos[line] + ")";
                // ELVSS
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.O] = _sheet.Range["A" + start_pos[line]].Value;
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.P] = "=MAX(D" + start_pos[line] + ":D" + stop_pos[line] + ")";
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.Q] = "=MIN(D" + start_pos[line] + ":D" + stop_pos[line] + ")";
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.R] = "=P" + (row + interval * 1 + 2) + "-Q" + (row + interval * 1 + 2);
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.S] = "=(P" + (row + interval * 1 + 2) + "-D" + start_pos[line] + ")*1000";
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.T] = "=(Q" + (row + interval * 1 + 2) + "-D" + start_pos[line] + ")*1000";
                _sheet.Cells[row + interval * 1 + 2, XLS_Table.U] = "=MAX(G" + start_pos[line] + ":G" + stop_pos[line] + ")";
                // AVDD
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.O] = _sheet.Range["A" + start_pos[line]].Value;
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.P] = "=MAX(E" + start_pos[line] + ":E" + stop_pos[line] + ")";
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.Q] = "=MIN(E" + start_pos[line] + ":E" + stop_pos[line] + ")";
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.R] = "=P" + (row + interval * 2 + 3) + "-Q" + (row + interval * 2 + 3);
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.S] = "=(P" + (row + interval * 2 + 3) + "-E" + start_pos[line] + ")*1000";
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.T] = "=(Q" + (row + interval * 2 + 3) + "-E" + start_pos[line] + ")*1000";
                _sheet.Cells[row + (interval * 2) + 3, XLS_Table.U] = "=MAX(H" + start_pos[line] + ":H" + stop_pos[line] + ")";
                // DVDD
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.O] = _sheet.Range["A" + start_pos[line]].Value;
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.P] = "=MAX(F" + start_pos[line] + ":F" + stop_pos[line] + ")";
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.Q] = "=MIN(F" + start_pos[line] + ":F" + stop_pos[line] + ")";
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.R] = "=P" + (row + interval * 3 + 4) + "-Q" + (row + interval * 3 + 4);
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.S] = "=(P" + (row + interval * 3 + 4) + "-F" + start_pos[line] + ")*1000";
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.T] = "=(Q" + (row + interval * 3 + 4) + "-F" + start_pos[line] + ")*1000";
                _sheet.Cells[row + (interval * 3) + 4, XLS_Table.U] = "=MAX(I" + start_pos[line] + ":I" + stop_pos[line] + ")";

                row++;
            }
        }
    }
}
