using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace OLEDLite
{
    public class ATE_Line : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        RTBBControl RTDev = new RTBBControl();

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            List<int> start_pos = new List<int>();
            List<int> stop_pos = new List<int>();
            List<string> Channel_num = new List<string>();

            Channel_num.Add("101"); // vin
            Channel_num.Add("102"); // vo1
            Channel_num.Add("103"); // vo2
            Channel_num.Add("104"); // vo3
            Channel_num.Add("105"); // vo4

            int row = 11;
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            int vin_cnt = test_parameter.vinList.Count;
            int iout_cnt = test_parameter.ioutList.Count;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);
            RTDev.BoadInit();
#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;

            _sheet.Cells[1, XLS_Table.A] = "Vin";
            _sheet.Cells[2, XLS_Table.A] = "Iout";
            _sheet.Cells[4, XLS_Table.A] = "Date";
            _sheet.Cells[5, XLS_Table.A] = "Note";
            _sheet.Cells[6, XLS_Table.A] = "Version";

            _sheet.Cells[1, XLS_Table.B] = test_parameter.vin_info;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.eload_info;
            _sheet.Cells[3, XLS_Table.B] = test_parameter.date_info;
            _sheet.Cells[6, XLS_Table.B] = test_parameter.ver_info;
#endif
            InsControl._power.AutoPowerOff();
            for (int bin_idx = 0;
                bin_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swire_cnt);
                bin_idx++)
            {

                for(int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                {

#if Report
                    _sheet.Cells[row, XLS_Table.A] = "VIN (V)";
                    _sheet.Cells[row, XLS_Table.B] = "Iin (mA)";
                    _sheet.Cells[row, XLS_Table.C] = "VO1 (V)";
                    _sheet.Cells[row, XLS_Table.D] = "VO2 (V)";
                    _sheet.Cells[row, XLS_Table.E] = "VO3 (V)";
                    _sheet.Cells[row, XLS_Table.F] = "VO4 (V)";
                    _sheet.Cells[row, XLS_Table.G] = "IO1 (mA)";
                    _sheet.Cells[row, XLS_Table.H] = "IO2 (mA)";
                    _sheet.Cells[row, XLS_Table.I] = "IO3 (mA)";
                    _sheet.Cells[row, XLS_Table.K] = "IO4 (mA)";
                    _range = _sheet.Range["A" + row, "K" + row];
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    row++;
#endif
                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
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

                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("Please connect DAQ !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }

                        if (binList[0] != "" && test_parameter.i2c_enable) RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                        else
                        {
                            // ic setting
                            int[] pulse_tmp;
                            bool[] Enable_state_table = new bool[] { test_parameter.ESwire_state, test_parameter.ASwire_state, test_parameter.ENVO4_state };
                            int[] Enable_num_table = new int[] { RTBBControl.ESwire, RTBBControl.ASwire, RTBBControl.ENVO4 };
                            pulse_tmp = test_parameter.ESwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(true, pulse_tmp[pulse_idx]);

                            pulse_tmp = test_parameter.ASwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(false, pulse_tmp[pulse_idx]);

                            for (int i = 0; i < Enable_state_table.Length; i++) RTBBControl.Swire_Control(Enable_num_table[i], Enable_state_table[i]);
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
                        row++;
                    }
                }

#if Report
                stop_pos.Add(row - 1);
                TimeSpan timeSpan = stopWatch.Elapsed;
                //AddAllCurve(start_pos, stop_pos, bin_idx, X_axis);
                MyLib.SaveExcelReport(test_parameter.wave_path, "Temp=" + temp + "_Efficiency @ ESwire=" + test_parameter.ESwireList[bin_idx]
                                      + "ASsire=" + test_parameter.ASwireList[bin_idx] + DateTime.Now.ToString("yyyyMMdd"), _book);
                _book.Close(false);
                _book = null;
                _app.Quit();
                _app = null;
                GC.Collect();
#endif
            }

        Stop:
            System.Windows.Forms.MessageBox.Show("Test finished!!!", "OLED Lite", System.Windows.Forms.MessageBoxButtons.OK);

        }
    }
}
