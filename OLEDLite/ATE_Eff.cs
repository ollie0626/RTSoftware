using System;
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
    public class ATE_Eff : TaskRun
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
            Channel_num.Add("102"); // vo12
            Channel_num.Add("103"); // vo3
            Channel_num.Add("104"); // vo4

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
            for(int bin_idx = 0; 
                bin_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swire_cnt);
                bin_idx++)
            {
                for(int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
#if Report
                    _sheet.Cells[row, XLS_Table.A] = "VIN (V)";
                    _sheet.Cells[row, XLS_Table.B] = "Iin (mA)";
                    _sheet.Cells[row, XLS_Table.C] = "VO12 (V)";
                    _sheet.Cells[row, XLS_Table.D] = "VO3 (V)";
                    _sheet.Cells[row, XLS_Table.E] = "VO4 (V)";
                    _sheet.Cells[row, XLS_Table.F] = "IO12 (mA)";
                    _sheet.Cells[row, XLS_Table.G] = "IO3 (mA)";
                    _sheet.Cells[row, XLS_Table.H] = "IO4 (mA)";
                    _sheet.Cells[row, XLS_Table.I] = "Pin";
                    _sheet.Cells[row, XLS_Table.J] = "Po";
                    _sheet.Cells[row, XLS_Table.K] = "Eff(%)";
                    _range = _sheet.Range["A" + row, "K" + row];
                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    row++;
                    _sheet.Cells[row, XLS_Table.A] = string.Format("{0:0.00}", test_parameter.vinList[vin_idx]);
                    _sheet.Cells[row, XLS_Table.B] = "ESwire=" + test_parameter.ESwireList[bin_idx] + ", ASwire=" + test_parameter.ASwireList[bin_idx];
                    row++;
#endif
                    for(int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        //if (test_parameter.run_stop == true) goto Stop;
                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        System.Threading.Thread.Sleep(500);
                        MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);
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
                        double Iin = InsControl._dmm1.GetCurrent(3);
                        // Io12, Io3, Io4
                        double[] Iout = InsControl._eload.GetAllChannel_Iout();

                        double Pin = measure_data[0] * Iin;
                        _sheet.Cells[row, XLS_Table.A] = string.Format("{0:0.000}", measure_data[0]);
                        _sheet.Cells[row, XLS_Table.B] = string.Format("{0:0.000}", Iin * 1000);
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.ESwire_state ? string.Format("{0:0.000}", measure_data[1]) : "0";
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.ASwire_state ? string.Format("{0:0.000}", measure_data[2]) : "0";
                        _sheet.Cells[row, XLS_Table.E] = test_parameter.ENVO4_state ? string.Format("{0:0.000}", measure_data[3]) : "0";
                        _sheet.Cells[row, XLS_Table.F] = test_parameter.ESwire_state ? string.Format("{0:0.000}", Iout[0] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.G] = test_parameter.ASwire_state ? string.Format("{0:0.000}", Iout[1] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.H] = test_parameter.ENVO4_state ? string.Format("{0:0.000}", Iout[2] * 1000) : "0";
                        _sheet.Cells[row, XLS_Table.I] = "=ABS(A" + row + "*B" + row + ")"; // pin
                        _sheet.Cells[row, XLS_Table.J] = "=ABS(C" + row + "*F" + row +
                                                             "+D" + row + "*G" + row +
                                                             "+E" + row + "*H" + row + ")"; // pout
                        _sheet.Cells[row, XLS_Table.K] = "=(J" + row + "/J" + row + ")*100";
                    }
                }

            } // interface loop

        }
    }
}
