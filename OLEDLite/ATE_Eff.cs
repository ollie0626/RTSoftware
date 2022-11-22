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
            _sheet.Cells[3, XLS_Table.A] = "Swire";
            _sheet.Cells[4, XLS_Table.A] = "Date";
            _sheet.Cells[5, XLS_Table.A] = "Note";
#endif
            InsControl._power.AutoPowerOff();
            for(int bin_idx = 0; 
                bin_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count);
                bin_idx++)
            {
                for(int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
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

                        // vin, vo12, vo3, vo4
                        double[] measure_data = InsControl._34970A.QuickMEasureDefine(100, Channel_num);
                        double Iin = InsControl._dmm1.GetCurrent(3);
                        // Io12, Io3, Io4
                        double[] Iout = InsControl._eload.GetAllChannel_Iout();

                        double Pin = measure_data[0] * Iin;
                        double Pvo12 = measure_data[1] * Iout[0];
                        double Pvo3 = measure_data[2] * Iout[1];
                        double Pvo4 = measure_data[3] * Iout[2];






                    }
                }

            } // interface loop

        }
    }
}
