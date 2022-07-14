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
            bool meter1_10A_en = false;
            bool meter2_10A_en = false;

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int row = 22;
            MyLib Mylib = new MyLib();
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = Mylib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            double[] vinList = new double[test_parameter.Vin_table.Count];
            Array.Copy(vinList, test_parameter.Vin_table.ToArray(), vinList.Length);

#if true
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Eff", bin_cnt, temp);
            printTitle(row); row++;
#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CH1_Loading(0);
            InsControl._eload.CCL_Mode();

            for(int freq_idx = 0; freq_idx < 2; freq_idx++)
            {
                InsControl._power.AutoPowerOff();
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();

                for(int iout_idx = 0; iout_idx < test_parameter.Iout_table.Count; iout_idx++)
                {
                    double Iin, level;
                    meter1_10A_en = false;
                    meter2_10A_en = false;
                    MyLib.Relay_Reset(false); // 10A level reset
#if true
                    printTitle(row);
#endif
                    level = test_parameter.Iout_table[iout_idx];
                    MyLib.Switch_ELoadLevel(level);
                    InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);
                    InsControl._eload.CH1_Loading(level);
                    Iin = InsControl._power.GetCurrent();
                    
                    if (!meter1_10A_en)
                        MyLib.Relay_Process(RTBBControl.GPIO2_0, Iin, true, ref meter1_10A_en);
                        
                    if (!meter2_10A_en)
                        MyLib.Relay_Process(RTBBControl.GPIO2_1, level, false, ref meter2_10A_en);

                    for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((iout_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                    } // vin loop
                } // iout loop
            } // freq loop

            Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;

#if true
            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

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

        }

        private void printTitle(int row)
        {
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iin(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Freq(MHz)";
            _sheet.Cells[row, XLS_Table.F] = "Vout(V)";
            _sheet.Cells[row, XLS_Table.G] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.H] = "LIR(%)";

            _range = _sheet.Range["A" + row, "E" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["F" + row, "I" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
        }
    }
}
