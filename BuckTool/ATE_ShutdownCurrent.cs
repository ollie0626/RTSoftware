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
    public class ATE_ShutdownCurrent : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public void printTitle(int row)
        {
            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iout(A)";
            _sheet.Cells[row, XLS_Table.E] = "EnOn Vout(V)";
            _sheet.Cells[row, XLS_Table.F] = "EnOn Iin(A)";
            _sheet.Cells[row, XLS_Table.G] = "EnOn Vout(V)";
            _sheet.Cells[row, XLS_Table.H] = "EnOff Iin(A)";

            _range = _sheet.Range["A" + row, "D" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["E" + row, "H" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
        }


        public override void ATETask()
        {
            int freq_cnt = (test_parameter.Freq_en[0] ? 1 : 0) + (test_parameter.Freq_en[1] ? 1 : 0);
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int row = 22;

            int no = 1;

            MyLib Mylib = new MyLib();
            double[] vinList = test_parameter.Vin_table.ToArray();

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Shutdown IQ", 0, temp);


#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CH1_Loading(0);
            InsControl._eload.CCL_Mode();


            for (int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();

                for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                {
                    printTitle(row);
                    for (int i = 1; i < 8; i++) _sheet.Columns[i].AutoFit();
                    row++;

                    for (int iout_idx = 0; iout_idx < test_parameter.Iout_table.Count; iout_idx++)
                    {
                        double iout = test_parameter.Iout_table[iout_idx];
                        double vin = test_parameter.Vin_table[vin_idx];
                        double temp = 25;
                        if (test_parameter.run_stop == true) goto Stop;
                        if (test_parameter.chamber_en == true) temp = InsControl._chamber.GetChamberTemperature();
                        
                        double on_current = 0, off_current = 0;
                        double on_vout = 0, off_vout = 0;
                        MyLib.Switch_ELoadLevel(iout);


                        // step 1. power on -> en -> eload on
                        InsControl._power.AutoSelPowerOn(vin);
                        RTBBControl.GpioEn_Enable();
                        InsControl._eload.CH1_Loading(iout);
                        MyLib.Delay1ms(test_parameter.en_ms * 1000);

                        // Measure Iin current (En on)
                        on_current = InsControl._dmm1.GetCurrent(0);
                        on_vout = InsControl._34970A.Get_10Vol(1);

                        // step 2. 
                        RTBBControl.GpioEn_Disable();
                        InsControl._eload.LoadOFF(1);
                        
                        // record ecah times shut down current
                        for (int test_idx = 0; test_idx < test_parameter.test_cnt; test_idx++)
                        {
                            // Measure Iin current (En off)
                            off_current = InsControl._dmm1.GetCurrent(0);
                            off_vout = InsControl._34970A.Get_10Vol(1);

                            // delay time
                            MyLib.Delay1ms(test_parameter.interval * 1000);

                            _sheet.Cells[row, XLS_Table.A] = no++;
                            _sheet.Cells[row, XLS_Table.B] = temp;
                            _sheet.Cells[row, XLS_Table.C] = vin;
                            _sheet.Cells[row, XLS_Table.D] = iout;
                            _sheet.Cells[row, XLS_Table.E] = on_current;
                            _sheet.Cells[row, XLS_Table.F] = on_vout;
                            _sheet.Cells[row, XLS_Table.G] = off_current;
                            _sheet.Cells[row, XLS_Table.H] = off_vout;
                            
                            row++;
                        }

                    } // eload loop
                } // vin loop
            } // freq loop

        Stop:
            InsControl._power.AutoPowerOff();
            InsControl._eload.CH1_Loading(0);

            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
#if Report
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_IQ_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        } // ATETask End



    }
}
