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
    public class ATE_OutputRipple : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        private void OSCInint()
        {
            InsControl._scope.CH1_ACoupling();
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_Off();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_On();


            InsControl._scope.CH1_Level(0.1);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH4_Level(0.3);
            InsControl._scope.CH4_Offset(0.3 * 3);

            InsControl._scope.Trigger_CH4();
            InsControl._scope.TimeScaleUs(20);
            InsControl._scope.TimeBasePositionUs(0);
            InsControl._scope.DoCommand(":MEASure:VPP CHANnel1");
        }


        public override void ATETask()
        {
            int freq_cnt = (test_parameter.Freq_en[0] ? 1 : 0) + (test_parameter.Freq_en[1] ? 1 : 0);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int row = 22;
            int bin_cnt = 1;
            MyLib Mylib = new MyLib();
            string[] binList = new string[1];
            binList = Mylib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            double[] vinList = new double[test_parameter.Vin_table.Count];
            Array.Copy(vinList, test_parameter.Vin_table.ToArray(), vinList.Length);

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Ripple", bin_cnt, temp);

#endif
            OSCInint();
            for (int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();
                
                for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                {
#if Report
                    printTitle(row); row++;
#endif
                    InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);
                    double target = vinList[vin_idx];
                    for (int iout_idx = 0; iout_idx < test_parameter.Iout_table.Count; iout_idx++)
                    {
                        double vin, iout, vpp, vout, iin;
                        string file_name = string.Format("{0}_Vin={1}_Iout={2}_Freq={3}",
                                                        row - 22,
                                                        test_parameter.Vin_table[vin_idx],
                                                        test_parameter.Iout_table[iout_idx],
                                                        test_parameter.Freq_des);
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((iout_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        MyLib.Switch_ELoadLevel(test_parameter.Iout_table[iout_idx]);

                        InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.Iout_table[iout_idx]);
                        MyLib.Delay1ms(150);
                        MyLib.Vincompensation(target, ref vinList[vin_idx]);
                        ChannelResize();
                        InsControl._scope.Root_STOP();

#if Report
                        vin = InsControl._34970A.Get_100Vol(1);
                        vout = InsControl._34970A.Get_100Vol(2);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();
                        vpp = InsControl._scope.Meas_CH1VPP();

                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = string.Format("{0:00.00}", vin);
                        _sheet.Cells[row, XLS_Table.D] = string.Format("{0:00.00}", iin);
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
                        _sheet.Cells[row, XLS_Table.F] = string.Format("{0:00.00}", vout);
                        _sheet.Cells[row, XLS_Table.G] = string.Format("{0:00.00}", iout);
                        _sheet.Cells[row, XLS_Table.H] = string.Format("{0:00.00}", vpp);

#endif
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);
                        InsControl._scope.Root_RUN();
                        row++;
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;


#if Report
            _sheet.Cells[2, 2] = str_temp;
            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_Eff" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        } // ATETask

        private void ChannelResize()
        {
            int time_cnt = 20;
            double max = 0, period, vpp = 0;
            double fail_num = 9.99 * Math.Pow(10, 9);
            for (int i = 0; i < 3; i++)
            {
                max = InsControl._scope.Meas_CH4MAX();
                InsControl._scope.CH4_Offset(max / 3);
                MyLib.Delay1ms(50);

                vpp = InsControl._scope.Meas_CH1VPP();
                InsControl._scope.CH1_Level(vpp / 3);

            }
            InsControl._scope.TriggerLevel_CH4(max * 0.75);
            period = InsControl._scope.Meas_CH4Period();
            while (period > fail_num)
            {
                time_cnt++;
                InsControl._scope.TimeScaleUs(time_cnt);
                period = InsControl._scope.Meas_CH4Period();
                if (time_cnt >= 100) break;
            }
            InsControl._scope.TimeScale(period / 2);

        }

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
            _sheet.Cells[row, XLS_Table.H] = "VPP(mV)";

            _range = _sheet.Range["A" + row, "E" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["F" + row, "H" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
        }
    }
}
