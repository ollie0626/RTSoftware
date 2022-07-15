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
    public class ATE_Loadtrans: TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        private void OSCInit()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_Off();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_On();

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH1_ACoupling();
            MyLib.WaveformCheck();

            InsControl._scope.Trigger_CH4();
            InsControl._scope.TriggerLevel_CH4(0.2);
            InsControl._scope.CH1_Level(1);
            InsControl._scope.CH4_Level(1);
            InsControl._scope.CH1_Offset(-2);
            InsControl._scope.CH4_Offset(3);
            MyLib.WaveformCheck();

            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
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

#if true
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "LoadTrans", bin_cnt, temp);
#endif

            OSCInit();
            MyLib.FuncGen_Fixedparameter(
                                        test_parameter.freq,
                                        test_parameter.duty,
                                        test_parameter.tr,
                                        test_parameter.tf);

            for (int freq_idx = 0; freq_idx < freq_cnt; freq_idx++)
            {
                if (freq_idx == 0 && test_parameter.Freq_en[0])
                    RTBBControl.Gpio_Enable();
                else
                    RTBBControl.Gpio_Disable();
                for (int vin_idx = 0; vin_idx < test_parameter.Vin_table.Count; vin_idx++)
                {
#if true
                    printTitle(row); row++;
#endif
                    InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[0]);
                    for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++)
                    {
                        double current_level, trigger_level;
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((func_idx % 20) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                        InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);

                        MyLib.FuncGen_loopparameter(
                            test_parameter.HiLo_table[func_idx].Highlevel,
                            test_parameter.HiLo_table[func_idx].LowLevel);


                        current_level = (test_parameter.HiLo_table[func_idx].Highlevel + test_parameter.HiLo_table[func_idx].LowLevel) / 4;
                        trigger_level = test_parameter.HiLo_table[func_idx].Highlevel * 0.6 + test_parameter.HiLo_table[func_idx].LowLevel * 0.4;
                        InsControl._scope.TriggerLevel_CH4(trigger_level);
                        InsControl._scope.CH4_Level(current_level);
                        InsControl._scope.CH4_Offset(current_level * 3);
                        InsControl._scope.SetTrigModeEdge(false);


                        InsControl._scope.Root_STOP();
                        InsControl._scope.NormalTrigger();
                        InsControl._scope.DoCommand(":MEASure: CLEar");
                        InsControl._scope.DoCommand(":MEASURE:VPP CHANnel1");
                        InsControl._scope.DoCommand(":MEASURE:VMAX CHANnel1");
                        InsControl._scope.DoCommand(":MEASURE:VMIN CHANnel1");
                        InsControl._scope.CH1_Level(0.3);
                        InsControl._scope.Root_RUN();
                        MyLib.WaveformCheck();
                        ChannelResize();








                    } // iout loop
                } // vin loop
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
            double max = InsControl._scope.Meas_CH1MAX();
            for(int i = 0; i < 3; i++)
            {
                InsControl._scope.CH1_Level(max / 3);
                max = InsControl._scope.Meas_CH1MAX();
                MyLib.ProcessCheck();
            }

            max = InsControl._scope.Meas_CH4MAX();
            for (int i = 0; i < 3; i++)
            {
                InsControl._scope.CH1_Level(max / 3);
                max = InsControl._scope.Meas_CH1MAX();
                MyLib.ProcessCheck();
            }
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
