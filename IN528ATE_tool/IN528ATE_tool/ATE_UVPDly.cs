using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;

namespace IN528ATE_tool
{
    public class ATE_UVPDly : TaskRun
    {

        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        MyLib MyLib = new MyLib();
        RTBBControl RTDev = new RTBBControl();

        private void OSCInint()
        {
            // CH1 Vout
            // CH2 LX
            // CH4 ILX
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH4_On();
            InsControl._scope.CH3_Off();

            InsControl._scope.CH1_Level(5);
            //InsControl._scope.CH2_Level(5);
            InsControl._scope.CH4_Level(1);
            // right position is negtive
            // up position is negtive 
            InsControl._scope.TimeScaleMs(test_parameter.cv_wait * 3); // trigger point
            InsControl._scope.TimeBasePositionMs(test_parameter.cv_wait * 3 * -3);
            //InsControl._scope.DoCommand(":FUNCtion1:VERTical AUTO");
            //InsControl._scope.DoCommand(string.Format(":FUNCTION1:ABSolute CHANNEL{0}", 1));
            //InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
            //InsControl._scope.DoCommand(":FUNCTION2:DISPLAY ON");
            InsControl._scope.Root_Clear();
            InsControl._scope.Root_RUN();
            InsControl._scope.Measure_Clear();
        }

        public void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            //int idx = 0;
            int vin_cnt = test_parameter.VinList.Count;
            int iout_cnt = test_parameter.IoutList.Count;
            int row = 22;
            string[] binList = new string[1];
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;

            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);
            binList = MyLib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            MyLib.ExcelReportInit(_sheet);
            MyLib.testCondition(_sheet, "UVP_Dly", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Bin file";
            _sheet.Cells[row, XLS_Table.E] = "UVP_DLY(ms)";
            _range = _sheet.Range["A" + row, "D" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["E" + row, "E" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            //InsControl._eload.CV_Mode();
            OSCInint();

            for(int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for(int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    if (test_parameter.run_stop == true) goto Stop;
                    if ((bin_idx % 5) == 0 && test_parameter.chamber_en) InsControl._chamber.GetChamberTemperature();
                    string file_name;
                    double ori_vol = 0;
                    string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                    file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V",
                            row - 22, res, temp,
                            test_parameter.VinList[vin_idx]
                            );

                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                    MyLib.Delay1ms(250);
                    if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                    MyLib.Delay1ms(100);
                    if (binList[0] != "") RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                    MyLib.Delay1ms(100);
                    ori_vol = InsControl._eload.GetVol();
                    InsControl._scope.Trigger_CH1();
                    InsControl._scope.CH1_Level(ori_vol / 5);
                    InsControl._scope.CH1_Offset((ori_vol / 5) * 3);
                    InsControl._scope.SetTrigModeEdge(true);
                    InsControl._scope.TriggerLevel_CH1(ori_vol * 0.65);
                    InsControl._scope.NormalTrigger();
                    InsControl._scope.Root_Clear();

                    // eload shot on to trigger uvp function
                    InsControl._eload.ShortOn();
                    MyLib.Delay1s(1);
                    InsControl._eload.ShortOff();

                    InsControl._scope.Root_STOP();
                    InsControl._scope.DoCommand(":MEASure:PPULses CHANNEL2");
                    InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                    InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement");
                    //:MARKer2:X:POSition?
                    double UVP_dly = InsControl._scope.doQueryNumber(":MARKer2:X:POSition?");
                    InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);

#if Report
                    _sheet.Cells[row, XLS_Table.A] = row - 22;
                    _sheet.Cells[row, XLS_Table.B] = temp;
                    _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                    _sheet.Cells[row, XLS_Table.D] = res;
                    _sheet.Cells[row, XLS_Table.J] = UVP_dly * 1000;
#endif
                    InsControl._power.AutoPowerOff();
                    InsControl._eload.AllChannel_LoadOff();
                    InsControl._scope.Root_RUN();
                    InsControl._scope.AutoTrigger();
                    InsControl._scope.Root_Clear();
                    row++;
                }
            }

        Stop:
            stopWatch.Stop();
#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;

            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_UVPDly_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

    }
}
