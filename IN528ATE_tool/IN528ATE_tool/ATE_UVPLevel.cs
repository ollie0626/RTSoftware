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
    public class ATE_UVPLevel : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        MyLib MyLib;
        RTBBControl RTDev = new RTBBControl();


        public void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            MyLib = new MyLib();
            int cv_cnt = (int)(test_parameter.cv_setting / test_parameter.cv_step);
            int bin_cnt = 1;
            int row = 22;
            string[] binList = MyLib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            RTDev.BoadInit();
            int vin_cnt = test_parameter.VinList.Count;

            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);
#if true
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            MyLib.ExcelReportInit(_sheet);
            MyLib.testCondition(_sheet, "Current_Limit", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Bin file";
            _sheet.Cells[row, XLS_Table.E] = "CV(%)";
            _sheet.Cells[row, XLS_Table.F] = "CV(V)";
            _sheet.Cells[row, XLS_Table.G] = "UVP(V)";
            _sheet.Cells[row, XLS_Table.H] = "UVP_Max(V)";
            _sheet.Cells[row, XLS_Table.I] = "UVP_Min(V)";
            _sheet.Cells[row, XLS_Table.J] = "UVP_DLY(ms)";

            _range = _sheet.Range["A" + row, "F" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["G" + row, "J" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);

            row++;
#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CV_Mode();

            OSCInint();

            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for(int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    double ori_vol = 0;
                    string file_name;
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

                    InsControl._scope.CH1_Level(ori_vol / 5);
                    InsControl._scope.CH1_Offset((ori_vol / 5) * 3);
                    InsControl._scope.SetTrigModeEdge(true);
                    InsControl._scope.TriggerLevel_CH1(ori_vol * 0.65);

                    for(int cv_idx = 0; cv_idx < cv_cnt; cv_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        double vol = 0;
                        double cv_vol = ori_vol * ((test_parameter.cv_setting - (test_parameter.cv_step * cv_idx)) / 100);
                        InsControl._eload.SetCV_Vol(cv_vol);
                        MyLib.Delay1ms(test_parameter.cv_wait);
                        vol = InsControl._eload.GetVol();
                        if(vol < (ori_vol * 0.5))
                        {
                            // Ic shoutdown
                            InsControl._scope.Root_STOP();
                            InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);
#if true
                            _sheet.Cells[row, XLS_Table.A] = row - 22;
                            _sheet.Cells[row, XLS_Table.B] = temp;
                            _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                            _sheet.Cells[row, XLS_Table.D] = res;
                            _sheet.Cells[row, XLS_Table.E] = test_parameter.cv_setting - (test_parameter.cv_step * cv_idx);
                            _sheet.Cells[row, XLS_Table.F] = cv_vol;
                            // check measure function
                            _sheet.Cells[row, XLS_Table.G] = "UVP(V)";
                            _sheet.Cells[row, XLS_Table.H] = "UVP_Max(V)";
                            _sheet.Cells[row, XLS_Table.I] = "UVP_Min(V)";
                            _sheet.Cells[row, XLS_Table.J] = "UVP_DLY(ms)";
#endif
                            break;
                        }
                        InsControl._power.AutoPowerOff();
                        InsControl._eload.AllChannel_LoadOff();
                        InsControl._scope.Root_RUN();
                        row++;
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);

            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_UVP_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
        }


        private void OSCInint()
        {
            // CH1 Vout
            // CH2 LX
            // CH3 ILX
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH4_On();
            InsControl._scope.CH3_Off();

            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH4_Level(1);

            InsControl._scope.TimeScaleMs(test_parameter.cv_wait * 3);
            InsControl._scope.TimeBasePositionMs(test_parameter.cv_wait * 3 * -3);
            InsControl._scope.DoCommand(":FUNCtion1:VERTical AUTO");
            InsControl._scope.DoCommand(string.Format(":FUNCTION1:ABSolute CHANNEL{0}", 1));
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
        }



    }
}
