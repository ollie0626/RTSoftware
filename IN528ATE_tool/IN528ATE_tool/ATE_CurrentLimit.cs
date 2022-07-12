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

    /* 
     * test Item: Current Limit
     * Measure Method: ELoad CV Mode to check ILX Level
     * Waveform Imformation: Vout and LX, ILX
     * Veriable: Bin file, Vout, Iout and Chamber
     */

    public class ATE_CurrentLimit : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        MyLib MyLib;
        RTBBControl RTDev = new RTBBControl();

        private void Channel_Resize()
        {
            InsControl._scope.TimeScaleUs(50);
            InsControl._scope.Trigger(2);
            InsControl._scope.AutoTrigger();
            InsControl._scope.SetTrigModeEdge(false);
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH3_Off();
            InsControl._scope.CH4_On();

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();

            InsControl._scope.CH1_Level(3.5);
            InsControl._scope.CH2_Level(3.5);
            InsControl._scope.CH4_Level(1);

            InsControl._scope.CH4_Offset(1 * 3);
            InsControl._scope.CH1_Offset(3.5 * 2);
            InsControl._scope.CH2_Offset(3.5 * 2);
            MyLib.Delay1ms(200);

            double vout, ILx;
            // Channel1: Vout
            // Channel2: Lx
            // Channel4: ILx
            vout = InsControl._scope.Meas_CH1MAX();
            InsControl._scope.TriggerLevel_CH2(vout * 0.6);
            ILx = InsControl._scope.Meas_CH4AVG(); // ILX
            InsControl._scope.CH4_Level(ILx / 3);
            InsControl._scope.CH4_Offset(ILx);
            MyLib.Delay1ms(200);

            for (int i = 0; i < 3; i++)
            {
                InsControl._scope.CH1_Level(vout / 4);
                InsControl._scope.CH2_Level(vout / 3);
                vout = InsControl._scope.Meas_CH1MAX();
                MyLib.Delay1ms(200);
            }


            double period = InsControl._scope.Meas_CH2Period();
            InsControl._scope.TimeScale(period);
        }


        public void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            MyLib = new MyLib();

            int bin_cnt = 1;
            int row = 22;
            string[] binList = MyLib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            RTDev.BoadInit();
            int vin_cnt = test_parameter.VinList.Count;
            
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if false
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            MyLib.ExcelReportInit(_sheet);
            MyLib.testCondition(_sheet, "Current_Limit", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "CV(%)";
            _sheet.Cells[row, XLS_Table.E] = "CV(V)";
            _sheet.Cells[row, XLS_Table.F] = "Bin file";
            _sheet.Cells[row, XLS_Table.G] = "Vout(V)";
            _sheet.Cells[row, XLS_Table.I] = "ILX_Max(A)";
            _sheet.Cells[row, XLS_Table.K] = "Power Off ILX_Max(A)";

            _range = _sheet.Range["A" + row, "F" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(124, 252, 0);

            _range = _sheet.Range["G" + row, "K" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CV_Mode();
            InsControl._scope.Measure_Clear();
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {
                for(int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    string file_name;
                    string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                    file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V_CV={4:0.##}%",
                            row - 22, res, temp,
                            test_parameter.VinList[vin_idx],
                            test_parameter.cv_setting
                            );

                    if (test_parameter.run_stop == true) goto Stop;
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                    System.Threading.Thread.Sleep(500);
                    if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), test_parameter.addr, test_parameter.specify_bin);
                    if (binList[0] != "") RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), test_parameter.addr, binList[bin_idx]);
                    InsControl._scope.AutoTrigger();
                    // CV enable
                    double cv_vol = InsControl._eload.GetVol() * (test_parameter.cv_setting / 100);
                    InsControl._eload.SetCV_Vol(cv_vol);
                    double tempVin = ori_vinTable[vin_idx];
                    //if (!MyLib.Vincompensation(InsControl._power, InsControl._34970A, ori_vinTable[vin_idx], ref tempVin))
                    //{
                    //    System.Windows.Forms.MessageBox.Show("34970 沒有連結 !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                    //    return;
                    //}
                    // channel resize and time scale resize. use channel 1, 2, 4.
                    Channel_Resize();
                    MyLib.Delay1ms(250);

                    InsControl._scope.DoCommand(":MEASure:VMAX CHANnel4"); // ILX max OCP
                    InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2"); // LX level max
                    InsControl._scope.DoCommand(":MEASure:VAVerage DISPlay, CHANnel1"); // Vout Level

                    InsControl._scope.Root_STOP();
                    MyLib.Delay1ms(50);
                    double max_ch4 = InsControl._scope.Meas_CH4MAX(); // ILX
                    double max_ch2 = InsControl._scope.Meas_CH2MAX(); // LX
                    double amp_ch1 = InsControl._scope.Meas_CH1MAX(); // Vout
                    InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name);
                    InsControl._scope.Root_RUN();
#if false
                    _sheet.Cells[row, XLS_Table.A] = row - 22;
                    _sheet.Cells[row, XLS_Table.B] = temp;
                    _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                    _sheet.Cells[row, XLS_Table.F] = res;
                    _sheet.Cells[row, XLS_Table.G] = amp_ch1;
                    _sheet.Cells[row, XLS_Table.D] = test_parameter.cv_setting;
                    _sheet.Cells[row, XLS_Table.E] = cv_vol;
                    _sheet.Cells[row, XLS_Table.I] = max_ch4; // current limit
#endif

                    // power off test
                    InsControl._scope.Trigger(1);
                    InsControl._scope.TriggerLevel_CH1(amp_ch1 * 0.7);
                    InsControl._scope.SetTrigModeEdge(true);
                    InsControl._scope.NormalTrigger();
                    InsControl._scope.TimeScaleMs(40);
                    MyLib.Delay1s(3);
                    

                    InsControl._power.AutoPowerOff();
                    double offset = InsControl._scope.doQueryNumber(":CHAN4:OFFSet?");
                    InsControl._scope.CH4_Offset(offset);
                    MyLib.Delay1s(2);
                    InsControl._scope.Root_STOP();
                    max_ch4 = InsControl._scope.Meas_CH4MAX();
#if false
                    _sheet.Cells[row, XLS_Table.K] = max_ch4; // power off ILX maximum
#endif
                    InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_OFF");
                    InsControl._scope.Root_RUN();
                    InsControl._eload.AllChannel_LoadOff();
                    MyLib.Delay1ms(150);
                    row++;
                }
            }
        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
#if false
            for (int i = 1; i < 12; i++) _sheet.Columns[i].AutoFit();
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_CurrentLimit_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

    }
}
