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
    public class ATE_UVPLevel : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        MyLib MyLib;
        RTBBControl RTDev = new RTBBControl();

        private void OSCInint()
        {
            // CH1 Vout
            // CH2 LX
            // CH4 ILX
            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH4_On();
            InsControl._scope.CH3_Off();

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();
            InsControl._scope.CH4_BWLimitOn();

            InsControl._scope.CH1_Level(5);
            //InsControl._scope.CH2_Level(5);
            InsControl._scope.CH4_Level(1);
            // right position is negtive
            // up position is negtive 
            InsControl._scope.TimeScaleMs(test_parameter.cv_wait * 3); // trigger point
            InsControl._scope.TimeBasePositionMs(test_parameter.cv_wait * 3 * -3);
            InsControl._scope.DoCommand(":FUNCtion1:VERTical AUTO");
            InsControl._scope.DoCommand(string.Format(":FUNCTION1:ABSolute CHANNEL{0}", 1));
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
            InsControl._scope.DoCommand(":FUNCTION2:DISPLAY ON");
            InsControl._scope.Root_Clear();
            InsControl._scope.Root_RUN();
            InsControl._scope.Measure_Clear();
        }


        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            MyLib = new MyLib();
            int cv_cnt = (int)(test_parameter.cv_setting / test_parameter.cv_step);
            int bin_cnt = 1;
            int row = 11;
            string[] binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            RTDev.BoadInit();
            int vin_cnt = test_parameter.vinList.Count;

            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);
#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            //MyLib.ExcelReportInit(_sheet);
            //MyLib.testCondition(_sheet, "UVP", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Bin file";
            _sheet.Cells[row, XLS_Table.E] = "CV(%)";
            _sheet.Cells[row, XLS_Table.F] = "CV(V)";
            _sheet.Cells[row, XLS_Table.G] = "UVP(V)";
            _sheet.Cells[row, XLS_Table.H] = "UVP_Max(V)";
            _sheet.Cells[row, XLS_Table.I] = "UVP_Min(V)";
            _sheet.Cells[row, XLS_Table.J] = "Vout(V)";
            //_sheet.Cells[row, XLS_Table.J] = "UVP_DLY(ms)";

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
                for (int bin_idx = 0; bin_idx < test_parameter.swire_cnt; bin_idx++)
                {
                    if (test_parameter.run_stop == true) goto Stop;
                    if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                    double ori_vol = 0;
                    string file_name = "";
                    string res = "";

                    if (!test_parameter.i2c_enable)
                    {
                        file_name = string.Format("{0}_Temp={1}_vin={2:0.##}V_CV={3:0.##}%_Pulse={4}",
                                    row - 22, temp,
                                    test_parameter.vinList[vin_idx],
                                    test_parameter.cv_setting,
                                    "ESwire=" + test_parameter.ESwireList[bin_idx] + ", ASwire=" + test_parameter.ASwireList[bin_idx]
                                    );
                    }
                    else
                    {
                        res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V",
                                    row - 22, res, temp,
                                    test_parameter.vinList[vin_idx]
                                    );
                    }


                    //res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                    //file_name = string.Format("{0}_{1}_Temp={2}C_vin={3:0.##}V",
                    //        row - 22, res, temp,
                    //        test_parameter.VinList[vin_idx]
                    //        );

                    InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                    MyLib.Delay1ms(250);

                    if (!test_parameter.i2c_enable)
                    {
                        int[] pulse_tmp;
                        bool[] Enable_state_table = new bool[] { test_parameter.ESwire_state, test_parameter.ASwire_state, test_parameter.ENVO4_state };
                        int[] Enable_num_table = new int[] { RTBBControl.ESwire, RTBBControl.ASwire, RTBBControl.ENVO4 };
                        pulse_tmp = test_parameter.ESwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                        for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(true, pulse_tmp[pulse_idx]);

                        pulse_tmp = test_parameter.ASwireList[bin_idx].Split(',').Select(int.Parse).ToArray();
                        for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTBBControl.SwirePulse(false, pulse_tmp[pulse_idx]);
                        for (int i = 0; i < Enable_state_table.Length; i++) RTBBControl.Swire_Control(Enable_num_table[i], Enable_state_table[i]);
                    }
                    else
                    {
                        //if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                        MyLib.Delay1ms(100);
                        if (binList[0] != "") RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, binList[bin_idx]);
                    }

                    MyLib.Delay1ms(100);
                    ori_vol = InsControl._eload.GetVol();

                    InsControl._scope.Trigger_CH1();
                    InsControl._scope.CH1_Level(ori_vol / 5);
                    InsControl._scope.CH1_Offset((ori_vol / 5) * 3);
                    InsControl._scope.SetTrigModeEdge(true);
                    InsControl._scope.TriggerLevel_CH1(ori_vol * 0.65);

                    double cv_vol = 0, cv_percent = 0;
                    InsControl._scope.NormalTrigger();
                    InsControl._scope.Root_Clear();
                    MyLib.Delay1s(1);
                    for (int cv_idx = 0; cv_idx < cv_cnt; cv_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        double vol = 0;
                        cv_vol = ori_vol * ((test_parameter.cv_setting - (test_parameter.cv_step * cv_idx)) / 100);
                        InsControl._eload.SetCV_Vol(cv_vol);
                        MyLib.Delay1ms((int)test_parameter.cv_wait * 15);
                        vol = InsControl._eload.GetVol();
                        cv_percent = test_parameter.cv_setting - (test_parameter.cv_step * cv_idx);
                        if (vol < (ori_vol * 0.5))
                        {
                            break;
                        }
                    }

                    MyLib.WaveformCheck();
                    // Ic shoutdown
                    InsControl._scope.Root_STOP();
                    InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);
                    /*
                        --------
                               |
                               --------
                                      |( wait )
                                      ---------
                                        ↓    ↓ | 
                                        ↓    ↓
                               wait * -0.7   wait * -0.3

                        Channel1: Vout
                        Channel2: Lx
                        Channel4: ILx
                        
                        Function1: Vout abs
                        Function2: Function1 gatting

                        use Lx get UVP delay time
                     */
                    double start_t = (test_parameter.cv_wait / 1000) * -1 * 0.7;
                    double stop_t = (test_parameter.cv_wait / 1000) * -1 * 0.3;
                    //InsControl._scope.DoCommand(string.Format(":FUNCtion1:GATing CHANnel1, {0}, {1}", start_t, stop_t));
                    InsControl._scope.DoCommand(string.Format(":FUNCtion2:GATing FUNCtion1, {0}, {1}", start_t, stop_t));


                    //InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
                    //InsControl._scope.DoCommand(string.Format(":FUNCtion1:GATing:STARt {0}", start_t));
                    //InsControl._scope.DoCommand(string.Format(":FUNCtion1:GATing:STOP {0}", stop_t));

                    InsControl._scope.DoCommand(":MEASure:SOURce FUNCtion2");
                    InsControl._scope.DoCommand(":MEASure:VMAX FUNCtion2");
                    InsControl._scope.DoCommand(":MEASure:VMIN FUNCtion2");
                    double UVP_amp = InsControl._scope.doQueryNumber(":MEASure:VAMPlitude? CHAN1");
                    double UVP_max = InsControl._scope.doQueryNumber(":MEASure:VMAX? CHAN1");
                    double UVP_min = InsControl._scope.doQueryNumber(":MEASure:VMIN? CHAN1");


                    //:MEASure:PPULses CHANNEL2
                    //:MARKer:MODE MEASurement
                    //:MARKer:MEASurement:MEASurement MEASurement1 --> ??
                    //InsControl._scope.DoCommand(":MEASure:PPULses CHANNEL2");
                    //InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                    //InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement");
                    ////:MARKer2:X:POSition?
                    //double UVP_dly = InsControl._scope.doQueryNumber(":MARKer2:X:POSition?");
#if Report
                    _sheet.Cells[row, XLS_Table.A] = row - 11;
                    _sheet.Cells[row, XLS_Table.B] = temp;
                    _sheet.Cells[row, XLS_Table.C] = test_parameter.vinList[vin_idx];
                    _sheet.Cells[row, XLS_Table.D] = "ESwire=" + test_parameter.ESwireList[bin_idx] + ", ASwire=" + test_parameter.ASwireList[bin_idx];
                    //_sheet.Cells[row, XLS_Table.E] = string.Format("{0}%", cv_percent);
                    _sheet.Cells[row, XLS_Table.E] = "=(H" + row + "/J" + row + ") * 100";
                    _sheet.Cells[row, XLS_Table.F] = cv_vol;
                    // check measure function
                    _sheet.Cells[row, XLS_Table.G] = UVP_amp;
                    _sheet.Cells[row, XLS_Table.H] = UVP_max;
                    _sheet.Cells[row, XLS_Table.I] = UVP_min;
                    _sheet.Cells[row, XLS_Table.J] = ori_vol;
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
            InsControl._scope.DoCommand(":MEASure:SOURce CHANnel1");
            stopWatch.Stop();

#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;

            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            for (int i = 1; i < 10; i++) _sheet.Columns[i].AutoFit();

            MyLib.SaveExcelReport(test_parameter.wave_path, temp + "C_UVP_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif

        }






    }
}
