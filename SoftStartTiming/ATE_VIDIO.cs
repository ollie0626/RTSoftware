
//#define Report_en
//#define Power_en
//#define Eload_en

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;



namespace SoftStartTiming
{
    public class ATE_VIDIO : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        RTBBControl RTDev = new RTBBControl();

        const int LPM = 0;
        const int G1 = 1;
        const int G2 = 2;

        List<double> overshoot_list = new List<double>();
        List<double> undershoot_list = new List<double>();
        List<double> slewrate_list = new List<double>();

        private void IOStateSetting(int lpm, int g1, int g2)
        {
            int value = (lpm << 0 | g1 << 1 | g2 << 2);
            int mask = 1 << LPM | 1 << G1  | 1 << G2;
            RTDev.GPIOnState((uint)mask, (uint)value);
        }

        private void OSCInit()
        {
            InsControl._oscilloscope.CHx_On(1); // vout
            InsControl._oscilloscope.CHx_On(2); // Lx
            InsControl._oscilloscope.CHx_On(3); // G1
            InsControl._oscilloscope.CHx_On(4); // G2

            // initial time scale
            InsControl._oscilloscope.SetTimeScale(4 * Math.Pow(10, -6));

            InsControl._oscilloscope.CHx_Level(3, 2);
            InsControl._oscilloscope.CHx_Level(4, 2);
            InsControl._oscilloscope.CHx_Position(3, 2.5);
            InsControl._oscilloscope.CHx_Position(4, 2.5);

            double max = test_parameter.vidio.vout_list[0] > test_parameter.vidio.vout_list_af[0] ?
                         test_parameter.vidio.vout_list[0] : test_parameter.vidio.vout_list_af[0];
            double min = test_parameter.vidio.vout_list[0] < test_parameter.vidio.vout_list_af[0] ?
                         test_parameter.vidio.vout_list[0] : test_parameter.vidio.vout_list_af[0];
            InsControl._oscilloscope.CHx_Level(1, max - min / 3);
            InsControl._oscilloscope.CHx_Offset(1, min);
            InsControl._oscilloscope.CHx_Position(1, -2);

            InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[0] / 1.5);
            InsControl._oscilloscope.CHx_Position(2, -4);

            InsControl._oscilloscope.SetAutoTrigger();
            InsControl._oscilloscope.SetTriggerLevel(2, max - min);
        }

        private void RefelevelSel(bool diff)
        {
            InsControl._oscilloscope.SetREFLevelMethod();
            if (diff)
            {
                InsControl._oscilloscope.SetREFLevel(80, 50, 20);
            }
            else
            {
                InsControl._oscilloscope.SetREFLevel(100, 50, 0);
            }
            
        }

        private void Phase1Test(int case_idx)
        {
            overshoot_list.Clear();
            undershoot_list.Clear();
            slewrate_list.Clear();

            double vout = test_parameter.vidio.vout_list[case_idx];
            double vout_af = test_parameter.vidio.vout_list_af[case_idx];
            bool rising_en = vout < vout_af ? true : false;

            bool diff = Math.Abs(vout - vout_af) > 0.13 ? true : false;
            RefelevelSel(diff);

            if (rising_en)
            {
                InsControl._oscilloscope.SetTriggerRise();
                InsControl._oscilloscope.CHx_Level(1, (vout_af - vout) / 4.5);
                InsControl._oscilloscope.CHx_Offset(1, vout);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout_af - vout) * 0.3 + vout);
            }
            else
            {
                InsControl._oscilloscope.SetTriggerFall();
                InsControl._oscilloscope.CHx_Level(1, (vout - vout_af) / 4.5);
                InsControl._oscilloscope.CHx_Offset(1, vout_af);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout - vout_af) * 0.3 + vout_af);
            }

            for (int repeat_idx = 0; repeat_idx < 23; repeat_idx++)
            {
                double slew_rate = 0;
                double over_shoot = 0;
                double under_shoot = 0;
                // initial sate setting
                IOStateSetting(
                                test_parameter.vidio.lpm_sel[case_idx],
                                test_parameter.vidio.g1_sel[case_idx],
                                test_parameter.vidio.g2_sel[case_idx]
                                );
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(500);
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(100);
                // transfer condition
                IOStateSetting(
                                test_parameter.vidio.lpm_sel_af[case_idx],
                                test_parameter.vidio.g1_sel_af[case_idx],
                                test_parameter.vidio.g2_sel_af[case_idx]
                                );
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetStop();

                if (repeat_idx > 3)
                {
                    if (rising_en)
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, 1);
                        MyLib.Delay1ms(50);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, 1);
                        slewrate_list.Add(slew_rate);

                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(1, 2);
                        MyLib.Delay1ms(50);
                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(1, 2);
                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(1, 2);
                        overshoot_list.Add(over_shoot);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, 1);
                        MyLib.Delay1ms(50);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, 1);
                        slewrate_list.Add(slew_rate);

                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(1, 2);
                        MyLib.Delay1ms(50);
                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(1, 2);
                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(1, 2);
                        undershoot_list.Add(under_shoot);
                    }
                }
                else
                {
                    // calculate time scale function disable
                    //double time_scale = 0;
                    //if (rising_en)
                    //{
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //}
                    //else
                    //{
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //}

                    //if (time_scale < Math.Pow(10, 6) & time_scale != 0)
                    //{
                    //    InsControl._oscilloscope.SetTimeScale(time_scale / 3);
                    //    InsControl._oscilloscope.SetTimeBasePosition(2);
                    //}
                }
            }

            InsControl._oscilloscope.SetCursorMode();
            InsControl._oscilloscope.SetCursorOn();
            InsControl._oscilloscope.SetAnnotation(1);
            double x1 = InsControl._oscilloscope.GetAnnotationXn(1);
            double x2 = InsControl._oscilloscope.GetAnnotationXn(2);
            

            InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);
            InsControl._oscilloscope.SetCursorScreenYpos(vout, vout_af);

        }

        private void Phase2Test(int case_idx)
        {
            overshoot_list.Clear();
            undershoot_list.Clear();
            slewrate_list.Clear();

            double vout = test_parameter.vidio.vout_list[case_idx];
            double vout_af = test_parameter.vidio.vout_list_af[case_idx];
            bool rising_en = vout_af < vout ? true : false;

            if (rising_en)
                InsControl._oscilloscope.SetTriggerRise();
            else
                InsControl._oscilloscope.SetTriggerFall();


            for (int repeat_idx = 0; repeat_idx < 23; repeat_idx++)
            {
                double slew_rate = 0;
                double over_shoot = 0;
                double under_shoot = 0;

                // initial sate setting
                IOStateSetting(
                                test_parameter.vidio.lpm_sel_af[case_idx],
                                test_parameter.vidio.g1_sel_af[case_idx],
                                test_parameter.vidio.g2_sel_af[case_idx]
                                );
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(500);
                InsControl._oscilloscope.SetNormalTrigger();
                // transfer condition
                IOStateSetting(
                                test_parameter.vidio.lpm_sel[case_idx],
                                test_parameter.vidio.g1_sel[case_idx],
                                test_parameter.vidio.g2_sel[case_idx]
                                );


                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetStop();

                if (repeat_idx > 3)
                {
                    if (rising_en)
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(2, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(2, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(2, 1);
                        slewrate_list.Add(slew_rate);

                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(2, 2);
                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(2, 2);
                        over_shoot = InsControl._oscilloscope.CHx_Meas_Overshoot(2, 2);
                        overshoot_list.Add(over_shoot);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(2, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(2, 1);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(2, 1);
                        slewrate_list.Add(slew_rate);

                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(2, 2);
                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(2, 2);
                        under_shoot = InsControl._oscilloscope.CHx_Meas_Undershoot(2, 2);
                        undershoot_list.Add(under_shoot);
                    }
                }
                else
                {
                    //double time_scale = 0;
                    //if (rising_en)
                    //{
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Rise(2);
                    //}
                    //else
                    //{
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //    time_scale = InsControl._oscilloscope.CHx_Meas_Fall(2);
                    //}

                    //if (time_scale < Math.Pow(10, 6) & time_scale != 0)
                    //{
                    //    InsControl._oscilloscope.SetTimeScale(time_scale / 3);
                    //    InsControl._oscilloscope.SetTimeBasePosition(2);
                    //}
                }
            }
        }

        public override void ATETask()
        {
            RTDev.BoadInit();
            OSCInit();
            int row = 10;
            int wave_row = 10;
            int wave_idx = 0;
            int idx = 0;
            string file_name = "";
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Cells.Font.Name = "Calibri";
            _sheet.Cells.Font.Size = 11;

            _sheet.Cells[1, XLS_Table.A] = "Item";
            _sheet.Cells[2, XLS_Table.A] = "Test Conditions";
            _sheet.Cells[3, XLS_Table.A] = "Result";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _range = _sheet.Range["A1", "A4"];
            _range.Font.Bold = true;
            _range.Interior.Color = Color.FromArgb(255, 178, 102);
            _range = _sheet.Range["A2"];
            _range.RowHeight = 150;
            _range = _sheet.Range["B1"];
            _range.ColumnWidth = 60;
            _range = _sheet.Range["A1", "B4"];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _sheet.Cells[1, XLS_Table.B] = "VID_IO";
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions;
            
            // report title
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "超連結";
            _sheet.Cells[row, XLS_Table.D] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.E] = "Vout Change(V)";
            _sheet.Cells[row, XLS_Table.F] = "Iout (A)";
            _sheet.Cells[row, XLS_Table.G] = "Rise SR (us/V)";
            _sheet.Cells[row, XLS_Table.H] = "Fall SR (us/V)";
            _sheet.Cells[row, XLS_Table.I] = "Overshoot (%)";
            _sheet.Cells[row, XLS_Table.J] = "Undershoot (%)";

            _range = _sheet.Range["B" + row, "I" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            row++;
#endif

            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                for (int iout_idx = 0; iout_idx < test_parameter.IoutList.Count; iout_idx++)
                {
                    for (int case_idx = 0; case_idx < test_parameter.vidio.g1_sel.Count; case_idx++)
                    {
                        file_name = string.Format("{0}_Temp={1}_VIN={2}_IOUT={3}_Vout={4}_{5}",
                                                idx, temp,
                                                test_parameter.VinList[vin_idx],
                                                test_parameter.IoutList[iout_idx],
                                                test_parameter.vidio.vout_list[case_idx],
                                                test_parameter.vidio.vout_list_af[case_idx]
                                                );
#if Power_en
                        InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                        MyLib.Delay1ms(200);
#endif

#if Eload_en
                        MyLib.Switch_ELoadLevel(test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
#endif
                        InsControl._oscilloscope.SetAutoTrigger();

                        double vout = test_parameter.vidio.vout_list[case_idx];
                        double vout_af = test_parameter.vidio.vout_list_af[case_idx];
                        bool rising_en = vout < vout_af ? true : false;
#if Report_en
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = "LINK";
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[row, XLS_Table.E] = test_parameter.vidio.vout_list[case_idx] + "->" + test_parameter.vidio.vout_list_af[case_idx];
                        _sheet.Cells[row, XLS_Table.F] = test_parameter.IoutList[iout_idx];
#endif
                        Phase1Test(case_idx);
                        InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, file_name + (rising_en ? "_rising" : "_falling"));

#if Report_en
                        if (rising_en)
                        {
                            _sheet.Cells[row, XLS_Table.G] = slewrate_list.Min();
                            _sheet.Cells[row, XLS_Table.I] = overshoot_list.Max();
                        }
                        else
                        {
                            _sheet.Cells[row, XLS_Table.H] = slewrate_list.Min();
                            _sheet.Cells[row, XLS_Table.J] = undershoot_list.Max();
                        }
#endif
//-----------------------------------------------------------------------------------------
                        // Phase2Test(case_idx);
                        // InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, file_name + (!rising_en ? "_rising" : "_falling"));
                        //MyLib.PastWaveform(test_parameter.waveform_path,)
#if Report_en
                        if (!rising_en)
                        {
                            _sheet.Cells[row, XLS_Table.G] = slewrate_list.Min();
                            _sheet.Cells[row, XLS_Table.I] = overshoot_list.Max();
                        }
                        else
                        {
                            _sheet.Cells[row, XLS_Table.H] = slewrate_list.Min();
                            _sheet.Cells[row, XLS_Table.J] = undershoot_list.Max();
                        }
#endif
                        //-----------------------------------------------------------------------------------------

#if Report_en
                        Excel.Range main_range = _sheet.Range["C" + row];
                        Excel.Range hyper = _sheet.Range["M" + wave_row + 1];
                        // A to B
                        _sheet.Hyperlinks.Add(main_range, "#'" + _sheet.Name + "'!M" + (wave_row + 1));
                        _sheet.Hyperlinks.Add(hyper, "#'" + _sheet.Name + "'!C" + row);


                        _sheet.Cells[wave_row, XLS_Table.M] = "超連結";
                        _sheet.Cells[wave_row, XLS_Table.N] = "VIN";
                        _sheet.Cells[wave_row, XLS_Table.O] = "Vout";
                        _sheet.Cells[wave_row, XLS_Table.P] = "Iout";

                        _sheet.Cells[wave_row + 1, XLS_Table.M] = "Go back";
                        _sheet.Cells[wave_row + 1, XLS_Table.N] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[wave_row + 1, XLS_Table.O] = test_parameter.vidio.vout_list[case_idx] + "->" + test_parameter.vidio.vout_list_af[case_idx];
                        _sheet.Cells[wave_row + 1, XLS_Table.P] = test_parameter.IoutList[case_idx];

                        _range = _sheet.Range["M" + (wave_row + 2), "U" + (wave_row + 16)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + (rising_en ? "_rising" : "_falling"));
                        _range = _sheet.Range["W" + (wave_row + 2), "AE" + (wave_row + 16)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + (!rising_en ? "_rising" : "_falling"));
#endif

                        InsControl._oscilloscope.SetAutoTrigger();
                        wave_row += 21;
                        row++;
                    }
                }
            }
        } // function end

    }
}
