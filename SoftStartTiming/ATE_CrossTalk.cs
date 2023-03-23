using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;


namespace SoftStartTiming
{
    public class ATE_CrossTalk : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        string[] cells = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z",
                                        "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"};

        int row = 20;
        int[] col_pos = new int[17];
        
        enum Col_List
        {
            b_Vmean = 0, b_Vmax, b_Vmin, b_jitter, b_delta_pos, b_delta_neg, 
            a_Vmax, a_min, a_jitter, a_delta_pos, a_delta_neg, 
            delta_pos, delta_neg, tol_pos, tol_neg, res_pos, res_neg,
            
        };

        public double temp;
        RTBBControl RTDev = new RTBBControl();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_CrossTalk()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Cross Talk test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void OSCInit()
        {
            InsControl._oscilloscope.SetAutoTrigger();
            InsControl._oscilloscope.SetTimeScale(test_parameter.offtime_scale_ms);

            for (int i = 0; i < 4; i++)
            {
                InsControl._oscilloscope.CHx_Off(i + 1);
            }

            for (int i = 0; i < test_parameter.ch_num; i++)
            {
                InsControl._oscilloscope.CHx_On(i + 1);
                InsControl._oscilloscope.CHx_Position(i + 1, 0);
            }

        }

        public override void ATETask()
        {
            RTDev.BoadInit();
#if true
            // Excel initial
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

            string item = "Cross Talk: ";

            switch (test_parameter.cross_mode)
            {
                case 0:
                    item += "CCM";
                    break;
                case 1:
                    item += "EN on/off";
                    break;
                case 2:
                    item += "VID";
                    break;
                case 3:
                    item += "LT";
                    break;
            }

            string rail_info = "";
            for (int i = 0; i < test_parameter.ch_num; i++)
            {
                rail_info += test_parameter.rail_name[i];
                switch (test_parameter.cross_mode)
                {
                    case 0: // ccm
                        rail_info += ": load=";
                        for (int j = 0; j < test_parameter.ccm_eload[i].Count; j++)
                        {
                            rail_info += test_parameter.ccm_eload[i][j] + ((j == test_parameter.ccm_eload[i].Count - 1) ? "A" : "A, ");
                        }
                        break;
                    case 1: // en on / off
                        rail_info += string.Format("Addr={0:X2}_ON={1:X2}_OFF={1:X2}",
                                                    test_parameter.en_addr[i],
                                                    test_parameter.en_data[i],
                                                    test_parameter.disen_data[i]
                                                    );
                        break;
                    case 2: // vid
                        rail_info += string.Format("Addr={0:X2}_Hi={1:X2}_Lo={1:X2}",
                                        test_parameter.vid_addr[i],
                                        test_parameter.hi_code[i],
                                        test_parameter.lo_code[i]
                                        );
                        break;
                    case 3: // LT
                        rail_info += ": L1=";

                        for(int j = 0; j < test_parameter.lt_l1[i].Count; j++)
                        {
                            rail_info += test_parameter.lt_l1[i][j] + (j == test_parameter.lt_l1[i].Count ? "A" : "A, ");
                        }

                        rail_info += " L2=";

                        for (int j = 0; j < test_parameter.lt_l2[i].Count; j++)
                        {
                            rail_info += test_parameter.lt_l2[i][j] + (j == test_parameter.lt_l2[i].Count ? "A" : "A, ");
                        }

                        break;
                }

                rail_info += ",full load=" + test_parameter.full_load[i] + "\r\n";
            }

            _sheet.Cells[1, XLS_Table.B] = item;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver
                                            + test_parameter.vin_conditions
                                            + rail_info;
#endif
            //OSCInit();
            if (test_parameter.cross_mode == 3)
                Cross_LT();
            else
                Cross_CCM();
        }

        private void MeasureVictim(int victim, int col_start, double vout, bool before)
        {
            double vmean = 0;
            double vmax = 0;
            double vmin = 0;
            double jitter = 0;

            for (int i = 0; i < 5; i++)
            {
                vmean = InsControl._oscilloscope.CHx_Meas_Mean(victim, 1);
                vmax = InsControl._oscilloscope.CHx_Meas_Max(victim, 2);
                vmin = InsControl._oscilloscope.CHx_Meas_Min(victim, 3);

                if (victim == 0 && test_parameter.Lx1) jitter = InsControl._oscilloscope.CHx_Meas_Jitter(3, 4);
                if (victim == 1 && test_parameter.Lx2) jitter = InsControl._oscilloscope.CHx_Meas_Jitter(4, 4);
            }

            InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name);
            InsControl._oscilloscope.SetMeasureOff(1);
            InsControl._oscilloscope.SetMeasureOff(2);
            InsControl._oscilloscope.SetMeasureOff(3);
            InsControl._oscilloscope.SetMeasureOff(4);
#if true
            // for measure victim channel
            int col_cnt = 7;
            double pos_delta = (vmax - vmean) * 1000;
            double neg_delta = (vmean - vmin) * 1000;
            if (before)
            {
                _sheet.Cells[row, col_pos[(int)Col_List.b_Vmean]] = vmean;
                _sheet.Cells[row, col_pos[(int)Col_List.b_Vmax]] = vmax;
                _sheet.Cells[row, col_pos[(int)Col_List.b_Vmin]] = vmin;
                _sheet.Cells[row, col_pos[(int)Col_List.b_jitter]] = jitter;
                _sheet.Cells[row, col_pos[(int)Col_List.b_delta_pos]] = string.Format("{0:0.000}", pos_delta);
                _sheet.Cells[row, col_pos[(int)Col_List.b_delta_neg]] = string.Format("{0:0.000}", neg_delta);

            }
            else
            {
                col_start++;
                _sheet.Cells[row, col_pos[(int)Col_List.a_Vmax]] = vmax;
                _sheet.Cells[row, col_pos[(int)Col_List.a_min]] = vmin;
                _sheet.Cells[row, col_pos[(int)Col_List.a_jitter]] = jitter;
                _sheet.Cells[row, col_pos[(int)Col_List.a_delta_pos]] = string.Format("{0:0.000}", pos_delta); ; // + delta
                _sheet.Cells[row, col_pos[(int)Col_List.a_delta_neg]] = string.Format("{0:0.000}", neg_delta); // - delta

                //col_start += 2;
                _sheet.Cells[row, col_pos[(int)Col_List.delta_pos]] = string.Format("={0}{1}-{2}{3}",
                                                            cells[col_pos[(int)Col_List.a_delta_pos] - 1], row, cells[col_pos[(int)Col_List.b_delta_pos] - 1], row);
                _sheet.Cells[row, col_pos[(int)Col_List.delta_neg]] = string.Format("={0}{1}-{2}{3}",
                                                            cells[col_pos[(int)Col_List.a_delta_neg] - 1], row, cells[col_pos[(int)Col_List.b_delta_neg] - 1], row);

                _range = _sheet.Cells[row, col_pos[(int)Col_List.tol_pos]];
                _range.NumberFormat = "0.000%";
                _sheet.Cells[row, col_pos[(int)Col_List.tol_pos]] = string.Format("={0}{1} / 1000 / {2}{3}",
                                                             cells[col_pos[(int)Col_List.delta_pos] - 1], row, cells[col_pos[(int)Col_List.b_Vmean] - 1], row);

                _range = _sheet.Cells[row, col_pos[(int)Col_List.tol_neg]];
                _range.NumberFormat = "0.000%";
                _sheet.Cells[row, col_pos[(int)Col_List.tol_neg]] = string.Format("={0}{1} / 1000 / {2}{3}",
                                                             cells[col_pos[(int)Col_List.delta_neg] - 1], row, cells[col_pos[(int)Col_List.b_Vmean] - 1], row);



                _sheet.Cells[row, col_pos[(int)Col_List.res_pos]] = string.Format("=IF({0}{1} < {2}, \"PASS\",\"FAIL\")",
                                                                    cells[col_pos[(int)Col_List.tol_pos] - 1], row, test_parameter.tolerance / 100);
                
                _sheet.Cells[row, col_pos[(int)Col_List.res_neg]] = string.Format("=IF({0}{1} < {2}, \"PASS\",\"FAIL\")",
                                                                    cells[col_pos[(int)Col_List.tol_neg] - 1], row, test_parameter.tolerance / 100);


                _range = _sheet.Range["$" + cells[col_pos[(int)Col_List.res_pos] - 1] + "$" + row + ":$" + cells[col_pos[(int)Col_List.res_neg] - 1] + "$" + row];
                Excel.FormatCondition format = _range.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "PASS");
                format.Interior.Color = Excel.XlRgbColor.rgbGreen;


                _range = _sheet.Range["$" + cells[col_pos[(int)Col_List.res_pos] - 1] + "$" + row + ":$" + cells[col_pos[(int)Col_List.res_neg] - 1] + "$" + row];
                format = _range.FormatConditions.Add(Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlEqual, "FAIL");
                format.Interior.Color = Excel.XlRgbColor.rgbRed;

            }
#endif
        }

        #region " ---- Cross Talk ---- "

        private void Cross_CCM()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int vin_cnt = test_parameter.VinList.Count;
            row = 18;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);
            int file_idx = 0;

            int switch_max = 0;
            int ch_sw_num = 1;
            for (int i = 0; i < test_parameter.ccm_eload.Count; i++)
            {
                if (test_parameter.cross_en[i])
                {
                    ch_sw_num = ch_sw_num * 2;
                    if (test_parameter.ccm_eload[i].Count > switch_max)
                        switch_max = test_parameter.ccm_eload[i].Count;
                }
            }
            // ch_sw_num just judge that need to run how many times active load switch
            ch_sw_num = ch_sw_num / 2;

            //InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1ms(500);

            // the select_idx equal to vimtic channel
            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                if (test_parameter.cross_en[select_idx]) // select equal to victim
                {
                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                    {
                        // vin loop
                        InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);

                        for (int vout_idx = 0; vout_idx < test_parameter.vout_data[select_idx].Count; vout_idx++)
                        {
                            // vout loop
                            for (int freq_idx = 0; freq_idx < test_parameter.freq_data[select_idx].Count; freq_idx++)
                            {
                                // freq loop
                                /* change victim vout */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.vout_addr[select_idx],
                                                new byte[] { test_parameter.vout_data[select_idx][vout_idx] });

                                /* change victim freq */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.freq_addr[select_idx],
                                                new byte[] { test_parameter.freq_data[select_idx][freq_idx] });

                                int cnt_max = 0;
                                for (int cnt_idx = 0; cnt_idx < test_parameter.cross_en.Length; cnt_idx++)
                                {
                                    if (select_idx != cnt_idx)
                                    {
                                        cnt_max = select_idx != 0 ? test_parameter.ccm_eload[0].Count : test_parameter.ccm_eload[1].Count;
                                        if (cnt_max < test_parameter.ccm_eload[cnt_idx].Count)
                                            cnt_max = test_parameter.ccm_eload[cnt_idx].Count;
                                    }
                                }

                                // victim current select
                                for (int group_idx = 0; group_idx < cnt_max; group_idx++) // how many iout group
                                {

                                    //double victim_iout = group_idx < test_parameter.ccm_eload[select_idx].Count() ?
                                    //                    test_parameter.ccm_eload[select_idx][group_idx] : test_parameter.ccm_eload[select_idx].Max();

                                    double victim_iout = test_parameter.full_load[select_idx];
                                    int col_base = (int)XLS_Table.C + 2 + test_parameter.ch_num;
                                    int col_start = col_base;

#if true
                                    _sheet.Cells[row, col_start] = string.Format("Vout={0}, Addr={1:X2}, Data={2:X2}"
                                                                    , test_parameter.vout_des[select_idx][vout_idx]
                                                                    , test_parameter.vout_addr[select_idx]
                                                                    , test_parameter.vout_data[select_idx][vout_idx]);

                                    _range = _sheet.Cells[row, col_start];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _range = _sheet.Range[cells[col_start - 1] + row, cells[col_start - 1] + (row + 2)];
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);


                                    _sheet.Cells[row++, XLS_Table.C] = "Vin=" + test_parameter.VinList[vin_idx] + "V";
                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row + 1)];
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);


                                    _sheet.Cells[row, XLS_Table.C] = "Aggressor";
                                    _range = _sheet.Range["C" + (row), cells[test_parameter.ch_num] + (row)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_start] = string.Format("Freq(KHz)={0}, Addr={1:X2}, Data={2:X2}"
                                                                    , test_parameter.freq_des[select_idx][freq_idx]
                                                                    , test_parameter.freq_addr[select_idx]
                                                                    , test_parameter.freq_data[select_idx][freq_idx]);


                                    row++;
                                    int col_idx = (int)XLS_Table.C;
                                    for (int i = 0; i < test_parameter.ch_num; i++)
                                    {
                                        if (i != select_idx)
                                        {
                                            _range = _sheet.Cells[row, col_idx];
                                            _sheet.Cells[row, col_idx++] = test_parameter.rail_name[i] + "(A), Vout=" + test_parameter.vout_des[i][vout_idx];
                                            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        }
                                    }

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + " (A)";


                                    col_pos[(int)Col_List.b_Vmean] = col_base;
                                    _sheet.Cells[row, col_base++] = "Vmean(V)";

                                    col_pos[(int)Col_List.b_Vmax] = col_base;
                                    _sheet.Cells[row, col_base] = "Victim Max Voltage";

                                    _sheet.Cells[row - 1, col_base] = "Before: no load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row)];
                                    _range.Interior.Color = Color.FromArgb(0xCC, 0xFF, 0xEF);
                                    col_base++;

                                    col_pos[(int)Col_List.b_Vmin] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.b_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";

                                    col_pos[(int)Col_List.b_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.b_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";
                                    //_sheet.Cells[row, col_base++] = "+ Tol (%)";
                                    //_sheet.Cells[row, col_base++] = "- Tol (%)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 4] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";

                                    col_pos[(int)Col_List.a_Vmax] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";

                                    col_pos[(int)Col_List.a_min] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.a_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";

                                    col_pos[(int)Col_List.a_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.a_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base] = "-VΔ (mV)";

                                    _range = _sheet.Range[cells[(int)XLS_Table.C + 2 + test_parameter.ch_num - 1] + (row - 1), cells[col_base - 1] + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    col_base += 2;

                                    col_pos[(int)Col_List.delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";

                                    col_pos[(int)Col_List.tol_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+ Tol (%)";

                                    col_pos[(int)Col_List.tol_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "- Tol (%)";

                                    col_pos[(int)Col_List.res_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+ Tol (Result)";

                                    col_pos[(int)Col_List.res_neg] = col_base;
                                    _sheet.Cells[row, col_base] = "- Tol (Result)";

                                    for (int i = 1; i < 25; i++)
                                        _sheet.Columns[i].AutoFit();
                                    row++;

                                    //FirstMeasure(select_idx);
                                    //_sheet.Cells[row - 2, col_start + 1] = test_parameter.freq_des[select_idx][freq_idx];
#endif
                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {

                                        double iout = (victim_idx == 0) ? 0 : victim_iout;
                                        test_parameter.waveform_name = string.Format("{0}_{1}_VIN={2}_Vout={3}_Freq={4}_Iout={5}",
                                                                        file_idx++,
                                                                        test_parameter.rail_name[select_idx],
                                                                        test_parameter.VinList[vin_idx],
                                                                        test_parameter.vout_des[select_idx][vout_idx],
                                                                        test_parameter.freq_des[select_idx][freq_idx],
                                                                        iout
                                                                        );
                                        int n = ch_sw_num == 2 ? 1 :
                                                ch_sw_num == 4 ? 2 :
                                                ch_sw_num == 8 ? 3 :
                                                ch_sw_num == 16 ? 4 :
                                                ch_sw_num == 32 ? 5 :
                                                ch_sw_num == 64 ? 6 : 7;

                                        if (iout != 0)
                                            InsControl._eload.Loading(select_idx + 1, iout);
                                        MeasureN(   n,
                                                    select_idx,
                                                    Convert.ToDouble(test_parameter.vout_des[select_idx][vout_idx]),
                                                    group_idx,
                                                    iout,
                                                    col_start,
                                                    victim_idx == 0 ? true : false);

                                        if (victim_idx == 0) row = row - ch_sw_num;
                                    }

                                    row += 3;
                                } // iout group loop
                            } // vout loop
                        } // freq loop
                    } // vin loop
                } // select aggressor loop
            }
            stopWatch.Stop();
#if true
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_CrossTalk_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

        private void Cross_LT()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int vin_cnt = test_parameter.VinList.Count;
            int file_idx = 0;
            row = 18;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

            int switch_max = 0;
            int ch_sw_num = 1;
            for (int i = 0; i < test_parameter.ccm_eload.Count; i++)
            {
                if (test_parameter.cross_en[i])
                {
                    ch_sw_num = ch_sw_num * 2;
                    if (test_parameter.lt_l1[i].Count > switch_max || test_parameter.lt_l2[i].Count > switch_max)
                    {
                        //switch_max = test_parameter.ccm_eload[i].Count;
                        if (test_parameter.lt_l1[i].Count > switch_max)
                            switch_max = test_parameter.lt_l1[i].Count;
                        else
                            switch_max = test_parameter.lt_l2[i].Count;
                    }

                }
            }
            // ch_sw_num just judge that need to run how many times active load switch
            ch_sw_num = ch_sw_num / 2;

            //InsControl._power.AutoPowerOff();
            //OSCInit();
            MyLib.Delay1ms(500);

            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                if (test_parameter.cross_en[select_idx]) // select equal to aggressor
                {
                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                    {
                        for (int vout_idx = 0; vout_idx < test_parameter.vout_data[select_idx].Count; vout_idx++)
                        {
                            for (int freq_idx = 0; freq_idx < test_parameter.freq_data[select_idx].Count; freq_idx++)
                            {

                                // freq loop
                                /* change aggressor vout */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.vout_addr[select_idx],
                                                new byte[] { test_parameter.vout_data[select_idx][vout_idx] });

                                /* change aggressor freq */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.freq_addr[select_idx],
                                                new byte[] { test_parameter.freq_data[select_idx][freq_idx] });

                                int cnt_max_l2 = 0;
                                int cnt_max_l1 = 0;
                                int cnt_max = 0;
                                for (int cnt_idx = 0; cnt_idx < test_parameter.cross_en.Length; cnt_idx++)
                                {
                                    if (select_idx != cnt_idx)
                                    {
                                        cnt_max_l2 = select_idx != 0 ? test_parameter.lt_l2[0].Count : test_parameter.lt_l2[1].Count;
                                        if (cnt_max_l2 < test_parameter.lt_l2[cnt_idx].Count)
                                            cnt_max_l2 = test_parameter.lt_l2[cnt_idx].Count;

                                        cnt_max_l1 = select_idx != 0 ? test_parameter.lt_l1[0].Count : test_parameter.lt_l1[1].Count;
                                        if (cnt_max_l1 < test_parameter.lt_l1[cnt_idx].Count)
                                            cnt_max_l1 = test_parameter.lt_l1[cnt_idx].Count;
                                    }
                                }
                                cnt_max = cnt_max_l2 > cnt_max_l1 ? cnt_max_l2 : cnt_max_l1;

                                for (int group_idx = 0; group_idx < cnt_max; group_idx++) // how many iout group
                                {
                                    double victim_iout = test_parameter.full_load[select_idx];
                                    int col_base = (int)XLS_Table.C + 2 + test_parameter.ch_num;
                                    int col_start = col_base;

#if true
                                    _sheet.Cells[row, col_start] = string.Format("Vout={0}, Addr={1:X2}, Data={2:X2}"
                                                                    , test_parameter.vout_des[select_idx][vout_idx]
                                                                    , test_parameter.vout_addr[select_idx]
                                                                    , test_parameter.vout_data[select_idx][vout_idx]);

                                    _range = _sheet.Cells[row, col_start];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _range = _sheet.Range[cells[col_start - 1] + row, cells[col_start - 1] + (row + 2)];
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);


                                    _sheet.Cells[row++, XLS_Table.C] = "Vin=" + test_parameter.VinList[vin_idx] + "V";
                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row + 1)];
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);


                                    _sheet.Cells[row, XLS_Table.C] = "Aggressor";
                                    _range = _sheet.Range["C" + (row), cells[test_parameter.ch_num] + (row)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_start] = string.Format("Freq(KHz)={0}, Addr={1:X2}, Data={2:X2}"
                                                                    , test_parameter.freq_des[select_idx][freq_idx]
                                                                    , test_parameter.freq_addr[select_idx]
                                                                    , test_parameter.freq_data[select_idx][freq_idx]);


                                    row++;
                                    int col_idx = (int)XLS_Table.C;
                                    for (int i = 0; i < test_parameter.ch_num; i++)
                                    {
                                        if (i != select_idx)
                                        {
                                            _range = _sheet.Cells[row, col_idx];
                                            _sheet.Cells[row, col_idx++] = test_parameter.rail_name[i] + "(A), Vout=" + test_parameter.vout_des[i][vout_idx];
                                            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                        }
                                    }

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + " (A)";


                                    col_pos[(int)Col_List.b_Vmean] = col_base;
                                    _sheet.Cells[row, col_base++] = "Vmean(V)";

                                    col_pos[(int)Col_List.b_Vmax] = col_base;
                                    _sheet.Cells[row, col_base] = "Victim Max Voltage";

                                    _sheet.Cells[row - 1, col_base] = "Before: no load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row)];
                                    _range.Interior.Color = Color.FromArgb(0xCC, 0xFF, 0xEF);
                                    col_base++;

                                    col_pos[(int)Col_List.b_Vmin] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.b_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";

                                    col_pos[(int)Col_List.b_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.b_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";
                                    //_sheet.Cells[row, col_base++] = "+ Tol (%)";
                                    //_sheet.Cells[row, col_base++] = "- Tol (%)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 4] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";

                                    col_pos[(int)Col_List.a_Vmax] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";

                                    col_pos[(int)Col_List.a_min] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.a_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";

                                    col_pos[(int)Col_List.a_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.a_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base] = "-VΔ (mV)";

                                    _range = _sheet.Range[cells[(int)XLS_Table.C + 2 + test_parameter.ch_num - 1] + (row - 1), cells[col_base - 1] + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    col_base += 2;

                                    col_pos[(int)Col_List.delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";

                                    col_pos[(int)Col_List.tol_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+ Tol (%)";

                                    col_pos[(int)Col_List.tol_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "- Tol (%)";

                                    col_pos[(int)Col_List.res_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+ Tol (Result)";

                                    col_pos[(int)Col_List.res_neg] = col_base;
                                    _sheet.Cells[row, col_base] = "- Tol (Result)";

                                    for (int i = 1; i < 25; i++)
                                        _sheet.Columns[i].AutoFit();
                                    row++;
#endif

                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {
                                        double iout = victim_idx == 0 ? 0 : victim_iout;
                                        test_parameter.waveform_name = string.Format("{0}_{1}_VIN={2}_Vout={3}_Freq={4}_Iout={5}",
                                                                file_idx++,
                                                                test_parameter.rail_name[select_idx],
                                                                test_parameter.VinList[vin_idx],
                                                                test_parameter.vout_des[select_idx][vout_idx],
                                                                test_parameter.freq_des[select_idx][freq_idx],
                                                                iout
                                                                );


                                        int n = ch_sw_num == 2 ? 1 :
                                        ch_sw_num == 4 ? 2 :
                                        ch_sw_num == 8 ? 3 :
                                        ch_sw_num == 16 ? 4 :
                                        ch_sw_num == 32 ? 5 :
                                        ch_sw_num == 64 ? 6 : 7;

                                        if (iout != 0)
                                            InsControl._eload.Loading(select_idx + 1, iout);

                                        MeasureN(   n,                                  // select total number
                                                    select_idx,                         // victim number
                                                    Convert.ToDouble(test_parameter.vout_des[select_idx][vout_idx]),                           // vout setting print to excel
                                                    group_idx,                          // maybe has more test conditions
                                                    iout,                                 // iout value to excel
                                                    col_start,                          // excel col start position
                                                    (victim_idx == 0 ? true : false),   // before & after
                                                    true);                              // lt mode enable

                                        if (victim_idx == 0) row = row - ch_sw_num;
                                    } // victim no load and full load

                                    row += 3;
                                } // group loop
                            } // freq loop
                        } // vout loop
                    } // vin loop
                } // channel select
            } // select

            stopWatch.Stop();
#if true
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_CrossTalk_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

        private void CHx_LevelReScale(int ch, double vout)
        {
            //double vout = Convert.ToDouble(test_parameter.vout_des[ch][vout_idx]);
            InsControl._oscilloscope.CHx_On(ch);
            InsControl._oscilloscope.CHx_Offset(ch, vout);
            InsControl._oscilloscope.CHx_Level(ch, 0.05); // set 50mV
            InsControl._oscilloscope.CHx_Position(ch, 0);
        }

        private void MeasureN(  int n, int select_idx, double vout,
                                int group, double iout_n, int col_start,
                                bool before, bool lt_mode = false)
        {
            int idx = 0;
            int[] sw_en = new int[n]; // save victim channel number
            double[] iout = new double[n];
            double[] l1 = new double[n];
            double[] l2 = new double[n];
            int loop_cnt = (int)Math.Pow(2, n);

            CHx_LevelReScale(select_idx + 1, vout);
            // save aggressor number and trun off aggressor channel 
            for (int aggressor = 0; aggressor < test_parameter.cross_en.Length; aggressor++)
            {
                if (aggressor != select_idx && test_parameter.cross_en[aggressor])
                {
                    sw_en[idx++] = aggressor;
                    InsControl._oscilloscope.CHx_Off(aggressor + 1);
                }
            }

            if (select_idx == 0 && test_parameter.Lx1) InsControl._oscilloscope.CHx_On(3);
            if (select_idx == 1 && test_parameter.Lx2) InsControl._oscilloscope.CHx_On(4);

            InsControl._oscilloscope.SetClear();
            InsControl._oscilloscope.SetPERSistence();

            // save aggressor iout conditions
            // iout select maximum setting if over iout list overflow.
            for (int i = 0; i < n; i++)
            {
                if (lt_mode)
                {
                    l1[i] = group < test_parameter.lt_l1[sw_en[i]].Count ?
                        test_parameter.lt_l1[sw_en[i]][group] : test_parameter.lt_l1[sw_en[i]].Max();

                    l2[i] = group < test_parameter.lt_l2[sw_en[i]].Count ?
                        test_parameter.lt_l2[sw_en[i]][group] : test_parameter.lt_l2[sw_en[i]].Max();
                }
                else
                {
                    iout[i] = group < test_parameter.ccm_eload[sw_en[i]].Count ?
                        test_parameter.ccm_eload[sw_en[i]][group] : test_parameter.ccm_eload[sw_en[i]].Max();
                }
            }

            // calculate and excute all of test conditions.
            for (int i = 0; i < loop_cnt; i++)
            {
                InsControl._eload.Loading(select_idx + 1, iout_n);
                List<double> data = new List<double>();
                List<double> data_l1 = new List<double>();
                List<double> data_l2 = new List<double>();
                int bit0 = (i & 0x01) >> 0;
                int bit1 = (i & 0x02) >> 1;
                int bit2 = (i & 0x04) >> 2;
                int bit3 = (i & 0x08) >> 3;
                int bit4 = (i & 0x10) >> 4;
                int bit5 = (i & 0x20) >> 5;
                int bit6 = (i & 0x40) >> 6;
                int bit7 = (i & 0x80) >> 7;
                int[] bit_list = new int[] { bit0, bit1, bit2, bit3, bit4, bit5, bit6, bit7 };

                // select test mode: CCM, EN on/off, VID, LT 
                for (int j = 0; j < n; j++)
                {
                    switch (test_parameter.cross_mode)
                    {
                        case 0:
                            data.Add(bit_list[j] == 0 ? 0 : iout[j]);
                            break;
                        case 1:
                        case 2:
                            // open active load
                            //InsControl._eload.Loading(sw_en[j] + 1, iout[j]);
                            data.Add(bit_list[j] == 0 ? 0 : 1);
                            break;
                        case 3:
                            data.Add(bit_list[j] == 0 ? 0 : 1);
                            data_l1.Add(bit_list[j] == 0 ? 0 : l1[j]);
                            data_l2.Add(bit_list[j] == 0 ? 0 : l2[j]);
                            break;
                    }
                }

                int aggressor_col = (int)XLS_Table.C;
                for (int j = 0; j < n; j++) // run each channel
                {
                    switch (test_parameter.cross_mode)
                    {
                        case 0: // ccm mode
                            if (data[j] != 0)
                            {
                                InsControl._eload.Loading(sw_en[j] + 1, iout[j]);
                                _sheet.Cells[row, j + aggressor_col] = InsControl._eload.GetIout();
                            }
                            else
                            {
                                _sheet.Cells[row, j + aggressor_col] = 0;
                            }

                            break;
                        case 1: // i2c on / off
                            _sheet.Cells[row, j + aggressor_col].NumberFormat = "@";
                            _sheet.Cells[row, j + aggressor_col] = (data[j] == 1) ? "Enable" : "0";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                if (data[j] == 0) break;
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.en_data[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.disen_data[j] });
                            }
                            break;
                        case 2: // i2c VID
                            _sheet.Cells[row, j + aggressor_col].NumberFormat = "@";
                            _sheet.Cells[row, j + aggressor_col] = (data[j] == 1) ? test_parameter.lo_code[j].ToString("X") + "->" + test_parameter.hi_code[j].ToString("X") : "0";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                if (data[j] == 0) break;
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.vid_addr[j], new byte[] { test_parameter.hi_code[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.vid_addr[j], new byte[] { test_parameter.lo_code[j] });
                            }
                            break;
                        case 3: // LT
                            _sheet.Cells[row, j + aggressor_col].NumberFormat = "@";
                            _sheet.Cells[row, j + aggressor_col] = (data[j] == 1) ? l1[j] + " -> " + l2[j] : "0";
                            // eload over 4CH need to select channel
                            InsControl._eload.DymanicLoad(sw_en[j] + 1, data_l1[j], data_l2[j], 500, 500); // 1KHz
                            break;
                    }
                }

                string temp = test_parameter.waveform_name;
                test_parameter.waveform_name = test_parameter.waveform_name + string.Format("_case{0}", i);
                MeasureVictim(select_idx + 1, col_start + 1, vout, before);
                test_parameter.waveform_name = temp;


                InsControl._eload.Loading(select_idx + 1, iout_n);
                _sheet.Cells[row, before ? col_start : col_start + 7] = InsControl._eload.GetIout();

                double[] read_iout = InsControl._eload.GetAllChannel_Iout();
                double[] read_vout = InsControl._eload.GetAllChannel_Vol();

                //Console.WriteLine("Vout1={0}\tVout2={1}\tVout3={2}\tVout3={3}", read_vout[0], read_vout[1], read_vout[2], read_vout[3]);
                //Console.WriteLine("[0]\t[1]\t[2]\t[3]");
                Console.Write("{0} = ", i);
                Console.WriteLine("Iout1={0}\tIout2={1}\tIout3={2}\tIout3={3}", read_iout[0], read_iout[1], read_iout[2], read_iout[3]);

                InsControl._eload.AllChannel_LoadOff();
                row++;
            }
        }

        #endregion


    }
}




