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
        string[] cells = new string[] { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

        int row = 20;

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
            InsControl._oscilloscope.SetTimeScale(0.002);
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
                rail_info += test_parameter.rail_name[i] + ": aggressor load=";

                for (int j = 0; j < test_parameter.ccm_eload[i].Count; j++)
                {
                    rail_info += test_parameter.ccm_eload[i][j] + ((j == test_parameter.ccm_eload[i].Count - 1) ? "A" : "A, ");
                }
                rail_info += ",full load=" + test_parameter.full_load[i] + "\r\n";
            }

            _sheet.Cells[1, XLS_Table.B] = item;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver
                                            + test_parameter.vin_conditions
                                            + rail_info;
#endif

            OSCInit();

            // first item CCM
            if (test_parameter.cross_mode == 3)
                Cross_LT();
            else
                Cross_CCM();
            
        }

        private void MeasureVictim(int victim, int col_start, bool before)
        {

            double vmean = 0;
            double vmax = 0;
            double vmin = 0;
            double jitter = 0;

            vmean = InsControl._oscilloscope.CHx_Meas_Mean(victim);
            vmax = InsControl._oscilloscope.CHx_Meas_Max(victim);
            vmin = InsControl._oscilloscope.CHx_Meas_Min(victim);

            if (test_parameter.jitter_ch != 0)
                jitter = InsControl._oscilloscope.CHx_Meas_Jitter(test_parameter.jitter_ch);

            InsControl._oscilloscope.SetMeasureOff(1);
#if true
            // for measure victim channel
            int col_cnt = 7;
            if (before)
            {
                _sheet.Cells[row, col_start++] = vmean;
                _sheet.Cells[row, col_start++] = vmax;
                _sheet.Cells[row, col_start++] = vmin;
                _sheet.Cells[row, col_start++] = jitter;
                _sheet.Cells[row, col_start++] = vmax - vmean;
                _sheet.Cells[row, col_start++] = vmean - vmin;
            }
            else
            {
                _sheet.Cells[row, col_start++ + col_cnt] = vmax;
                _sheet.Cells[row, col_start++ + col_cnt] = vmin;
                _sheet.Cells[row, col_start++ + col_cnt] = jitter;
                _sheet.Cells[row, col_start++ + col_cnt] = vmax - vmean;
                _sheet.Cells[row, col_start + col_cnt] = vmean - vmin;
            }
#endif
        }

#region "Cross Talk CCM Mode"

        private void Cross_CCM()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int vin_cnt = test_parameter.VinList.Count;
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
                    if (test_parameter.ccm_eload[i].Count > switch_max)
                        switch_max = test_parameter.ccm_eload[i].Count;
                }
            }
            // ch_sw_num just judge that need to run how many times active load switch
            ch_sw_num = ch_sw_num / 2;

            //InsControl._power.AutoPowerOff();
            OSCInit();
            MyLib.Delay1ms(500);

            // the select_idx equal to aggressor channel
            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                if (test_parameter.cross_en[select_idx]) // select equal to aggressor
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
                                /* change aggressor vout */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.vout_addr[select_idx],
                                                new byte[] { test_parameter.vout_data[select_idx][vout_idx] });

                                /* change aggressor freq */
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                                                test_parameter.freq_addr[select_idx],
                                                new byte[] { test_parameter.freq_data[select_idx][freq_idx] });

                                int cnt_max = 0;
                                for(int cnt_idx = 0; cnt_idx < test_parameter.cross_en.Length; cnt_idx++)
                                {
                                    if(select_idx != cnt_idx)
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
                                    _sheet.Cells[row, col_start] = "Vout=" + test_parameter.vout_des[select_idx][vout_idx];
                                    _sheet.Cells[row++, XLS_Table.C] = "Vin=" + test_parameter.VinList[vin_idx] + "V";
                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, XLS_Table.C] = "Aggressor";
                                    _range = _sheet.Range["C" + (row), cells[test_parameter.ch_num] + (row)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_start] = "Freq (KHz) =" + test_parameter.freq_des[select_idx][freq_idx];
                                    row++;
                                    int col_idx = (int)XLS_Table.C;

                                    for (int i = 0; i < test_parameter.ch_num; i++)
                                    {
                                        if (i != select_idx)
                                        {
                                            _sheet.Cells[row, col_idx++] = test_parameter.rail_name[i] + "(A)";
                                        }
                                    }

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + " (A)";
                                    _sheet.Cells[row, col_base++] = "Vmean(V)";
                                    _sheet.Cells[row, col_base] = "Victim Max Voltage";
                                    _sheet.Cells[row - 1, col_base] = "Before: no load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 2] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    col_base++;

                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";
                                    _sheet.Cells[row, col_base] = "-VΔ (mV)";

                                    _range = _sheet.Range[cells[(int)XLS_Table.C + 2 + test_parameter.ch_num - 1] + (row - 1), cells[col_base - 1] + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    for (int i = 1; i < 25; i++)
                                        _sheet.Columns[i].AutoFit();
                                    row++;

                                    //FirstMeasure(select_idx);
                                    //_sheet.Cells[row - 2, col_start + 1] = test_parameter.freq_des[select_idx][freq_idx];
#endif
                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {

                                        double iout = (victim_idx == 0) ? 0 : victim_iout;
                                        int n = ch_sw_num == 2 ? 1 :
                                                ch_sw_num == 4 ? 2 :
                                                ch_sw_num == 8 ? 3 :
                                                ch_sw_num == 16 ? 4 :
                                                ch_sw_num == 32 ? 5 :
                                                ch_sw_num == 64 ? 6 : 7;

                                        MeasureN(   n,
                                                    select_idx,
                                                    vout_idx,
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
#if false
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_CrossTalk_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

        private void CHx_LevelReScale(int ch, int vout_idx)
        {
            double vout = Convert.ToDouble(test_parameter.vout_des[ch][vout_idx]);
            InsControl._oscilloscope.CHx_On(ch);
            InsControl._oscilloscope.CHx_Offset(ch, vout);
            InsControl._oscilloscope.CHx_Level(ch, 0.05); // set 50mV
        }

        private void MeasureN(int n, int select_idx, int vout_idx, int group, double iout_n, int col_start,
                                bool before, bool lt_mode = false)
        {
            int idx = 0;
            int[] sw_en = new int[n]; // save victim channel number
            double[] iout = new double[n];
            double[] l1 = new double[n];
            double[] l2 = new double[n];
            int loop_cnt = (int)Math.Pow(2, n);
            //Dictionary<int, List<double>> iout_list = new Dictionary<int, List<double>>();

            // save aggressor number and trun off aggressor channel 
            for (int aggressor = 0; aggressor < test_parameter.cross_en.Length; aggressor++)
            {
                if (aggressor != select_idx && test_parameter.cross_en[aggressor])
                {
                    sw_en[idx++] = aggressor;
                    InsControl._oscilloscope.CHx_Off(aggressor);
                }
            }

            CHx_LevelReScale(select_idx + 1, vout_idx);
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
                for (int j = 0; j < n; j++)
                {
                    switch (test_parameter.cross_mode)
                    {
                        case 0: // ccm mode
                            InsControl._eload.Loading(sw_en[j] + 1, iout[j]);
                            _sheet.Cells[row, j + aggressor_col] = data[j];
                            break;
                        case 1: // i2c on / off
                            _sheet.Cells[row, j + aggressor_col] = (data[j] == 1) ? "Enable" : "0";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                if (data[j] == 0) break;
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.en_data[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.disen_data[j] });
                            }
                            break;
                        case 2: // i2c VID
                            _sheet.Cells[row, j + aggressor_col] = (data[j] == 1) ? test_parameter.lo_code[j] + "->" + test_parameter.hi_code[j] : "0";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                if (data[j] == 0) break;
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.hi_code[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.lo_code[j] });
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

                
                if (lt_mode)
                {
                    // load transient mode
                    MeasureVictim(select_idx, col_start + 1, before);
                    if(before)
                        _sheet.Cells[row, before ? col_start : col_start + 7] = 0;
                    else
                        _sheet.Cells[row, before ? col_start : col_start + 7] =  "0 -> " + test_parameter.lt_full[select_idx].ToString();
                }
                else
                {
                    // others mode
                    MeasureVictim(select_idx, col_start + 1, before);
                    _sheet.Cells[row, before ? col_start : col_start + 7] = iout_n;
                }
                row++;
            }

            // turn on all of scope channel
            for (int i = 0; i < 4; i++)
                InsControl._oscilloscope.CHx_On(i + 1);
        }

#endregion


#region "Cross Talk Load transient" 

        private void Cross_LT()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int vin_cnt = test_parameter.VinList.Count;
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
#if true
                                    int col_base = (int)XLS_Table.C + 2 + test_parameter.ch_num;
                                    int col_start = col_base;
                                    _sheet.Cells[row, col_start] = "Vout=" + test_parameter.vout_des[select_idx][vout_idx];
                                    _sheet.Cells[row++, XLS_Table.C] = "Vin=" + test_parameter.VinList[vin_idx] + "V";
                                    _range = _sheet.Range["C" + (row - 1), cells[test_parameter.ch_num] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, XLS_Table.C] = "Aggressor";
                                    _range = _sheet.Range["C" + (row), cells[test_parameter.ch_num] + (row)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_start] = "Freq (KHz)=" + test_parameter.freq_des[select_idx][freq_idx];
                                    row++;
                                    int col_idx = (int)XLS_Table.C;
                                    for (int i = 0; i < test_parameter.ch_num; i++)
                                    {
                                        if (i != select_idx)
                                        {
                                            _sheet.Cells[row, col_idx++] = test_parameter.rail_name[i] + "(A)";
                                        }
                                    }

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + " (A)";
                                    _sheet.Cells[row, col_base++] = "Vmean(V)";
                                    _sheet.Cells[row, col_base] = "Victim Max Voltage";
                                    _sheet.Cells[row - 1, col_base] = "Before: no load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    col_base++;

                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 4] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";
                                    _sheet.Cells[row, col_base] = "-VΔ (mV)";

                                    _range = _sheet.Range[cells[(int)XLS_Table.C + 2 + test_parameter.ch_num - 1] + (row - 1), cells[col_base - 1] + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    for (int i = 1; i < 25; i++)
                                        _sheet.Columns[i].AutoFit();
                                    row++;
#endif

                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {
                                        double l1 = 0;
                                        double l2 = victim_idx == 0 ? 0 : test_parameter.lt_full[select_idx];
                                        InsControl._eload.DymanicLoad(select_idx + 1, l1, l2, 500, 500); // t1, t2 unit is us
                                        int n = ch_sw_num == 2 ? 1 :
                                        ch_sw_num == 4 ? 2 :
                                        ch_sw_num == 8 ? 3 :
                                        ch_sw_num == 16 ? 4 :
                                        ch_sw_num == 32 ? 5 :
                                        ch_sw_num == 64 ? 6 : 7;

                                        MeasureN(   n,                                  // select total number
                                                    select_idx,                         // victim number
                                                    vout_idx,                           // vout setting print to excel
                                                    group_idx,                          // maybe has more test conditions
                                                    l2,                                 // iout value to excel
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
        }

#endregion
    }
}



//private void Measure2(int aggressor, int group, double iout)
//{
//    int sw_ch = 0;
//    for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
//    {
//        if (aggressor != victim && test_parameter.cross_en[victim]) sw_ch = victim;
//    }

//    //ReserveMeasureChannel(aggressor);

//    for (int idx = 0; idx < 2; idx++)
//    {
//        switch (idx)
//        {
//            case 0:
//                //InsControl._eload.Loading(sw_ch, 0);
//                _sheet.Cells[row, XLS_Table.B] = 0;
//                break;
//            case 1:
//                //InsControl._eload.Loading(sw_ch, test_parameter.ccm_eload[sw_ch][group]);
//                _sheet.Cells[row, XLS_Table.B] = test_parameter.ccm_eload[sw_ch][group];
//                break;
//        }
//        //MeasureVictim(aggressor);
//        _sheet.Cells[row, XLS_Table.M] = iout;
//        row++;
//    }
//}

//private void Measure4(int aggressor, int group, double iout_n, int col_start, bool before)
//{
//    // program flow
//    // find enable channel
//    // get group iout setting
//    int idx = 0;
//    int[] sw_en = new int[2];
//    double[] iout = new double[2];
//    for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
//    {
//        if (victim != aggressor && test_parameter.cross_en[victim])
//        {
//            sw_en[idx++] = victim;
//        }
//    }

//    //ReserveMeasureChannel(aggressor);

//    // if group setting is differenct. I need found max current setting
//    iout[0] = group < test_parameter.ccm_eload[sw_en[0]].Count ?
//        test_parameter.ccm_eload[sw_en[0]][group] : test_parameter.ccm_eload[sw_en[0]].Max();

//    iout[1] = group < test_parameter.ccm_eload[sw_en[1]].Count ?
//        test_parameter.ccm_eload[sw_en[1]][group] : test_parameter.ccm_eload[sw_en[1]].Max();

//    for (idx = 0; idx < 4; idx++)
//    {
//        switch (idx)
//        {
//            case 0:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);
//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                break;
//            case 1:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                break;
//            case 2:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);

//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                break;
//            case 3:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                break;
//        }
//        MeasureVictim(aggressor, col_start + 1, before);
//        _sheet.Cells[row, before ? col_start : col_start + 6] = iout_n;
//        row++;
//    }
//}

//private void Measure8(int aggressor, int group, double iout_n, int col_start, bool before)
//{
//    int idx = 0;
//    int[] sw_en = new int[3];
//    double[] iout = new double[3];
//    for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
//    {
//        if (victim != aggressor && test_parameter.cross_en[victim])
//        {
//            sw_en[idx++] = victim;
//        }
//    }

//    //ReserveMeasureChannel(aggressor);

//    iout[0] = group < test_parameter.ccm_eload[sw_en[0]].Count ?
//        test_parameter.ccm_eload[sw_en[0]][group] : test_parameter.ccm_eload[sw_en[0]].Max();

//    iout[1] = group < test_parameter.ccm_eload[sw_en[1]].Count ?
//        test_parameter.ccm_eload[sw_en[1]][group] : test_parameter.ccm_eload[sw_en[1]].Max();

//    iout[2] = group < test_parameter.ccm_eload[sw_en[2]].Count ?
//        test_parameter.ccm_eload[sw_en[2]][group] : test_parameter.ccm_eload[sw_en[2]].Max();

//    for (idx = 0; idx < 8; idx++)
//    {
//        switch (idx)
//        {
//            case 0:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);
//                //InsControl._eload.Loading(sw_en[2] + 1, 0);

//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                _sheet.Cells[row, XLS_Table.D] = 0;
//                break;
//            case 1:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);
//                //InsControl._eload.Loading(sw_en[2] + 1, 0);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                _sheet.Cells[row, XLS_Table.D] = 0;
//                break;
//            case 2:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
//                //InsControl._eload.Loading(sw_en[2] + 1, 0);

//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                _sheet.Cells[row, XLS_Table.D] = 0;
//                break;
//            case 3:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
//                //InsControl._eload.Loading(sw_en[2] + 1, 0);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                _sheet.Cells[row, XLS_Table.D] = 0;
//                break;
//            case 4:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);
//                //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                _sheet.Cells[row, XLS_Table.D] = iout[2];
//                break;
//            case 5:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, 0);
//                //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = 0;
//                _sheet.Cells[row, XLS_Table.D] = iout[2];
//                break;
//            case 6:
//                //InsControl._eload.Loading(sw_en[0] + 1, 0);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
//                //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

//                _sheet.Cells[row, XLS_Table.B] = 0;
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                _sheet.Cells[row, XLS_Table.D] = iout[2];
//                break;
//            case 7:
//                //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
//                //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
//                //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

//                _sheet.Cells[row, XLS_Table.B] = iout[0];
//                _sheet.Cells[row, XLS_Table.C] = iout[1];
//                _sheet.Cells[row, XLS_Table.D] = iout[2];
//                break;
//        }
//        MeasureVictim(aggressor, col_start + 1, before);
//        _sheet.Cells[row, before ? col_start : col_start + 6] = iout_n;
//        row++;
//    }
//}
