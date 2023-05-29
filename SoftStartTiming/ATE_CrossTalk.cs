
#define Report_en
#define Power_en
#define Eload_en


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
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;

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
        int progress = 0;

        double first_vmean;

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
        CrossTalk updateMain;

        Thread dont_stop;
        ParameterizedThreadStart p_dont_stop;
        static volatile int dont_stop_cnt = 0;


        public void Channel_Switch_event(object obj)
        {
            int aggressor_col = (int)XLS_Table.C;

            CrossTalkParameter parameter = (CrossTalkParameter)obj;

            while (true)
            {
                switch (test_parameter.cross_mode)
                {
                    case 0: // ccm mode
                        if (parameter.data[parameter.idx] != 0)
                        {
                            double temp = 0;
#if Eload_en
                            InsControl._eload.Loading(parameter.sw_en[parameter.idx] + 1, parameter.iout[parameter.idx]);
                            temp = InsControl._eload.GetIout();
#endif

#if Report_en

                            _sheet.Cells[row, parameter.idx + aggressor_col] = temp;
#endif
                        }
                        else
                        {
#if Eload_en
                            InsControl._eload.LoadOFF(parameter.sw_en[parameter.idx] + 1);
#endif

#if Report_en
                            _sheet.Cells[row, parameter.idx + aggressor_col] = 0;
#endif
                        }
                        break;
                    case 1:
                        _sheet.Cells[row, parameter.idx + aggressor_col].NumberFormat = "@";
                        _sheet.Cells[row, parameter.idx + aggressor_col] = (parameter.data[parameter.idx] == 1) ? "Enable" : "0";

                        List<byte> en_addr = new List<byte>();
                        List<byte> en_data = new List<byte>();
                        List<byte> disen_data = new List<byte>();
                        en_addr = test_parameter.en_addr.ToList();
                        en_data = test_parameter.en_data.ToList();
                        disen_data = test_parameter.disen_data.ToList();

                        en_addr.Remove(test_parameter.en_addr[parameter.select_idx]);
                        en_data.Remove(test_parameter.en_data[parameter.select_idx]);
                        disen_data.Remove(test_parameter.disen_data[parameter.select_idx]);

                        WriteEn(parameter.data, en_addr.ToArray(), en_data.ToArray(), disen_data.ToArray(), parameter.select_idx);
                        break;
                    case 2: // i2c VID

                        List<byte> vid_addr = new List<byte>();
                        List<byte> vid_low = new List<byte>();
                        List<byte> vid_high = new List<byte>();

                        for (int k = 0; k < test_parameter.vid_addr.Length; k++)
                        {
                            if (k != parameter.select_idx)
                            {
                                vid_addr.Add(test_parameter.vid_addr[k]);
                                vid_low.Add(test_parameter.lo_code[k]);
                                vid_high.Add(test_parameter.hi_code[k]);
                            }
                        }

                        _sheet.Cells[row, parameter.idx + aggressor_col].NumberFormat = "@";
                        _sheet.Cells[row, parameter.idx + aggressor_col] = (parameter.data[parameter.idx] == 1) ? vid_low[parameter.idx].ToString("X") + "<->" + vid_high[parameter.idx].ToString("X") : "0";
                        //WriteEn(data, test_parameter.vid_addr, test_parameter.hi_code, test_parameter.lo_code);
                        for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                        {
                            if (parameter.data[parameter.idx] == 0) break;

                            RTDev.I2C_Write((byte)(test_parameter.slave), vid_addr[parameter.idx], new byte[] { vid_low[parameter.idx] });
                            RTDev.I2C_Write((byte)(test_parameter.slave), vid_addr[parameter.idx], new byte[] { vid_high[parameter.idx] });
                        }
                        break;
                    case 3: // LT
                        _sheet.Cells[row, parameter.idx + aggressor_col].NumberFormat = "@";
                        _sheet.Cells[row, parameter.idx + aggressor_col] = (parameter.data[parameter.idx] == 1) ? parameter.l1[parameter.idx] + " <-> " + parameter.l2[parameter.idx] : "0";
                        // eload over 4CH need to select channel
#if Eload_en
                        if (parameter.data[parameter.idx] != 0)
                            InsControl._eload.DymanicLoad(parameter.sw_en[parameter.idx] + 1, parameter.data_l1[parameter.idx], parameter.data_l2[parameter.idx], 500, 500);
                        else
                            InsControl._eload.LoadOFF(parameter.sw_en[parameter.idx] + 1);
#endif
                        break;
                }

                dont_stop_cnt++;
            } // while loop end
        }

        public ATE_CrossTalk(CrossTalk main)
        {
            delegate_mess = new FinishNotification(MessageNotify);
            updateMain = main;
            p_dont_stop = new ParameterizedThreadStart(Channel_Switch_event);
            dont_stop = new Thread(p_dont_stop);
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

            InsControl._oscilloscope.CHx_Off(1);
            InsControl._oscilloscope.CHx_Off(2);
            InsControl._oscilloscope.CHx_Off(3);
            InsControl._oscilloscope.CHx_Off(4);

            InsControl._oscilloscope.CHx_BWLimitOn(1);
            InsControl._oscilloscope.CHx_BWLimitOn(2);
            InsControl._oscilloscope.CHx_BWLimitOn(3);
            InsControl._oscilloscope.CHx_BWLimitOn(4);
        }

        public void WriteFreq(int sealect, int freq_idx)
        {
            int len = test_parameter.freq_addr.Length;
            List<byte> freq_addr = new List<byte>();
            List<byte> freq_data = new List<byte>();

            for (int i = 0; i < len; i++)
            {
                for (int j = i + 1; j < len; j++)
                {
                    if (test_parameter.freq_addr[i] == test_parameter.freq_addr[j])
                    {
                        freq_addr.Add(test_parameter.freq_addr[i]);
                        freq_data.Add((byte)(test_parameter.freq_data[i][freq_idx < test_parameter.freq_data[i].Count ? freq_idx : test_parameter.freq_data[i].Count - 1]
                            | test_parameter.freq_data[j][freq_idx < test_parameter.freq_data[j].Count ? freq_idx : test_parameter.freq_data[j].Count - 1]));
                        break;
                    }
                    else
                    {
                        if (freq_addr.IndexOf(test_parameter.freq_addr[i]) == -1)
                        {
                            freq_addr.Add(test_parameter.freq_addr[i]);
                            freq_data.Add((byte)(test_parameter.freq_data[i][freq_idx < test_parameter.freq_data[i].Count ? freq_idx : test_parameter.freq_data[i].Count - 1]));
                        }
                    }
                }
                if (i == len - 1 && freq_addr.IndexOf(test_parameter.freq_addr[i]) == -1)
                {
                    freq_addr.Add(test_parameter.freq_addr[i]);
                    freq_data.Add((byte)(test_parameter.freq_data[i][freq_idx < test_parameter.freq_data[i].Count ? freq_idx : test_parameter.freq_data[i].Count - 1]));
                }
            }

            freq_addr = freq_addr.Distinct().ToList();


            for (int i = 0; i < freq_addr.Count; i++)
            {
                RTDev.I2C_Write((byte)(test_parameter.slave), freq_addr[i], new byte[] { freq_data[i] });
            }
        }

        public void WriteEn(List<double> data, byte[] addr, byte[] en_on, byte[] dis_off, int select_idx)
        {
            int len = test_parameter.en_addr.Length;
            //List<byte> wr_en = new List<byte>();
            List<byte> en_addr = new List<byte>();

            Dictionary<int, byte> wr_en = new Dictionary<int, byte>();
            Dictionary<int, byte> wr_dis = new Dictionary<int, byte>();

            // find same enable address
            for (int i = 0; i < data.Count; i++)
            {
                for (int j = i + 1; j < data.Count; j++)
                {
                    if (addr[i] == addr[j])
                    {
                        en_addr.Add(addr[i]);

                        // truth tabel state
                        if (data[i] == 1 && data[j] == 1)
                        {
                            // i, j =  1
                            if (wr_en.ContainsKey(addr[i]))
                            {
                                wr_en[addr[i]] |= en_on[j];
                            }
                            else
                            {
                                wr_en.Add(addr[i], (byte)(en_on[i] | en_on[j]));
                            }
                        }
                        else if (data[i] == 1)
                        {
                            // i = 1
                            if (wr_en.ContainsKey(addr[i]))
                            {
                                wr_en[addr[i]] |= (byte)(en_on[i]);
                            }
                            else
                            {
                                wr_en.Add(addr[i], (byte)(en_on[i]));
                            }
                        }
                        else if (data[j] == 1)
                        {
                            // j = 1
                            if (wr_en.ContainsKey(addr[i]))
                            {
                                wr_en[addr[i]] |= (byte)(en_on[j]);
                            }
                            else
                            {
                                wr_en.Add(addr[i], (byte)en_on[j]);
                            }

                        }
                        else
                        {
                            // i, j == 0
                            if (wr_en.ContainsKey(addr[i]))
                            {
                                wr_en[addr[i]] |= (byte)(dis_off[i] | dis_off[j]);
                            }
                            else
                            {
                                wr_en.Add(addr[i], (byte)(dis_off[i] | dis_off[j]));
                            }
                        }

                        break;
                    }
                    else
                    {
                        if (en_addr.IndexOf(addr[i]) == -1)
                        {
                            en_addr.Add(addr[i]);
                            wr_en.Add(addr[i], (byte)(en_on[i]));
                        }
                    }
                }

                if (i == data.Count - 1 && en_addr.IndexOf(addr[i]) == -1)
                {
                    en_addr.Add(addr[i]);
                    wr_en.Add(addr[i], (byte)(en_on[i]));
                }
            }

            en_addr = en_addr.Distinct().ToList();

            byte must_en_addr = test_parameter.en_addr[select_idx];
            byte must_en_data = test_parameter.en_data[select_idx];

            if (en_addr.IndexOf(must_en_addr) == -1)
            {
                en_addr.Add(must_en_addr);
                wr_en.Add(must_en_addr, must_en_data);
            }
            else
            {
                wr_en[must_en_addr] |= must_en_data;
            }

            // channel on off 100 times
            for (int idx = 0; idx < 100; idx++)
            {
                for (int i = 0; i < en_addr.Count; i++)
                {
                    //// turn off all rails
                    for (int j = 0; j < addr.Length; j++)
                    {
                        if (must_en_addr == addr[j])
                            RTDev.I2C_Write((byte)(test_parameter.slave), addr[j], new byte[] { (byte)(dis_off[i] | must_en_data) });
                        else
                            RTDev.I2C_Write((byte)(test_parameter.slave), addr[j], new byte[] { dis_off[i] });
                    }


                    // turn on rails
                    RTDev.I2C_Write((byte)(test_parameter.slave), en_addr[i], new byte[] { wr_en[en_addr[i]] });
                }
            }
        }

        public override void ATETask()
        {
            progress = 0;
            updateMain.UpdateProgressBar(0);
            RTDev.BoadInit();
#if Report_en
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

                        for (int j = 0; j < test_parameter.lt_l1[i].Count; j++)
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

                //rail_info += ",full load=" + test_parameter.full_load[i] + "\r\n";


                for (int j = 0; j < test_parameter.full_load[i].Count; j++)
                {
                    rail_info += test_parameter.full_load[i][j] + ((j == test_parameter.full_load[i].Count - 1) ? "A" : "A, ");
                }

                rail_info += "\r\n";
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

        private void MeasureVictim(int victim, int col_start, double vout, bool before, double vin)
        {
            double vmean = 0;
            double vmax = 0;
            double vmin = 0;
            double jitter = 0;

#if Scope_en
            for (int i = 0; i < 5; i++)
            {
                vmean = InsControl._oscilloscope.CHx_Meas_Mean(victim, 1);
                vmax = InsControl._oscilloscope.CHx_Meas_Max(victim, 2);
                vmin = InsControl._oscilloscope.CHx_Meas_Min(victim, 3);

                vmean = InsControl._oscilloscope.CHx_Meas_Mean(victim, 1);
                vmax = InsControl._oscilloscope.CHx_Meas_Max(victim, 2);
                vmin = InsControl._oscilloscope.CHx_Meas_Min(victim, 3);
                string res = "";
                if (victim <= test_parameter.scope_lx.Count)
                    res = test_parameter.scope_lx[victim - 1];
                if (res.IndexOf("CH") != -1)
                {
                    res = res.Replace("CH", "");
                    int int_res = Convert.ToInt32(res);
                    InsControl._oscilloscope.CHx_Offset(int_res, 0);
                    InsControl._oscilloscope.SetTimeScale(0.00001);

                    vmax = InsControl._oscilloscope.CHx_Meas_Max(int_res, 4);
                    vmax = InsControl._oscilloscope.CHx_Meas_Max(int_res, 4);
                    vmax = InsControl._oscilloscope.CHx_Meas_Max(int_res, 4);

                    InsControl._oscilloscope.CHx_Level(int_res, vin / 4);
                    InsControl._oscilloscope.CHx_Position(int_res, -3);
                    InsControl._oscilloscope.SetNormalTrigger();
                    InsControl._oscilloscope.SetTriggerFall();
                    InsControl._oscilloscope.SetTriggerLevel(int_res, vin * 0.9);

                    double period = 0;
                    period = InsControl._oscilloscope.CHx_Meas_Period(int_res, 4);
                    period = InsControl._oscilloscope.CHx_Meas_Period(int_res, 4);
                    period = InsControl._oscilloscope.CHx_Meas_Period(int_res, 4);
                    period = InsControl._oscilloscope.CHx_Meas_Period(int_res, 4);
                    if (period < Math.Pow(10, 10))
                        InsControl._oscilloscope.SetTimeScale(period);

                    InsControl._oscilloscope.SetDPXOn();
                    MyLib.Delay1ms(300);
                    jitter = InsControl._oscilloscope.CHx_Meas_Jitter(int_res, 4);
                    jitter = InsControl._oscilloscope.CHx_Meas_Jitter(int_res, 4);
                    MyLib.Delay1s(test_parameter.accumulate);
                    jitter = InsControl._oscilloscope.CHx_Meas_Jitter(int_res, 4) * Math.Pow(10, 9);
                    InsControl._oscilloscope.SetDPXOff();
                }
            }

            InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name);
            InsControl._oscilloscope.SetMeasureOff(1);
            InsControl._oscilloscope.SetMeasureOff(2);
            InsControl._oscilloscope.SetMeasureOff(3);
            InsControl._oscilloscope.SetMeasureOff(4);
#endif


#if Report_en
            // for measure victim channel
            //int col_cnt = 7;
            double pos_delta = (vmax - vmean) * 1000;
            double neg_delta = (vmean - vmin) * 1000;
            if (before)
            {
                first_vmean = vmean;
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
                _sheet.Cells[row, col_pos[(int)Col_List.a_delta_pos]] = string.Format("{0:0.000}", (vmax - first_vmean) * 1000); // + delta
                _sheet.Cells[row, col_pos[(int)Col_List.a_delta_neg]] = string.Format("{0:0.000}", (first_vmean - vmin) * 1000); // - delta

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
                                                                    cells[col_pos[(int)Col_List.tol_pos] - 1], row, test_parameter.tolerance);

                _sheet.Cells[row, col_pos[(int)Col_List.res_neg]] = string.Format("=IF({0}{1} < {2}, \"PASS\",\"FAIL\")",
                                                                    cells[col_pos[(int)Col_List.tol_neg] - 1], row, test_parameter.tolerance);


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
#if Scope_en
            OSCInit();
            MyLib.Delay1ms(500);
#endif

            // the select_idx equal to vimtic channel
            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
                    // vin loop
#if Power_en
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
#endif
                    for (int freq_idx = 0; freq_idx < test_parameter.freq_data[select_idx].Count; freq_idx++)
                    {
                        for (int vout_idx = 0; vout_idx < test_parameter.vout_data[select_idx].Count; vout_idx++)
                        {
                            /* change victim vout */
                            for (int i = 0; i < test_parameter.vout_addr.Length; i++)
                            {
                                RTDev.I2C_Write((byte)(test_parameter.slave),
                                                test_parameter.vout_addr[i],
                                                new byte[] { test_parameter.vout_data[i][vout_idx] });
                            }

                            WriteFreq(select_idx, freq_idx);

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
                            for (int full_idx = 0; full_idx < test_parameter.full_load[select_idx].Count; full_idx++)
                            {
                                double victim_iout = test_parameter.full_load[select_idx][full_idx];
                                //double iout = victim_iout;
#if Eload_en
                                InsControl._eload.Loading(select_idx + 1, victim_iout);
#endif

                                // victim current select
                                for (int group_idx = 0; group_idx < cnt_max; group_idx++) // how many iout group
                                {

                                    int col_base = (int)XLS_Table.C + 2 + test_parameter.ch_num;
                                    int col_start = col_base;

#if Report_en
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


                                    _range = _sheet.Cells[row - 2, col_start];
                                    _sheet.Cells[row - 2, col_start] = "Victim Info";
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);

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

                                    _sheet.Cells[row - 1, col_base] = "Before: no load on aggressor";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row)];
                                    _range.Interior.Color = Color.FromArgb(0xCC, 0xFF, 0xEF);
                                    col_base++;

                                    col_pos[(int)Col_List.b_Vmin] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.b_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(ns)";

                                    col_pos[(int)Col_List.b_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.b_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";
                                    //_sheet.Cells[row, col_base++] = "+ Tol (%)";
                                    //_sheet.Cells[row, col_base++] = "- Tol (%)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on aggressor";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 4] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";

                                    col_pos[(int)Col_List.a_Vmax] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";

                                    col_pos[(int)Col_List.a_min] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.a_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(ns)";

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
                                        test_parameter.waveform_name = string.Format("{0}_{1}_VIN={2}_Vout={3}_Freq={4}_Iout={5}",
                                                                        file_idx++,
                                                                        test_parameter.rail_name[select_idx],
                                                                        test_parameter.VinList[vin_idx],
                                                                        test_parameter.vout_des[select_idx][vout_idx],
                                                                        test_parameter.freq_des[select_idx][freq_idx],
                                                                        victim_iout
                                                                        );
                                        int n = ch_sw_num == 2 ? 1 :
                                                ch_sw_num == 4 ? 2 :
                                                ch_sw_num == 8 ? 3 :
                                                ch_sw_num == 16 ? 4 :
                                                ch_sw_num == 32 ? 5 :
                                                ch_sw_num == 64 ? 6 : 7;

                                        MeasNParameter measNParameter = new MeasNParameter();

                                        measNParameter.N = n;
                                        measNParameter.select_idx = select_idx;
                                        measNParameter.vout = Convert.ToDouble(test_parameter.vout_des[select_idx][vout_idx]);
                                        measNParameter.iout_n = victim_iout;
                                        measNParameter.col_start = col_start;
                                        measNParameter.lt_mode = false;
                                        measNParameter.before = victim_idx == 0 ? true : false;
                                        measNParameter.vin = test_parameter.VinList[vin_idx];

                                        MeasureN(measNParameter);

                                        if (victim_idx == 0) row = row - ch_sw_num;
                                    } // test 2 without aggressor and with aggressor
                                    row += 3;
                                } // change full load conditions


                            } // iout group loop
                        } // vout loop
                    } // freq loop
                } // vin loop
            } // select aggressor loop


            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
#if Report_en
            string conditions = (string)_sheet.Cells[2, XLS_Table.B].Value + "\r\n";
            conditions = conditions + time;
            _sheet.Cells[2, XLS_Table.B] = conditions;

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

            OSCInit();
            MyLib.Delay1ms(500);

            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
                    for (int freq_idx = 0; freq_idx < test_parameter.freq_data[select_idx].Count; freq_idx++)
                    {
                        for (int vout_idx = 0; vout_idx < test_parameter.vout_data[select_idx].Count; vout_idx++)
                        {
                            /* change victim vout */
                            for (int i = 0; i < test_parameter.vout_addr.Length; i++)
                            {
                                RTDev.I2C_Write((byte)(test_parameter.slave),
                                                test_parameter.vout_addr[i],
                                                new byte[] { test_parameter.vout_data[i][vout_idx] });
                            }

                            WriteFreq(select_idx, freq_idx);

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


                            for (int full_idx = 0; full_idx < test_parameter.full_load[select_idx].Count; full_idx++)
                            {
                                double victim_iout = test_parameter.full_load[select_idx][full_idx];
                                //double iout = victim_iout;
#if Eload_en
                                InsControl._eload.Loading(select_idx + 1, victim_iout);
#endif

                                for (int group_idx = 0; group_idx < cnt_max; group_idx++) // how many iout group
                                {
                                    //double victim_iout = test_parameter.full_load[select_idx];
                                    int col_base = (int)XLS_Table.C + 2 + test_parameter.ch_num;
                                    int col_start = col_base;

#if Report_en
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

                                    _range = _sheet.Cells[row - 2, col_start];
                                    _sheet.Cells[row - 2, col_start] = "Victim Info";
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range.Interior.Color = Color.FromArgb(0xFF, 0xFF, 0xCC);

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

                                    _sheet.Cells[row - 1, col_base] = "Before: no load on aggressor";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row)];
                                    _range.Interior.Color = Color.FromArgb(0xCC, 0xFF, 0xEF);
                                    col_base++;

                                    col_pos[(int)Col_List.b_Vmin] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.b_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(ns)";

                                    col_pos[(int)Col_List.b_delta_pos] = col_base;
                                    _sheet.Cells[row, col_base++] = "+VΔ (mV)";

                                    col_pos[(int)Col_List.b_delta_neg] = col_base;
                                    _sheet.Cells[row, col_base++] = "-VΔ (mV)";

                                    _sheet.Cells[row - 1, col_base] = "After: with load on aggressor";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 4] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";

                                    col_pos[(int)Col_List.a_Vmax] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";

                                    col_pos[(int)Col_List.a_min] = col_base;
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";

                                    col_pos[(int)Col_List.a_jitter] = col_base;
                                    _sheet.Cells[row, col_base++] = "Jitter(ns)";

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
                                        //double iout = victim_idx == 0 ? 0 : victim_iout;
                                        test_parameter.waveform_name = string.Format("{0}_{1}_VIN={2}_Vout={3}_Freq={4}_Iout={5}",
                                                                file_idx++,
                                                                test_parameter.rail_name[select_idx],
                                                                test_parameter.VinList[vin_idx],
                                                                test_parameter.vout_des[select_idx][vout_idx],
                                                                test_parameter.freq_des[select_idx][freq_idx],
                                                                victim_iout
                                                                );


                                        int n = ch_sw_num == 2 ? 1 :
                                        ch_sw_num == 4 ? 2 :
                                        ch_sw_num == 8 ? 3 :
                                        ch_sw_num == 16 ? 4 :
                                        ch_sw_num == 32 ? 5 :
                                        ch_sw_num == 64 ? 6 : 7;

                                        MeasNParameter measNParameter = new MeasNParameter();

                                        measNParameter.N = n;
                                        measNParameter.select_idx = select_idx;
                                        measNParameter.vout = Convert.ToDouble(test_parameter.vout_des[select_idx][vout_idx]);
                                        measNParameter.iout_n = victim_iout;
                                        measNParameter.col_start = col_start;
                                        measNParameter.lt_mode = true;
                                        measNParameter.before = victim_idx == 0 ? true : false;
                                        measNParameter.vin = test_parameter.VinList[vin_idx];

                                        MeasureN(measNParameter);

                                        if (victim_idx == 0) row = row - ch_sw_num;
                                    } // victim no load and full load
                                    row += 3;
                                }

                            } // group loop
                        } // freq loop
                    } // vout loop
                } // vin loop
            } // channel select

            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
#if Report_en
            string conditions = (string)_sheet.Cells[2, XLS_Table.B].Value + "\r\n";
            conditions = conditions + time;
            _sheet.Cells[2, XLS_Table.B] = conditions;
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
            InsControl._oscilloscope.CHx_On(ch);
            InsControl._oscilloscope.CHx_Offset(ch, vout);
            InsControl._oscilloscope.CHx_Level(ch, 1); // set 100mV
            InsControl._oscilloscope.CHx_Position(ch, 0);

        RE_Scale:
            double vpp = 0;
            for (int i = 0; i < 5; i++)
            {
                vpp = InsControl._oscilloscope.CHx_Meas_VPP(ch, 4);
                MyLib.Delay1ms(200);
                vpp = InsControl._oscilloscope.CHx_Meas_VPP(ch, 4);
                vpp = InsControl._oscilloscope.CHx_Meas_VPP(ch, 4);
                InsControl._oscilloscope.CHx_Level(ch, vpp / 2);
            }

            double res = InsControl._oscilloscope.doQueryNumber(string.Format("CH{0}:SCAle?", ch));
            if (res >= 0.1)
                goto RE_Scale;
        }

        private void MeasureN(MeasNParameter measNParameter)
        {
            int idx = 0;
            int[] sw_en = new int[measNParameter.N]; // save victim channel number
            double[] iout = new double[measNParameter.N];
            double[] l1 = new double[measNParameter.N];
            double[] l2 = new double[measNParameter.N];
            int loop_cnt = (int)Math.Pow(2, measNParameter.N);

            // turn vout channel
            string name = test_parameter.scope_chx[measNParameter.select_idx];
            string res = test_parameter.scope_lx[measNParameter.select_idx];
#if Scope_en
            switch (name)
            {
                case "CH1": CHx_LevelReScale(1, measNParameter.vout); break;
                case "CH2": CHx_LevelReScale(2, measNParameter.vout); break;
                case "CH3": CHx_LevelReScale(3, measNParameter.vout); break;
                case "CH4": CHx_LevelReScale(4, measNParameter.vout); break;
            }

            // enable lx channel
            switch (res)
            {
                case "CH1": InsControl._oscilloscope.CHx_On(1); break;
                case "CH2": InsControl._oscilloscope.CHx_On(2); break;
                case "CH3": InsControl._oscilloscope.CHx_On(3); break;
                case "CH4": InsControl._oscilloscope.CHx_On(4); break;
            }
#endif

            for (int aggressor = 0; aggressor < test_parameter.scope_chx.Count; aggressor++)
            {
                if (test_parameter.eload_chx[aggressor] != test_parameter.eload_chx[measNParameter.select_idx])
                {
                    sw_en[idx++] = test_parameter.eload_chx[aggressor] - 1;
                }
            }
#if Scope_en
            InsControl._oscilloscope.SetClear();
            InsControl._oscilloscope.SetPERSistence();
#endif

            // save aggressor iout conditions
            // iout select maximum setting if over iout list overflow.
            for (int i = 0; i < measNParameter.N; i++)
            {
                if (measNParameter.lt_mode)
                {
                    l1[i] = measNParameter.group < test_parameter.lt_l1[sw_en[i]].Count ?
                        test_parameter.lt_l1[sw_en[i]][measNParameter.group] : test_parameter.lt_l1[sw_en[i]].Max();

                    l2[i] = measNParameter.group < test_parameter.lt_l2[sw_en[i]].Count ?
                        test_parameter.lt_l2[sw_en[i]][measNParameter.group] : test_parameter.lt_l2[sw_en[i]].Max();
                }
                else
                {
                    iout[i] = measNParameter.group < test_parameter.ccm_eload[sw_en[i]].Count ?
                        test_parameter.ccm_eload[sw_en[i]][measNParameter.group] : test_parameter.ccm_eload[sw_en[i]].Max();
                }
            }

            // calculate and excute all of test conditions.
            for (int i = 0; i < loop_cnt; i++)
            {
#if Scope_en
                InsControl._oscilloscope.SetClear();
#endif
                updateMain.UpdateProgressBar(++progress);

#if Scope_en
                InsControl._oscilloscope.SetAutoTrigger();
#endif

#if Eload_en
                if (measNParameter.iout_n != 0)
                    InsControl._eload.Loading(test_parameter.eload_chx[measNParameter.select_idx], measNParameter.iout_n);
                else
                    InsControl._eload.LoadOFF(test_parameter.eload_chx[measNParameter.select_idx]);
#endif

                //TODO: Issue1
                MyLib.Delay1s(1);

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
                for (int j = 0; j < measNParameter.N; j++)
                {
                    switch (test_parameter.cross_mode)
                    {
                        case 0:
                            data.Add(bit_list[j] == 0 ? 0 : iout[j]);
                            break;
                        case 1:
                        case 2:
                            data.Add(bit_list[j] == 0 ? 0 : 1);
                            break;
                        case 3:
                            data.Add(bit_list[j] == 0 ? 0 : 1);
                            data_l1.Add(bit_list[j] == 0 ? 0 : l1[j]);
                            data_l2.Add(bit_list[j] == 0 ? 0 : l2[j]);
                            break;
                    }
                }

                if (!measNParameter.before)
                {
                    for (int j = 0; j < measNParameter.N; j++) // run each channel
                    {
                        dont_stop_cnt = 0;
                        CrossTalkParameter input = new CrossTalkParameter();
                        input.idx = j;
                        input.select_idx = measNParameter.select_idx;
                        input.data = data;
                        input.sw_en = sw_en;
                        input.iout = iout;
                        input.data_l1 = data_l1;
                        input.data_l2 = data_l2;
                        input.l1 = l1;
                        input.l2 = l2;
#if Scope_en
                        InsControl._oscilloscope.CHx_On(measNParameter.select_idx + 1);
#endif
                        dont_stop = new Thread(p_dont_stop);
                        dont_stop.Start(input);
                        MyLib.Delay1s(test_parameter.accumulate);

                        //while (dont_stop_cnt <= 100) ;

                        dont_stop.Abort();
                        dont_stop = null;
                    }
                }


                string temp = test_parameter.waveform_name;
                test_parameter.waveform_name = test_parameter.waveform_name + string.Format("_case{0}", i);
                MeasureVictim(
                                Convert.ToInt32(name.Replace("CH", "")),
                                measNParameter.col_start + 1,
                                measNParameter.vout,
                                measNParameter.before,
                                measNParameter.vin);
                test_parameter.waveform_name = temp;

#if Eload_en
                InsControl._eload.Loading(test_parameter.eload_chx[measNParameter.select_idx], measNParameter.iout_n);
#if Report_en
                _sheet.Cells[row, measNParameter.before ? measNParameter.col_start : measNParameter.col_start + 7] = InsControl._eload.GetIout();
#endif
#endif

#if Eload_en
                InsControl._eload.AllChannel_LoadOff();
#endif
                row++;
            }
#if Scope_en
            InsControl._oscilloscope.CHx_Off(1);
            InsControl._oscilloscope.CHx_Off(2);
            InsControl._oscilloscope.CHx_Off(3);
            InsControl._oscilloscope.CHx_Off(4);
#endif
        }

#endregion
    }

    public class CrossTalkParameter
    {
        public int idx;
        public List<double> data = new List<double>();
        public List<double> data_l1 = new List<double>();
        public List<double> data_l2 = new List<double>();
        public int[] sw_en;
        public double[] iout;
        public int select_idx;
        public double[] l1;
        public double[] l2;

        public CrossTalkParameter()
        { }
    }

    public class MeasNParameter
    {
        public int N;
        public int select_idx;
        public double vout;
        public int group;
        public double iout_n;
        public int col_start;
        public bool before;
        public bool lt_mode = false;
        public double vin;

        public MeasNParameter()
        { }
    }

}





