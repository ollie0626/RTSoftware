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
            InsControl._oscilloscope.SetTimeScale(2 / 1000);
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
#endif

            // first item CCM
            Cross_CCM();
        }

        private void MeasureVictim(int victim, int col_start, bool before)
        {
            // for measure victim channel
            if (before)
            {
                _sheet.Cells[row, col_start++] = "before Vmean(V)";
                _sheet.Cells[row, col_start++] = "before Victim Max Voltage";
                _sheet.Cells[row, col_start++] = "before Victim Min Voltage";
                _sheet.Cells[row, col_start++] = "before Jitter(%)";
                _sheet.Cells[row, col_start++] = "before +VΔ (mV)";
            }
            else
            {
                _sheet.Cells[row, col_start++ + 6] = "after Victim Max Voltage";
                _sheet.Cells[row, col_start++ + 6] = "after Victim Min Voltage";
                _sheet.Cells[row, col_start++ + 6] = "after Jitter(%)";
                _sheet.Cells[row, col_start + 6] = "+VΔ (mV)";
            }
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
            //OSCInit();
            MyLib.Delay1ms(500);

            // the select_idx equal to aggressor channel
            for (int select_idx = 0; select_idx < test_parameter.cross_en.Length; select_idx++)
            {
                if (test_parameter.cross_en[select_idx]) // select equal to aggressor
                {
                    for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                    {
                        // vin loop
                        //InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);

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

                                // victim current select
                                for (int group_idx = 0; group_idx < switch_max; group_idx++) // how many iout group
                                {

                                    //double victim_iout = group_idx < test_parameter.ccm_eload[select_idx].Count() ?
                                    //                    test_parameter.ccm_eload[select_idx][group_idx] : test_parameter.ccm_eload[select_idx].Max();

                                    double victim_iout = test_parameter.full_load[select_idx];
#if true
                                    int col_base = (int)XLS_Table.B + 2 + test_parameter.ch_num;
                                    int col_start = col_base;
                                    _sheet.Cells[row, col_start] = "Vout=" + test_parameter.vout_des[select_idx][vout_idx];
                                    _sheet.Cells[row++, XLS_Table.B] = "Vin=" + test_parameter.VinList[vin_idx] + "V";
                                    _sheet.Cells[row, XLS_Table.B] = "Aggressor";
                                    _range = _sheet.Range["B" + (row - 1).ToString(), "B" + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    _sheet.Cells[row, col_start] = "Freq (KHz)";
                                    row++;
                                    int col_idx = (int)XLS_Table.B;

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

                                    _sheet.Cells[row - 1, col_base] = "After: with load on victim";
                                    _range = _sheet.Range[cells[col_base - 1] + (row - 1), cells[col_base + 3] + (row - 1)];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, col_base++] = test_parameter.rail_name[select_idx] + "(A)";
                                    _sheet.Cells[row, col_base++] = "Victim Max Voltage";
                                    _sheet.Cells[row, col_base++] = "Victim Min Voltage";
                                    _sheet.Cells[row, col_base++] = "Jitter(%)";
                                    _sheet.Cells[row, col_base] = "+VΔ (mV)";

                                    _range = _sheet.Range[cells[(int)XLS_Table.B + 2 + test_parameter.ch_num - 1] + (row - 1), cells[col_base - 1] + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    for (int i = 1; i < 25; i++)
                                        _sheet.Columns[i].AutoFit();
                                    row++;

                                    //FirstMeasure(select_idx);
                                    _sheet.Cells[row - 2, col_start + 1] = test_parameter.freq_des[select_idx][freq_idx];
#endif
                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {

                                        double iout = (victim_idx == 0) ? 0 : victim_iout;

                                        switch (ch_sw_num)
                                        {
                                            case 2:
                                                //Measure2(select_idx, group_idx, iout);
                                                break;
                                            case 4:
                                                Measure4(select_idx, group_idx, iout, col_start, victim_idx == 0 ? true : false);
                                                if (victim_idx == 0) row = row - 4;
                                                break;
                                            case 8:
                                                Measure8(select_idx, group_idx, iout, col_start, victim_idx == 0 ? true : false);
                                                if (victim_idx == 0) row = row - 8;
                                                break;
                                            case 16: // 4
                                            case 32: // 5
                                            case 64: // 6
                                            case 128: // 7
                                                int n = ch_sw_num == 16 ? 4 :
                                                        ch_sw_num == 32 ? 5 :
                                                        ch_sw_num == 64 ? 6 : 7;
                                                MeasureN(n, select_idx, group_idx, iout, col_start, victim_idx == 0 ? true : false);
                                                if (victim_idx == 0) row = row - ch_sw_num;
                                                break;
                                        }
                                    }

                                    row += 3;
                                } // iout group loop


                            } // vout loop
                        } // freq loop
                    } // vin loop
                } // select aggressor loop

            }

            stopWatch.Stop();
        }

        private void ReserveMeasureChannel(int victim)
        {
            for (int ch_idx = 0; ch_idx < 4; ch_idx++)
            {
                if (ch_idx == victim)
                    InsControl._oscilloscope.CHx_On(ch_idx + 1);
                else
                    InsControl._oscilloscope.CHx_Off(ch_idx + 1);
            }
        }

        private void Measure2(int aggressor, int group, double iout)
        {
            int sw_ch = 0;
            for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
            {
                if (aggressor != victim && test_parameter.cross_en[victim]) sw_ch = victim;
            }

            //ReserveMeasureChannel(aggressor);

            for (int idx = 0; idx < 2; idx++)
            {
                switch (idx)
                {
                    case 0:
                        //InsControl._eload.Loading(sw_ch, 0);
                        _sheet.Cells[row, XLS_Table.B] = 0;
                        break;
                    case 1:
                        //InsControl._eload.Loading(sw_ch, test_parameter.ccm_eload[sw_ch][group]);
                        _sheet.Cells[row, XLS_Table.B] = test_parameter.ccm_eload[sw_ch][group];
                        break;
                }
                //MeasureVictim(aggressor);
                _sheet.Cells[row, XLS_Table.M] = iout;
                row++;
            }
        }

        private void Measure4(int aggressor, int group, double iout_n, int col_start, bool before)
        {
            // program flow
            // find enable channel
            // get group iout setting
            int idx = 0;
            int[] sw_en = new int[2];
            double[] iout = new double[2];
            for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
            {
                if (victim != aggressor && test_parameter.cross_en[victim])
                {
                    sw_en[idx++] = victim;
                }
            }

            //ReserveMeasureChannel(aggressor);

            // if group setting is differenct. I need found max current setting
            iout[0] = group < test_parameter.ccm_eload[sw_en[0]].Count ?
                test_parameter.ccm_eload[sw_en[0]][group] : test_parameter.ccm_eload[sw_en[0]].Max();

            iout[1] = group < test_parameter.ccm_eload[sw_en[1]].Count ?
                test_parameter.ccm_eload[sw_en[1]][group] : test_parameter.ccm_eload[sw_en[1]].Max();

            for (idx = 0; idx < 4; idx++)
            {
                switch (idx)
                {
                    case 0:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);
                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        break;
                    case 1:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        break;
                    case 2:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);

                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        break;
                    case 3:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        break;
                }
                MeasureVictim(aggressor, col_start + 1, before);
                _sheet.Cells[row, before ? col_start : col_start + 6] = iout_n;
                row++;
            }
        }

        private void Measure8(int aggressor, int group, double iout_n, int col_start, bool before)
        {
            int idx = 0;
            int[] sw_en = new int[3];
            double[] iout = new double[3];
            for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
            {
                if (victim != aggressor && test_parameter.cross_en[victim])
                {
                    sw_en[idx++] = victim;
                }
            }

            //ReserveMeasureChannel(aggressor);

            iout[0] = group < test_parameter.ccm_eload[sw_en[0]].Count ?
                test_parameter.ccm_eload[sw_en[0]][group] : test_parameter.ccm_eload[sw_en[0]].Max();

            iout[1] = group < test_parameter.ccm_eload[sw_en[1]].Count ?
                test_parameter.ccm_eload[sw_en[1]][group] : test_parameter.ccm_eload[sw_en[1]].Max();

            iout[2] = group < test_parameter.ccm_eload[sw_en[2]].Count ?
                test_parameter.ccm_eload[sw_en[2]][group] : test_parameter.ccm_eload[sw_en[2]].Max();

            for (idx = 0; idx < 8; idx++)
            {
                switch (idx)
                {
                    case 0:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);
                        //InsControl._eload.Loading(sw_en[2] + 1, 0);

                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        _sheet.Cells[row, XLS_Table.D] = 0;
                        break;
                    case 1:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);
                        //InsControl._eload.Loading(sw_en[2] + 1, 0);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        _sheet.Cells[row, XLS_Table.D] = 0;
                        break;
                    case 2:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
                        //InsControl._eload.Loading(sw_en[2] + 1, 0);

                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        _sheet.Cells[row, XLS_Table.D] = 0;
                        break;
                    case 3:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
                        //InsControl._eload.Loading(sw_en[2] + 1, 0);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        _sheet.Cells[row, XLS_Table.D] = 0;
                        break;
                    case 4:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);
                        //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        _sheet.Cells[row, XLS_Table.D] = iout[2];
                        break;
                    case 5:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, 0);
                        //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = 0;
                        _sheet.Cells[row, XLS_Table.D] = iout[2];
                        break;
                    case 6:
                        //InsControl._eload.Loading(sw_en[0] + 1, 0);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
                        //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

                        _sheet.Cells[row, XLS_Table.B] = 0;
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        _sheet.Cells[row, XLS_Table.D] = iout[2];
                        break;
                    case 7:
                        //InsControl._eload.Loading(sw_en[0] + 1, iout[0]);
                        //InsControl._eload.Loading(sw_en[1] + 1, iout[1]);
                        //InsControl._eload.Loading(sw_en[2] + 1, iout[2]);

                        _sheet.Cells[row, XLS_Table.B] = iout[0];
                        _sheet.Cells[row, XLS_Table.C] = iout[1];
                        _sheet.Cells[row, XLS_Table.D] = iout[2];
                        break;
                }
                MeasureVictim(aggressor, col_start + 1, before);
                _sheet.Cells[row, before ? col_start : col_start + 6] = iout_n;
                row++;
            }
        }

        private void MeasureN(int n, int aggressor, int group, double iout_n, int col_start, bool before)
        {
            int idx = 0;
            int[] sw_en = new int[n];
            double[] iout = new double[n];
            int loop_cnt = (int)Math.Pow(2, n);
            Dictionary<int, List<double>> iout_list = new Dictionary<int, List<double>>();
            for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
            {
                if (victim != aggressor && test_parameter.cross_en[victim])
                {
                    sw_en[idx++] = victim;
                }
            }

            for (int i = 0; i < n; i++)
            {
                iout[i] = group < test_parameter.ccm_eload[sw_en[i]].Count ?
                    test_parameter.ccm_eload[sw_en[i]][group] : test_parameter.ccm_eload[sw_en[i]].Max();
            }

            for (int i = 0; i < loop_cnt; i++)
            {
                List<double> data = new List<double>();
                int bit0 = (i & 0x01) >> 0;
                int bit1 = (i & 0x02) >> 1;
                int bit2 = (i & 0x04) >> 2;
                int bit3 = (i & 0x08) >> 3;
                int bit4 = (i & 0x10) >> 4;
                int bit5 = (i & 0x20) >> 5;
                int bit6 = (i & 0x40) >> 6;
                int bit7 = (i & 0x80) >> 7;
                int[] bit_list = new int[] { bit0, bit1, bit2, bit3, bit4, bit5, bit6, bit7 };

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
                            //data.Add(bit_list[j] == 0 ? 0 : test_parameter.en_data[j]);

                            break;
                    }

                }
                iout_list.Add(i, data);

                int aggressor_col = (int)XLS_Table.B;
                for (int j = 0; j < n; j++)
                {
                    switch (test_parameter.cross_mode)
                    {
                        case 0: // ccm mode
                            //InsControl._eload.Loading(sw_en[j] + 1, iout[j]);
                            _sheet.Cells[row, j + aggressor_col] = data[j];
                            break;
                        case 1: // i2c on / off
                            _sheet.Cells[row, j + aggressor_col] = "Enable -> Disable";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.en_data[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.disen_data[j] });
                            }
                            break;
                        case 2: // i2c VIP
                            _sheet.Cells[row, j + aggressor_col] = "VIP";
                            for (int repeat_idx = 0; repeat_idx < 100; repeat_idx++)
                            {
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.hi_code[j] });
                                RTDev.I2C_Write((byte)(test_parameter.slave >> 1), test_parameter.en_addr[j], new byte[] { test_parameter.lo_code[j] });
                            }
                            break;
                    }
                }

                MeasureVictim(aggressor, col_start + 1, before);
                _sheet.Cells[row, before ? col_start : col_start + 6] = iout_n;
                row++;
            }
        }

        #endregion


        #region "Cross Talk Hi/Lo Code & Channel On/Off" 

        private void Cross_I2C()
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
                            // vout loop
                            for (int freq_idx = 0; freq_idx < test_parameter.freq_data[select_idx].Count; freq_idx++)
                            {
                                for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                {

                                } // victim no load and full load
                            } // freq loop
                        } // vout loop
                    } // vin loop
                } // channel select
            } // select
        }

        #endregion
    }
}