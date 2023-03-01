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
            for (int select_idx = 0; select_idx < test_parameter.cross_select.Length; select_idx++)
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


                                for (int group_idx = 0; group_idx < switch_max; group_idx++) // how many iout group
                                {

                                    double victim_iout = group_idx < test_parameter.ccm_eload[select_idx].Count() ?
                                                        test_parameter.ccm_eload[select_idx][group_idx] : test_parameter.ccm_eload[select_idx].Max();


#if true
                                    _sheet.Cells[row++, XLS_Table.M] = "No Load Vout";
                                    _range = _sheet.Range["M" + (row - 1).ToString()];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, XLS_Table.M] = "Full Load Vout";
                                    _range = _sheet.Range["M" + (row).ToString()];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row++, XLS_Table.B] = "Vin= " + test_parameter.VinList[vin_idx] + "V";
                                    _range = _sheet.Range["B" + (row - 1).ToString()];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row, XLS_Table.M] = "Freq";
                                    _range = _sheet.Range["M" + (row).ToString()];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    _sheet.Cells[row++, XLS_Table.B] = "Aggressor";
                                    _range = _sheet.Range["B" + (row - 1).ToString()];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                    switch (ch_sw_num)
                                    {
                                        case 2:
                                            _range = _sheet.Range["B" + row, "B" + (row + 1)];
                                            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                            _sheet.Cells[row, XLS_Table.B] = "Rail1";
                                            _sheet.Cells[row + 1, XLS_Table.B] = "Iout(A)";
                                            break;
                                        case 4:
                                            _range = _sheet.Range["B" + row, "C" + (row + 1)];
                                            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                            _sheet.Cells[row, XLS_Table.B] = "Rail1";
                                            _sheet.Cells[row, XLS_Table.C] = "Rail2";
                                            _sheet.Cells[row + 1, XLS_Table.B] = "Iout(A)";
                                            _sheet.Cells[row + 1, XLS_Table.C] = "Iout(A)";
                                            break;
                                        case 8:
                                            _range = _sheet.Range["B" + row, "D" + (row + 1)];
                                            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                            _sheet.Cells[row, XLS_Table.B] = "Rail1";
                                            _sheet.Cells[row, XLS_Table.C] = "Rail2";
                                            _sheet.Cells[row, XLS_Table.D] = "Rail3";
                                            _sheet.Cells[row + 1, XLS_Table.B] = "Iout(A)";
                                            _sheet.Cells[row + 1, XLS_Table.C] = "Iout(A)";
                                            _sheet.Cells[row + 1, XLS_Table.D] = "Iout(A)";
                                            break;
                                    }

                                    _sheet.Cells[row, XLS_Table.M] = "Victim";
                                    _range = _sheet.Range["M" + row, "T" + row];
                                    _range.Merge();
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    row++;

                                    _sheet.Cells[row, XLS_Table.M] = "Iout(A)";
                                    _sheet.Cells[row, XLS_Table.N] = "Vmean(V)";
                                    _sheet.Cells[row, XLS_Table.O] = "Victim Max Voltage";
                                    _sheet.Cells[row, XLS_Table.P] = "Victim Min Voltage";
                                    _sheet.Cells[row, XLS_Table.Q] = "Jitter(%)";
                                    _sheet.Cells[row, XLS_Table.R] = "+VΔ (mV)";
                                    _sheet.Cells[row, XLS_Table.S] = "+ Tol (%)";
                                    _sheet.Cells[row, XLS_Table.T] = "+ Tol";
                                    _range = _sheet.Range["M" + row, "T" + row];
                                    _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    row++;

                                    FirstMeasure(select_idx);
                                    _sheet.Cells[row - 3, XLS_Table.N] = test_parameter.freq_des[select_idx][freq_idx];
#endif
                                    for (int victim_idx = 0; victim_idx < 2; victim_idx++)
                                    {

                                        double iout = victim_idx == 0 ? 0 : victim_iout;
                                        //InsControl._eload.Loading(select_idx + 1, iout);

                                        // victim info
                                        //_sheet.Cells[row, XLS_Table.M] = test_parameter.freq_des[select_idx][freq_idx];
                                        //_sheet.Cells[row, XLS_Table.N] = iout;

                                        switch (ch_sw_num)
                                        {
                                            case 2:
                                                Measure2(select_idx, group_idx, iout);
                                                break;
                                            case 4:
                                                Measure4(select_idx, group_idx, iout);
                                                break;
                                            case 8:
                                                Measure8(select_idx, group_idx, iout);
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

        private void MeasureVictim(int victim)
        {
            //=IF(AD23="","",IF(AF23>AD23,"PASS","FAIL"))
            //= IF(AE23 = "", "", IF(AG23 > AE23, "PASS", "FAIL"))
            //_sheet.Cells[row, XLS_Table.Q] = InsControl._oscilloscope.CHx_Meas_Mean(victim);
            //_sheet.Cells[row, XLS_Table.R] = InsControl._oscilloscope.CHx_Meas_Max(victim);
            //_sheet.Cells[row, XLS_Table.S] = InsControl._oscilloscope.CHx_Meas_Min(victim);


            _sheet.Cells[row, XLS_Table.Q] = "Jitter";
            _sheet.Cells[row, XLS_Table.R] = string.Format("=IF(Q{0}=\"\",\"\",(R{1}-Q{2}) * 1000)", row, row, row);
            _sheet.Cells[row, XLS_Table.S] = "+ Tol (%)";
            _sheet.Cells[row, XLS_Table.T] = "+ Tol";
        }

        private void FirstMeasure(int sel)
        {
            // disable others channel and scope control
            double iout_max = test_parameter.ccm_eload[sel].Max();

            for (int ch_idx = 0; ch_idx < 4; ch_idx++)
            {
                if (ch_idx == sel)
                {
                    //InsControl._oscilloscope.CHx_On(ch_idx + 1);
                    RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                        test_parameter.en_addr[ch_idx],
                        new byte[] { test_parameter.en_code[ch_idx] });
                }
                else
                {
                    //InsControl._oscilloscope.CHx_Off(ch_idx + 1);
                    RTDev.I2C_Write((byte)(test_parameter.slave >> 1),
                        test_parameter.en_addr[ch_idx],
                        new byte[] { test_parameter.disable_code[ch_idx] });
                }
            }

            //InsControl._eload.Loading((sel + 1), 0);
            //_sheet.Cells[row - 5, XLS_Table.G] = InsControl._oscilloscope.CHx_Meas_Mean(sel + 1);
            
            //InsControl._eload.Loading((sel + 1), iout_max);
            //_sheet.Cells[row - 4, XLS_Table.G] = InsControl._oscilloscope.CHx_Meas_Mean(sel + 1);

            //InsControl._eload.Loading((sel + 1), 0);
            //InsControl._eload.AllChannel_LoadOff();
        }

        private void Measure2(int aggressor, int group, double iout)
        {
            int sw_ch = 0;
            for (int victim = 0; victim < test_parameter.cross_en.Length; victim++)
            {
                if (aggressor != victim && test_parameter.cross_en[victim]) sw_ch = victim + 1;
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
                MeasureVictim(aggressor);
                _sheet.Cells[row, XLS_Table.M] = iout;
                row++;
            }
        }

        private void Measure4(int aggressor, int group, double iout_n)
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
                MeasureVictim(aggressor);
                _sheet.Cells[row, XLS_Table.M] = iout_n;
                row++;
            }
        }

        private void Measure8(int aggressor, int group, double iout_n)
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
                MeasureVictim(aggressor);
                _sheet.Cells[row, XLS_Table.M] = iout_n;
                row++;
            }
        }

    }
}
