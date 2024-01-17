﻿

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
using System.Text.RegularExpressions;

namespace SoftStartTiming
{
    public class delayTime_parameter
    {
        public double vin;

        // each times test reflash test condition
        public double CH1Lev;
        public double CH2Lev;
        public double CH3Lev;
        public double CH4Lev;

        public int[] meas_posCH1 = new int[2]; // meas channel 0: start, 1: stop
        public int[] meas_posCH2 = new int[2];
        public int[] meas_posCH3 = new int[2];
        public int[] meas_posCH4 = new int[2];

        public double[] precentCH1 = new double[2]; // meas percentage 100% to 0%
        public double[] precentCH2 = new double[2];
        public double[] precentCH3 = new double[2];
        public double[] precentCH4 = new double[2];

        public double idealTime0; // ideal spec
        public double idealTime1;
        public double idealTime2;
        public double idealTime3;

        public int eloadCH1; // eload select???
        public int eloadCH2;
        public int eloadCH3;
        public int eloadCH4;

        public double loading1; // eload loading
        public double loading2;
        public double loading3;
        public double loading4;

        public string seq0; // sequence reg
        public string seq1;
        public string seq2;
        public string seq3;

        public string ideal0; // ideal reg
        public string ideal1;
        public string ideal2;
        public string ideal3;

        public int trigger_ch;
    }

    public class ATE_DelayTime : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        //Excel.Chart _chart;

        //public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        //TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;
        //const int meas_dt1 = 1;
        //const int meas_dt2 = 2;
        //const int meas_dt3 = 3;

        const int meas_sst1 = 1;
        const int meas_sst2 = 2;
        const int meas_sst3 = 3;
        const int meas_sst4 = 4;

        const int meas_vmax1 = 5;
        const int meas_vmax2 = 6;
        const int meas_vmax3 = 7;
        const int meas_vmax4 = 8;

        delayTime_parameter dt_test = new delayTime_parameter();


        public ATE_DelayTime()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }


        private void SetMeasurePercent(int meas_n, double hi, double mid, double lo)
        {
            string cmd = string.Format("MEASUrement:MEAS{0}:REFLevel:METHod PERCent", meas_n);
            InsControl._tek_scope.DoCommand(cmd);

            cmd = string.Format("MEASUrement:MEAS{0}:REFLevel:PERCent:HIGH {1}", meas_n, hi);
            InsControl._tek_scope.DoCommand(cmd);

            cmd = string.Format("MEASUrement:MEAS{0}:REFLevel:PERCent:MID {1}", meas_n, mid);
            InsControl._tek_scope.DoCommand(cmd);

            cmd = string.Format("MEASUrement:MEAS{0}:REFLevel:PERCent:LOW {1}", meas_n, lo);
            InsControl._tek_scope.DoCommand(cmd);
        }

        private void GetParameter(int idx)
        {
            DataGridView seq_dg = test_parameter.seq_dg;
            dt_test.vin = Convert.ToDouble(seq_dg[0, idx].Value);

            #region "scope info"
            // init level
            dt_test.CH1Lev = Convert.ToDouble(seq_dg[5, idx].Value.ToString().Split(',')[0]);
            dt_test.CH2Lev = Convert.ToDouble(seq_dg[5, idx].Value.ToString().Split(',')[1]);
            dt_test.CH3Lev = Convert.ToDouble(seq_dg[5, idx].Value.ToString().Split(',')[2]);
            dt_test.CH4Lev = Convert.ToDouble(seq_dg[5, idx].Value.ToString().Split(',')[3]);

            // get meas channel 
            string[] tmp = seq_dg[2, idx].Value.ToString().Split(',');
            string[] pos = tmp[0].Split('→');
            dt_test.meas_posCH1[0] = Convert.ToInt32(pos[0].Replace("CH", ""));
            dt_test.meas_posCH1[1] = Convert.ToInt32(pos[1].Replace("CH", ""));

            pos = tmp[1].Split('→');
            dt_test.meas_posCH2[0] = Convert.ToInt32(pos[0].Replace("CH", ""));
            dt_test.meas_posCH2[1] = Convert.ToInt32(pos[1].Replace("CH", ""));

            pos = tmp[2].Split('→');
            dt_test.meas_posCH3[0] = Convert.ToInt32(pos[0].Replace("CH", ""));
            dt_test.meas_posCH3[1] = Convert.ToInt32(pos[1].Replace("CH", ""));

            pos = tmp[3].Split('→');
            dt_test.meas_posCH4[0] = Convert.ToInt32(pos[0].Replace("CH", ""));
            dt_test.meas_posCH4[1] = Convert.ToInt32(pos[1].Replace("CH", ""));

            // percentage
            tmp = seq_dg[3, idx].Value.ToString().Split(',');
            string[] per = tmp[0].Split('→');
            dt_test.precentCH1[0] = Convert.ToDouble(per[0]);
            dt_test.precentCH1[1] = Convert.ToDouble(per[1]);

            per = tmp[1].Split('→');
            dt_test.precentCH2[0] = Convert.ToDouble(per[0]);
            dt_test.precentCH2[1] = Convert.ToDouble(per[1]);

            per = tmp[2].Split('→');
            dt_test.precentCH3[0] = Convert.ToDouble(per[0]);
            dt_test.precentCH3[1] = Convert.ToDouble(per[1]);

            per = tmp[3].Split('→');
            dt_test.precentCH4[0] = Convert.ToDouble(per[0]);
            dt_test.precentCH4[1] = Convert.ToDouble(per[1]);


            tmp = seq_dg[7, idx].Value.ToString().Split(',');
            dt_test.idealTime0 = Convert.ToDouble(tmp[0]);
            dt_test.idealTime1 = Convert.ToDouble(tmp[1]);
            dt_test.idealTime2 = Convert.ToDouble(tmp[2]);
            dt_test.idealTime3 = Convert.ToDouble(tmp[3]);

            // example: CH1[3.15]
            tmp = seq_dg[6, idx].Value.ToString().Split(',');
            string input = tmp[0];
            string pattern1 = @"CH(\d+)";
            string pattern2 = @"\[(\d+(\.\d+)?)\]";
            Match match1 = Regex.Match(input, pattern1);
            Match match2 = Regex.Match(input, pattern2);

            //match.Success
            dt_test.eloadCH1 = Convert.ToInt32(match1.Groups[1].Value);
            dt_test.loading1 = Convert.ToDouble(match2.Groups[1].Value);

            input = tmp[1];
            match1 = Regex.Match(input, pattern1);
            match2 = Regex.Match(input, pattern2);
            dt_test.eloadCH2 = Convert.ToInt32(match1.Groups[1].Value);
            dt_test.loading2 = Convert.ToDouble(match2.Groups[1].Value);

            input = tmp[2];
            match1 = Regex.Match(input, pattern1);
            match2 = Regex.Match(input, pattern2);
            dt_test.eloadCH3 = Convert.ToInt32(match1.Groups[1].Value);
            dt_test.loading3 = Convert.ToDouble(match2.Groups[1].Value);

            input = tmp[3];
            match1 = Regex.Match(input, pattern1);
            match2 = Regex.Match(input, pattern2);
            dt_test.eloadCH4 = Convert.ToInt32(match1.Groups[1].Value);
            dt_test.loading4 = Convert.ToDouble(match2.Groups[1].Value);
            #endregion

            tmp = seq_dg[1, idx].Value.ToString().Split(',');
            dt_test.seq0 = tmp[0];
            dt_test.seq1 = tmp[1];
            dt_test.seq2 = tmp[2];
            dt_test.seq3 = tmp[3];

            tmp = seq_dg[4, idx].Value.ToString().Split(',');
            dt_test.ideal0 = tmp[0];
            dt_test.ideal1 = tmp[1];
            dt_test.ideal2 = tmp[2];
            dt_test.ideal3 = tmp[3];

            string trigger_ch = seq_dg[8, idx].Value.ToString();
            dt_test.trigger_ch = Convert.ToInt32(trigger_ch.Replace("CH", ""));
        }

        private void GetSameAddr(ref Dictionary<int, int> map, int addr, int data)
        {
            if (map.ContainsKey(addr))
            {
                map[addr] |= data;
            }
            else
            {
                map.Add(addr, data);
            }
        }

        private void SeqAndIdealWrite()
        {
            string pattern = @"[A-Za-z0-9][A-Za-z0-9]";

            string[] seqTable = new string[] { dt_test.seq0, dt_test.seq1, dt_test.seq2, dt_test.seq3 };
            string[] idealTable = new string[] { dt_test.ideal0, dt_test.ideal1, dt_test.ideal2, dt_test.ideal3 };
            Dictionary<int, int> addr_map = new Dictionary<int, int>();
            List<int> addrList = new List<int>();

            for (int i = 0; i < 4; i++)
            {
                string input = seqTable[i];
                MatchCollection match = Regex.Matches(input, pattern);
                int seq_addr = Convert.ToInt32(match[0].ToString(), 16);
                int seq_data = Convert.ToInt32(match[1].ToString(), 16);
                GetSameAddr(ref addr_map, seq_addr, seq_data);
                input = idealTable[i];
                match = Regex.Matches(input, pattern);
                int ideal_addr = Convert.ToInt32(match[0].ToString(), 16);
                int ideal_data = Convert.ToInt32(match[1].ToString(), 16);
                GetSameAddr(ref addr_map, ideal_addr, ideal_data);
                addrList.Add(seq_addr);
                addrList.Add(ideal_addr);
            }

            addrList = addrList.Distinct().ToList();

            for (int i = 0; i < addr_map.Count; i++)
            {
                int addr = addrList[i];
                int data = addr_map[addr];
                byte[] buf = new byte[] { (byte)data };
                RTDev.I2C_Write((byte)test_parameter.slave, (byte)addr, buf);
            }
        }

        private void OSCInit()
        {
            InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
            InsControl._tek_scope.SetTimeBasePosition(15);
            InsControl._tek_scope.SetRun();
            InsControl._tek_scope.SetTriggerMode(); // auto trigger
            InsControl._tek_scope.SetTriggerSource(dt_test.trigger_ch);

            InsControl._tek_scope.CHx_On(1);
            InsControl._tek_scope.CHx_On(2);
            InsControl._tek_scope.CHx_On(3);
            InsControl._tek_scope.CHx_On(4);

            InsControl._tek_scope.CHx_BWlimitOn(1);
            InsControl._tek_scope.CHx_BWlimitOn(2);
            InsControl._tek_scope.CHx_BWlimitOn(3);
            InsControl._tek_scope.CHx_BWlimitOn(4);

            InsControl._tek_scope.CHx_Position(1, -1);
            InsControl._tek_scope.CHx_Position(2, -2.5);
            InsControl._tek_scope.CHx_Position(3, -3);
            InsControl._tek_scope.CHx_Position(4, -3.5);

            //SetMeasurePercent(meas_scope_ch1, 100, 50, 0);

            /* initial level setting */
            InsControl._tek_scope.CHx_Level(1, dt_test.CH1Lev);
            InsControl._tek_scope.CHx_Level(2, dt_test.CH2Lev);
            InsControl._tek_scope.CHx_Level(3, dt_test.CH3Lev);
            InsControl._tek_scope.CHx_Level(4, dt_test.CH4Lev);

            //InsControl._tek_scope.SetMeasureSource(1, meas_sst1, "RISe");
            if (test_parameter.sleep_mode) InsControl._tek_scope.SetMeasureSource(1, meas_sst1, "RISe");
            else InsControl._tek_scope.SetMeasureSource(1, meas_sst1, "FALL");
            InsControl._tek_scope.SetMeasureSource(2, meas_sst2, "RISe");
            InsControl._tek_scope.SetMeasureSource(3, meas_sst3, "RISe");
            InsControl._tek_scope.SetMeasureSource(4, meas_sst4, "RISe");


            double hi = dt_test.precentCH1[0] > dt_test.precentCH1[1] ? dt_test.precentCH1[0] : dt_test.precentCH1[1];
            double lo = dt_test.precentCH1[0] < dt_test.precentCH1[1] ? dt_test.precentCH1[0] : dt_test.precentCH1[1];
            SetMeasurePercent(meas_sst1, hi, hi * 0.5, lo);

            hi = dt_test.precentCH2[0] > dt_test.precentCH2[1] ? dt_test.precentCH2[0] : dt_test.precentCH2[1];
            lo = dt_test.precentCH2[0] < dt_test.precentCH2[1] ? dt_test.precentCH2[0] : dt_test.precentCH2[1];
            SetMeasurePercent(meas_sst2, hi, hi * 0.5, lo);

            hi = dt_test.precentCH3[0] > dt_test.precentCH3[1] ? dt_test.precentCH3[0] : dt_test.precentCH3[1];
            lo = dt_test.precentCH3[0] < dt_test.precentCH3[1] ? dt_test.precentCH3[0] : dt_test.precentCH3[1];
            SetMeasurePercent(meas_sst3, hi, hi * 0.5, lo);


            hi = dt_test.precentCH4[0] > dt_test.precentCH4[1] ? dt_test.precentCH4[0] : dt_test.precentCH4[1];
            lo = dt_test.precentCH4[0] < dt_test.precentCH4[1] ? dt_test.precentCH4[0] : dt_test.precentCH4[1];
            SetMeasurePercent(meas_sst4, hi, hi * 0.5, lo);

            InsControl._tek_scope.DoCommand("HORizontal:ROLL OFF");
            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            InsControl._tek_scope.PersistenceDisable();
        }

        private void TriggerEvent(double vin)
        {
            switch (test_parameter.trigger_event)
            {
                case 0: // gpio trigger
                    InsControl._tek_scope.SetTriggerSource(1);
                    InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                    InsControl._tek_scope.CHx_Position(1, 1.5);

                    if (test_parameter.sleep_mode)
                        GpioOnSelect(test_parameter.gpio_pin);
                    else
                        GpioOffSelect(test_parameter.gpio_pin);
                    break;
                case 1: // i2c trigger
                    InsControl._tek_scope.SetTriggerSource(1);
                    InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                    InsControl._tek_scope.CHx_Position(1, 1.5);

                    // rails enable
                    I2C_DG_Write(test_parameter.i2c_init_dg);
                    MyLib.Delay1ms(50);
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                    break;
                case 2: // vin trigger
#if Power_en
                    InsControl._power.AutoSelPowerOn(vin);
#endif
                    InsControl._tek_scope.SetTriggerSource(1);
                    InsControl._tek_scope.SetTriggerLevel(vin * 0.35);
                    break;
                case 3: // rail trigger
                    I2C_DG_Write(test_parameter.i2c_init_dg);
                    MyLib.Delay1ms(50);
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                    break;
            }
        }

        private void LevelEvent()
        {
            InsControl._tek_scope.SetMeasureSource(1, meas_vmax1, "MAXimum");

            InsControl._tek_scope.CHx_Level(1, dt_test.CH1Lev);
            InsControl._tek_scope.CHx_Level(2, dt_test.CH2Lev);
            InsControl._tek_scope.CHx_Level(3, dt_test.CH3Lev);
            InsControl._tek_scope.CHx_Level(4, dt_test.CH4Lev);

            InsControl._tek_scope.SetMeasureSource(1, meas_vmax1, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(2, meas_vmax2, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(3, meas_vmax3, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(4, meas_vmax4, "MAXimum");

            int re_cnt = 0;
            for (int ch_idx = 0; ch_idx < 4; ch_idx++)
            {
                if (test_parameter.auto_en[ch_idx])
                {
                re_scale:;
                    if (re_cnt > 3)
                    {
                        re_cnt = 0;
                        continue;
                    }

                    double vmax = 0;
                    for (int k = 0; k < 3; k++)
                    {
                        vmax = InsControl._tek_scope.MeasureMax(meas_vmax1 + ch_idx);
                        vmax = InsControl._tek_scope.MeasureMax(meas_vmax1 + ch_idx);
                        MyLib.Delay1ms(50);
                        vmax = InsControl._tek_scope.MeasureMax(meas_vmax1 + ch_idx);

                        if (vmax > 0.3 && vmax < Math.Pow(10, 3))
                        {
                            InsControl._tek_scope.CHx_Level(ch_idx + 1, vmax / 3);
                            if(ch_idx == 0)
                            {
                                InsControl._tek_scope.SetTriggerLevel((vmax / 3) * 0.8);
                            }
                        }
                        MyLib.Delay1ms(300);
                    }
                    MyLib.Delay1ms(300);
                }
            }

            InsControl._tek_scope.SetMeasureSource(1, meas_vmax1, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(2, meas_vmax2, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(3, meas_vmax3, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(4, meas_vmax4, "MAXimum");
        }

        private void Scope_Channel_Resize(double vin)
        {
            double time_scale = 0;
            InsControl._tek_scope.SetRun();
            InsControl._tek_scope.SetTriggerMode();
            time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
            if (time_scale <= 55 * Math.Pow(10, -6) || time_scale > 100 * Math.Pow(10, -3))
            {
                time_scale = test_parameter.ontime_scale_ms / 1000;
            }
            InsControl._tek_scope.SetTimeScale((25 * Math.Pow(10, -12)));
            InsControl._tek_scope.SetRun();
            InsControl._tek_scope.SetTriggerMode();
#if Power_en
            InsControl._power.AutoSelPowerOn(vin);
            MyLib.Delay1ms(1000);
            I2C_DG_Write(test_parameter.i2c_init_dg);
            SeqAndIdealWrite();
            I2C_DG_Write(test_parameter.i2c_mtp_dg); // i2c mtp program
            MyLib.Delay1s(2); // wait for program time
            InsControl._power.AutoPowerOff();
            MyLib.Delay1s(1);
            InsControl._power.AutoSelPowerOn(vin);
            MyLib.Delay1ms(1000);
            I2C_DG_Write(test_parameter.i2c_init_dg);
            SeqAndIdealWrite();
#endif
            TriggerEvent(vin); // gpio, i2c(initial), vin trigger

            if (test_parameter.trigger_event == 1)
            {
                I2C_DG_Write(test_parameter.i2c_init_dg);
                RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_dis });
                RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
            }

            if (InsControl._tek_scope_en) MyLib.Delay1s(1);

            MyLib.Delay1ms(900);
            LevelEvent();
            PowerOffEvent();

            InsControl._tek_scope.SetTimeScale(time_scale);
            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");

            //MyLib.Delay1ms(250);
        }

        private void I2C_DG_Write(DataGridView dg)
        {
            for (int i = 0; i < dg.RowCount; i++)
            {
                byte addr = Convert.ToByte(dg[0, i].Value.ToString(), 16);
                byte data = Convert.ToByte(dg[1, i].Value.ToString(), 16);
                RTDev.I2C_Write((byte)(test_parameter.slave), addr, new byte[] { data });
                MyLib.Delay1ms(200);
            }
        }

        private bool TriggerStatus()
        {
            int cnt = 0;
            while (InsControl._tek_scope.doQueryNumber("ACQuire:NUMACq?") == 0)
            {
                cnt++;
                MyLib.Delay1ms(50);
                if (cnt > 100) return false;
            }
            return true;
        }

        private double CursorFunction(int ch1, int ch2, bool direct)
        {
            double res = 0;

            // bool hi_to_lo = dt_test.meas_posCH1[sel] > dly_end_list[sel];
            // int meas_start = start_list[sel];
            // int meas_end = end_list[sel];
            TriggerStatus();

            for(int i = 0; i < 3; i++)
            {
                // enable start channel annotation
                InsControl._tek_scope.DoCommand(string.Format("MEASUrement:ANNOTation:STATE MEAS{0}", ch1));
                MyLib.Delay1ms(800);
                double x1 = direct ?
                    InsControl._tek_scope.doQueryNumber(string.Format("MEASUrement:ANNOTation:X2?")) :
                    InsControl._tek_scope.doQueryNumber(string.Format("MEASUrement:ANNOTation:X1?"));

                InsControl._tek_scope.DoCommand(string.Format("MEASUrement:ANNOTation:STATE MEAS{0}", ch2));
                MyLib.Delay1ms(800);
                double x2 = !direct ?
                    InsControl._tek_scope.doQueryNumber(string.Format("MEASUrement:ANNOTation:X2?")) :
                    InsControl._tek_scope.doQueryNumber(string.Format("MEASUrement:ANNOTation:X1?"));


                InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH" + ch1.ToString());
                MyLib.Delay1ms(600);
                InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH" + ch2.ToString());
                MyLib.Delay1ms(600);
                InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                MyLib.Delay1ms(600);
                InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                MyLib.Delay1ms(600);
                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 " + x1.ToString());
                InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + x2.ToString());
                MyLib.Delay1ms(600);

                res = InsControl._tek_scope.doQueryNumber("CURSor:VBArs:DELTa?");
                MyLib.Delay1ms(100);
                res = InsControl._tek_scope.doQueryNumber("CURSor:VBArs:DELTa?");
            }

            return res;
        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            RTDev.BoadInit();
            RTDev.GpioInit();
            GetParameter(0);

            int vin_cnt = test_parameter.VinList.Count;
            int row = 8;
            int wave_row = 8;
            int wave_pos = 0;
            //string[] binList;
            //double[] ori_vinTable = new double[vin_cnt];
            //int bin_cnt = 1;
            //Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if Report_en
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
#endif
            //InsControl._power.AutoPowerOff();
            OSCInit();
            //double total_ideal = Convert.ToDouble(dt_test.ideal0) +
            //    Convert.ToDouble(dt_test.ideal1) +
            //    Convert.ToDouble(dt_test.ideal2)+
            //    Convert.ToDouble(dt_test.ideal3);


            InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms * Math.Pow(10, -3));
            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            InsControl._tek_scope.SetTimeBasePosition(15);
            MyLib.Delay1s(1);
            int cnt = 0;
            
            #region "Report initial"
#if Report_en
            _sheet = _book.Worksheets.Add();
            _sheet.Name = "DelayTime";
            _sheet.Cells.Font.Name = "Calibri";
            _sheet.Cells.Font.Size = 11;
            row = 8;
            wave_row = 8;
            wave_pos = 0;
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

            // print test conditions
            _sheet.Cells[1, XLS_Table.B] = "Delay time/Slot time";
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions + test_parameter.conditions;

            _sheet.Cells[row, XLS_Table.D] = "No.";
            _sheet.Cells[row, XLS_Table.E] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.F] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.G] = "Conditions";
            _range = _sheet.Range["D" + row, "G" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // major measure timing
            _sheet.Cells[row, XLS_Table.H] = test_parameter.delay_us_en ? "DT1 (us)" : "DT1 (ms)";
            _sheet.Cells[row, XLS_Table.I] = test_parameter.delay_us_en ? "DT2 (us)" : "DT2 (ms)";
            _sheet.Cells[row, XLS_Table.J] = test_parameter.delay_us_en ? "DT3 (us)" : "DT3 (ms)";
            _sheet.Cells[row, XLS_Table.K] = test_parameter.delay_us_en ? "DT4 (us)" : "DT4 (ms)";

            // Add new measure
            _sheet.Cells[row, XLS_Table.L] = "V1 Top (V)";
            _sheet.Cells[row, XLS_Table.M] = "V1 Base (V)";
            _sheet.Cells[row, XLS_Table.N] = "V2 Top (V)";
            _sheet.Cells[row, XLS_Table.O] = "V2 Base (V)";
            _sheet.Cells[row, XLS_Table.P] = "V3 Top (V)";
            _sheet.Cells[row, XLS_Table.Q] = "V3 Base (V)";
            _sheet.Cells[row, XLS_Table.R] = "V4 Top (V)";
            _sheet.Cells[row, XLS_Table.S] = "V4 Base (V)";
            _sheet.Cells[row, XLS_Table.T] = "Pass/Fail";
            //_sheet.Cells[row, XLS_Table.Q] = "Max (V)";
            //_sheet.Cells[row, XLS_Table.R] = "Min (V)";

            _range = _sheet.Range["H" + row, "S" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["T" + row, "T" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            row++;
#endif
            #endregion

            for (int bin_idx = 0; bin_idx < test_parameter.seq_dg.RowCount; bin_idx++)
            {
                int retry_cnt = 0;
                if (test_parameter.run_stop == true) goto Stop;
                if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                /* get test conditions */
                GetParameter(bin_idx);

                double idel_time_total = (dt_test.idealTime0 + dt_test.idealTime1 + dt_test.idealTime2 + dt_test.idealTime2) * Math.Pow(10, -3);
                InsControl._tek_scope.SetTimeScale(idel_time_total / 4);
                /* Eload current setting */
                InsControl._eload.CH1_Loading(dt_test.loading1);
                InsControl._eload.CH2_Loading(dt_test.loading2);
                InsControl._eload.CH3_Loading(dt_test.loading3);
                InsControl._eload.CH4_Loading(dt_test.loading4);

                /* test initial setting */
                string file_name;
                MyLib.Delay1ms(500);
                file_name = string.Format("{0}_Temp={1}C_vin={2:0.##}V_Idealtime1={3}_Idealtime2={4}_Idealtime3={5}_Idealtime4={6}",
                                            cnt, temp,
                                            dt_test.vin,
                                            dt_test.idealTime0,
                                            dt_test.idealTime1,
                                            dt_test.idealTime2,
                                            dt_test.idealTime3
                                         );

                string res = string.Format("Idealtime1={0}_Idealtime2={1}_Idealtime3={2}",
                                            dt_test.idealTime0,
                                            dt_test.idealTime1,
                                            dt_test.idealTime2
                                         );

                double time_scale = 0;
                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
            retest:;
                Scope_Channel_Resize(dt_test.vin);
                InsControl._tek_scope.SetTriggerSource(dt_test.trigger_ch);
                double tempVin = dt_test.vin;
                if (retry_cnt > 3)
                {
                    _sheet.Cells[row, XLS_Table.F] = "sATE test fail";
                    //InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                    InsControl._tek_scope.SetTimeScale(idel_time_total / 4);
                    InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                    InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                    retry_cnt = 0;
                    row++;
                    continue;
                }

                InsControl._tek_scope.SetTriggerMode(false);
                MyLib.Delay1s(2);
                // power on trigger
                switch (test_parameter.trigger_event)
                {
                    case 0:
                        // GPIO trigger event
                        if (InsControl._tek_scope_en)
                            InsControl._tek_scope.SetClear();
                        else
                            InsControl._scope.Root_Clear();

                        //MyLib.Delay1ms(1500);
                        if (test_parameter.sleep_mode)
                        {
                            // sleep mode
                            InsControl._tek_scope.SetTriggerRise();

                            MyLib.Delay1ms(800);
                            GpioOnSelect(test_parameter.gpio_pin);
                        }
                        else
                        {
                            // PWRDis
                            InsControl._tek_scope.SetTriggerFall();

                            MyLib.Delay1ms(1000);
                            GpioOffSelect(test_parameter.gpio_pin);
                        }

                        if (InsControl._tek_scope_en) MyLib.Delay1s(1);
                        break;
                    case 1:
                        RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                        MyLib.Delay1s(1);
                        break;
                    case 2:
                        // Power supply trigger event
                        InsControl._power.AutoSelPowerOn(dt_test.vin);
                        MyLib.Delay1ms((int)((time_scale * 10) * 1.2) + 500);
                        break;
                    case 3:
                        // rail trigger
                        RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                        MyLib.Delay1s(1);
                        break;
                }
                InsControl._tek_scope.SetStop();

                time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
                if (time_scale >= 0.005) MyLib.Delay1s(5);
                int ch1;
                int ch2;
                bool direct;

                double delay_time_res1 = 0, delay_time_res2 = 0, delay_time_res3 = 0, delay_time_res4 = 0;
                double delay_time_res = 0;

                if (test_parameter.seq_en[0])
                {
                    ch1 = dt_test.meas_posCH1[0];
                    ch2 = dt_test.meas_posCH1[1];
                    direct = dt_test.precentCH1[0] > dt_test.precentCH1[1];
                    delay_time_res1 = CursorFunction(ch1, ch2, direct); // ideal time1
                    MyLib.Delay1ms(100);
                    if (!test_parameter.cursor_disable) InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_seq0");
                    delay_time_res = delay_time_res1;
                }

                if (test_parameter.seq_en[1])
                {
                    ch1 = dt_test.meas_posCH2[0];
                    ch2 = dt_test.meas_posCH2[1];
                    direct = dt_test.precentCH2[0] > dt_test.precentCH2[1];
                    delay_time_res2 = CursorFunction(ch1, ch2, direct); // ideal time2
                    MyLib.Delay1ms(100);
                    if (!test_parameter.cursor_disable) InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_seq1");
                    delay_time_res += delay_time_res2;
                }


                if (test_parameter.seq_en[2])
                {
                    ch1 = dt_test.meas_posCH3[0];
                    ch2 = dt_test.meas_posCH3[1];
                    direct = dt_test.precentCH3[0] > dt_test.precentCH3[1];
                    delay_time_res3 = CursorFunction(ch1, ch2, direct); // ideal time3
                    MyLib.Delay1ms(100);
                    if (!test_parameter.cursor_disable) InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_seq2");
                    delay_time_res += delay_time_res3;
                }

                if (test_parameter.seq_en[3])
                {
                    ch1 = dt_test.meas_posCH4[0];
                    ch2 = dt_test.meas_posCH4[1];
                    direct = dt_test.precentCH4[0] > dt_test.precentCH4[1];
                    if(!test_parameter.cursor_disable) delay_time_res4 = CursorFunction(ch1, ch2, direct); // ideal time4
                    MyLib.Delay1ms(100);
                    InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_seq3");
                    delay_time_res += delay_time_res4;
                }

                double us_unit = Math.Pow(10, -6);
                double ms_unit = Math.Pow(10, -3);
                double[] time_table = new double[] {
                                500 * us_unit, 400 * us_unit, 200 * us_unit, 100 * us_unit, 50 * us_unit, 20 * us_unit, 10 * us_unit,
                                40 * ms_unit, 20 * ms_unit, 10 * ms_unit, 4 * ms_unit, 2 * ms_unit, 1 * ms_unit
                            };
                List<double> min_list = new List<double>();
                double time_temp = (delay_time_res) / 4.5;
                double time_div = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");

                // scope time scale re-size
                if (delay_time_res > Math.Pow(10, 20) || delay_time_res < 0)
                {
                    //InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                    InsControl._tek_scope.SetTimeScale(idel_time_total / 4);
                    InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                    InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                    InsControl._tek_scope.SetTimeBasePosition(15);
                    InsControl._tek_scope.SetRun();
                    InsControl._tek_scope.SetTriggerMode();
                    PowerOffEvent();
                    retry_cnt++;
                    goto retest;
                }
                else if (delay_time_res > time_div * 4)
                {
                    InsControl._tek_scope.SetTimeScale(time_temp);
                    InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                    InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                    InsControl._tek_scope.SetTimeBasePosition(15);

                    if (!(time_div == InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?")))
                    {
                        InsControl._tek_scope.SetRun();
                        InsControl._tek_scope.SetTriggerMode();
                        PowerOffEvent();

                        retry_cnt++;
                        goto retest;
                    }
                }


                MyLib.Delay1ms(100);
                //InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name);
#if true
                double vin = 0, dt1 = 0, dt2 = 0, dt3 = 0, dt4 = 0;
                double vtop1, vtop2, vtop3, vtop4;
                double vbase1, vbase2, vbase3, vbase4;

                InsControl._tek_scope.SetMeasureSource(1, meas_vmax1, "HIGH");
                InsControl._tek_scope.SetMeasureSource(2, meas_vmax2, "HIGH");
                InsControl._tek_scope.SetMeasureSource(3, meas_vmax3, "HIGH");
                InsControl._tek_scope.SetMeasureSource(4, meas_vmax4, "HIGH");
                vtop1 = InsControl._tek_scope.MeasureMax(meas_vmax1);
                vtop2 = InsControl._tek_scope.MeasureMax(meas_vmax2);
                vtop3 = InsControl._tek_scope.MeasureMax(meas_vmax3);
                vtop4 = InsControl._tek_scope.MeasureMax(meas_vmax4);

                InsControl._tek_scope.SetMeasureSource(1, meas_vmax1, "LOW");
                InsControl._tek_scope.SetMeasureSource(2, meas_vmax2, "LOW");
                InsControl._tek_scope.SetMeasureSource(3, meas_vmax3, "LOW");
                InsControl._tek_scope.SetMeasureSource(4, meas_vmax4, "LOW");
                vbase1 = InsControl._tek_scope.MeasureMean(meas_vmax1);
                vbase2 = InsControl._tek_scope.MeasureMean(meas_vmax2);
                vbase3 = InsControl._tek_scope.MeasureMean(meas_vmax3);
                vbase4 = InsControl._tek_scope.MeasureMean(meas_vmax4);
                dt1 = delay_time_res1;
                dt2 = delay_time_res2;
                dt3 = delay_time_res3;
                dt4 = delay_time_res4;
#if Power_en
                vin = InsControl._power.GetVoltage();
#endif
#if Report_en
                _sheet.Cells[row, XLS_Table.D] = cnt++;
                _sheet.Cells[row, XLS_Table.E] = temp;
                _sheet.Cells[row, XLS_Table.F] = vin;
                _sheet.Cells[row, XLS_Table.G] = res; // ideal / delay time condtions

                double calculate_dt = (test_parameter.delay_us_en ? dt1 * Math.Pow(10, 6) : dt1 * Math.Pow(10, 3));
                _sheet.Cells[row, XLS_Table.H] = calculate_dt.ToString();
                _sheet.Cells[row, XLS_Table.L] = vtop1.ToString();
                _sheet.Cells[row, XLS_Table.M] = vbase1.ToString();
                calculate_dt = (test_parameter.delay_us_en ? dt2 * Math.Pow(10, 6) : dt2 * Math.Pow(10, 3));
                _sheet.Cells[row, XLS_Table.I] = calculate_dt.ToString();
                _sheet.Cells[row, XLS_Table.N] = vtop2.ToString();
                _sheet.Cells[row, XLS_Table.O] = vbase2.ToString();
                calculate_dt = (test_parameter.delay_us_en ? dt3 * Math.Pow(10, 6) : dt3 * Math.Pow(10, 3));
                _sheet.Cells[row, XLS_Table.J] = calculate_dt.ToString();
                _sheet.Cells[row, XLS_Table.P] = vtop3.ToString();
                _sheet.Cells[row, XLS_Table.Q] = vbase3.ToString();
                calculate_dt = (test_parameter.delay_us_en ? dt4 * Math.Pow(10, 6) : dt4 * Math.Pow(10, 3));
                _sheet.Cells[row, XLS_Table.K] = calculate_dt.ToString();
                _sheet.Cells[row, XLS_Table.R] = vtop4.ToString();
                _sheet.Cells[row, XLS_Table.S] = vbase4.ToString();

                /* spec judge */
                double value_base = dt_test.idealTime0 * (test_parameter.delay_us_en ? Math.Pow(10, -6) : Math.Pow(10, -3));
                double criteria_up = (test_parameter.judge_percent * value_base) + value_base;
                double criteria_down = value_base - (test_parameter.judge_percent * value_base);
                double value = 0;
                string jd_res = "Pass";

                value = dt1;
                if (value > criteria_up || value < criteria_down)
                {
                    jd_res = "Fail";
                }

                value = dt2;
                value_base = dt_test.idealTime1 * (test_parameter.delay_us_en ? Math.Pow(10, -6) : Math.Pow(10, -3));
                criteria_up = (test_parameter.judge_percent * value_base) + value_base;
                criteria_down = value_base - (test_parameter.judge_percent * value_base);
                if (value > criteria_up || value < criteria_down)
                {
                    jd_res = "Fail";
                }

                value = dt3;
                value_base = dt_test.idealTime2 * (test_parameter.delay_us_en ? Math.Pow(10, -6) : Math.Pow(10, -3));
                criteria_up = (test_parameter.judge_percent * value_base) + value_base;
                criteria_down = value_base - (test_parameter.judge_percent * value_base);
                if (value > criteria_up || value < criteria_down)
                {
                    jd_res = "Fail";
                }

                if (jd_res == "Fail")
                {
                    _range = _sheet.Range["T" + row];
                    _range.Interior.Color = Color.Red;
                }
                else
                {
                    _range = _sheet.Range["T" + row];
                    _range.Interior.Color = Color.LightGreen;
                }
                _sheet.Cells[row, XLS_Table.T] = jd_res;

                _sheet.Cells[wave_row, XLS_Table.AA] = "No.";
                _sheet.Cells[wave_row, XLS_Table.AB] = "Temp(C)";
                _sheet.Cells[wave_row, XLS_Table.AC] = "Vin(V)";
                _sheet.Cells[wave_row, XLS_Table.AD] = "Conditions";
                _range = _sheet.Range["AA" + wave_row, "AD" + wave_row];
                _range.Interior.Color = Color.FromArgb(124, 252, 0);
                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                _sheet.Cells[wave_row + 1, XLS_Table.AA] = "=D" + row;
                _sheet.Cells[wave_row + 1, XLS_Table.AB] = "=E" + row;
                _sheet.Cells[wave_row + 1, XLS_Table.AC] = "=F" + row;
                _sheet.Cells[wave_row + 1, XLS_Table.AD] = "=G" + row;
                _range = _sheet.Range["AA" + (wave_row + 2).ToString(), "AG" + (wave_row + 16).ToString()];
                MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_seq0");

                _range = _sheet.Range["AI" + (wave_row + 2).ToString(), "AO" + (wave_row + 16).ToString()];
                MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_seq1");

                _range = _sheet.Range["AQ" + (wave_row + 2).ToString(), "AW" + (wave_row + 16).ToString()];
                MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_seq2");

                #region "old version past waveform"
                
                //switch (wave_pos)
                //{
                //    case 0:
                //        _sheet.Cells[wave_row, XLS_Table.AA] = "No.";
                //        _sheet.Cells[wave_row, XLS_Table.AB] = "Temp(C)";
                //        _sheet.Cells[wave_row, XLS_Table.AC] = "Vin(V)";
                //        _sheet.Cells[wave_row, XLS_Table.AD] = "Conditions";
                //        _range = _sheet.Range["AA" + wave_row, "AD" + wave_row];
                //        _range.Interior.Color = Color.FromArgb(124, 252, 0);
                //        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //        _sheet.Cells[wave_row + 1, XLS_Table.AA] = "=D" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AB] = "=E" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AC] = "=F" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AD] = "=G" + row;
                //        _range = _sheet.Range["AA" + (wave_row + 2).ToString(), "AG" + (wave_row + 16).ToString()];
                //        wave_pos++;
                //        break;
                //    case 1:
                //        _sheet.Cells[wave_row, XLS_Table.AL] = "No.";
                //        _sheet.Cells[wave_row, XLS_Table.AM] = "Temp(C)";
                //        _sheet.Cells[wave_row, XLS_Table.AN] = "Vin(V)";
                //        _sheet.Cells[wave_row, XLS_Table.AO] = "Conditions";
                //        _range = _sheet.Range["AL" + wave_row, "AO" + wave_row];
                //        _range.Interior.Color = Color.FromArgb(124, 252, 0);
                //        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //        _sheet.Cells[wave_row + 1, XLS_Table.AL] = "=D" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AM] = "=E" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AN] = "=F" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AO] = "=G" + row;
                //        _range = _sheet.Range["AL" + (wave_row + 2).ToString(), "AR" + (wave_row + 16).ToString()];
                //        wave_pos++;
                //        break;
                //    case 2:
                //        _sheet.Cells[wave_row, XLS_Table.AW] = "No.";
                //        _sheet.Cells[wave_row, XLS_Table.AX] = "Temp(C)";
                //        _sheet.Cells[wave_row, XLS_Table.AY] = "Vin(V)";
                //        _sheet.Cells[wave_row, XLS_Table.AZ] = "Conditions";
                //        _range = _sheet.Range["AW" + wave_row, "AZ" + wave_row];
                //        _range.Interior.Color = Color.FromArgb(124, 252, 0);
                //        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //        _sheet.Cells[wave_row + 1, XLS_Table.AW] = "=D" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AX] = "=E" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AY] = "=F" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.AZ] = "=G" + row;
                //        _range = _sheet.Range["AW" + (wave_row + 2).ToString(), "BC" + (wave_row + 16).ToString()];
                //        wave_pos++;
                //        break;
                //    case 3:
                //        _sheet.Cells[wave_row, XLS_Table.BH] = "No.";
                //        _sheet.Cells[wave_row, XLS_Table.BI] = "Temp(C)";
                //        _sheet.Cells[wave_row, XLS_Table.BJ] = "Vin(V)";
                //        _sheet.Cells[wave_row, XLS_Table.BK] = "Conditions";
                //        _range = _sheet.Range["BH" + wave_row, "BK" + wave_row];
                //        _range.Interior.Color = Color.FromArgb(124, 252, 0);
                //        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                //        _sheet.Cells[wave_row + 1, XLS_Table.BH] = "=D" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.BI] = "=E" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.BJ] = "=F" + row;
                //        _sheet.Cells[wave_row + 1, XLS_Table.BK] = "=G" + row;
                //        _range = _sheet.Range["BH" + (wave_row + 2).ToString(), "BN" + (wave_row + 16).ToString()];
                //        wave_pos = 0; wave_row = wave_row + 19;
                //        break;
                //} // switch case

                //MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name);
                #endregion
#endif
                row++;
#endif
                InsControl._tek_scope.SetRun();
                PowerOffEvent();
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string str_temp = _sheet.Cells[2, XLS_Table.B].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
#if Report_en
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_DT_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif

        }

        private void PowerOffEvent()
        {
            switch (test_parameter.trigger_event)
            {
                case 0: // gpio power disable
                    if (test_parameter.sleep_mode)
                        GpioOffSelect(test_parameter.gpio_pin);
                    else
                        GpioOnSelect(test_parameter.gpio_pin);
                    break;
                case 1:
                    // rails disable
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_dis });
                    I2C_DG_Write(test_parameter.i2c_init_dg);
                    break;
                case 2: // vin trigger
#if Power_en
                    InsControl._power.AutoPowerOff();
#endif
                    break;
                case 3:
                    // rail trigger
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_dis });
                    break;
            }
        }

        private void GpioOnSelect(int num)
        {
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Enable();
                    break;
                case 1:
                    RTDev.Gp2En_Enable();
                    break;
                case 2:
                    RTDev.Gp3En_Enable();
                    break;
            }
        }

        private void GpioOffSelect(int num)
        {
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Disable();
                    break;
                case 1:
                    RTDev.Gp2En_Disable();
                    break;
                case 2:
                    RTDev.Gp3En_Disable();
                    break;
            }
        }

    }
}





