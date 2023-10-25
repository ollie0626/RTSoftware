
#define Report_en
#define Power_en
#define Eload_en

using RTBBLibDotNet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoftStartTiming
{
    public class ATE_VIDIO : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        //public new double temp;
        RTBBControl RTDev = new RTBBControl();

        const int LPM = 0;
        const int G1 = 1;
        const int G2 = 2;
        //const int test_cnt = 5;

        List<double> overshoot_list = new List<double>();
        List<double> undershoot_list = new List<double>();
        List<double> slewrate_list = new List<double>();

        List<double> rise_time_list = new List<double>();
        List<double> fall_time_list = new List<double>();
        List<double> delay_list = new List<double>();
        List<double> delay100_list = new List<double>();

        List<double> vmax_list = new List<double>();
        List<double> vmin_list = new List<double>();
        List<string> phase1_name = new List<string>();
        List<string> phase2_name = new List<string>();


        int meas_rising = 1;
        int meas_falling = 2;
        int meas_vmax = 3;
        int meas_vmin = 4;
        int meas_delay = 5;
        int meas_delay100 = 6;
        const int cnt_rest = 10000;
        const int cursor_fail_cnt = 10;


        public delegate void FinishNotification();
        FinishNotification delegate_mess;
        VIDIO updateMain;
        int progress = 0;

        public ATE_VIDIO(VIDIO main)
        {
            delegate_mess = new FinishNotification(MessageNotify);
            updateMain = main;
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("VIDIO test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void IOStateSetting(int state)
        {
            //int value = (lpm << 0 | g1 << 1 | g2 << 2);
            int mask = 1 << LPM | 1 << G1 | 1 << G2;

            RTDev.GPIOnState((uint)mask, (uint)state);
        }

        private void OSCInit()
        {
            InsControl._oscilloscope.SetRST();
            MyLib.Delay1s(3);
            InsControl._oscilloscope.CHx_On(1); // vout
            InsControl._oscilloscope.CHx_On(2); // Lx
            InsControl._oscilloscope.CHx_On(3); // G1
            InsControl._oscilloscope.CHx_On(4); // G2
            MyLib.Delay1s(2);
            InsControl._oscilloscope.CHx_BWLimitOn(1);
            InsControl._oscilloscope.CHx_BWLimitOn(2);
            InsControl._oscilloscope.CHx_BWLimitOn(3);
            InsControl._oscilloscope.CHx_BWLimitOn(4);

            // initial time scale
            InsControl._oscilloscope.SetTimeScale(4 * Math.Pow(10, -6));
            InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
            InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
            InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");

            InsControl._oscilloscope.CHx_Level(3, 5);
            InsControl._oscilloscope.CHx_Level(4, 5);
            InsControl._oscilloscope.CHx_Position(3, 3);
            InsControl._oscilloscope.CHx_Position(4, 3);

            InsControl._oscilloscope.SetMeasureSource(1, meas_rising, "RISE");
            InsControl._oscilloscope.SetMeasureSource(1, meas_falling, "FALL");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmax, "MAXimum");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmin, "MINImum");

            double vout = 0;
            double vout_af = 0;

            try
            {
                vout = (double)test_parameter.vidio.vout_list[0];
            }
            catch
            {

            }

            try
            {
                vout_af = (double)test_parameter.vidio.vout_list_af[0];
            }
            catch
            {

            }

            double max = vout > vout_af ? vout : vout_af;
            double min = vout < vout_af ? vout : vout_af;

            InsControl._oscilloscope.CHx_Level(1, max - min / 3);
            InsControl._oscilloscope.CHx_Offset(1, min);
            InsControl._oscilloscope.CHx_Position(1, -2);
            MyLib.Delay1s(2);
            InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[0] / 1.5);
            InsControl._oscilloscope.CHx_Position(2, -4);
            InsControl._oscilloscope.SetAutoTrigger();
            InsControl._oscilloscope.SetTriggerLevel(2, max - min);
            InsControl._oscilloscope.SetTimeBasePosition(25);
        }

        private void Initial_TimeScale(bool rising_en, int case_idx)
        {
            double time_scale = 5;
            double vout = 0;
            double vout_af = 0;

            if (test_parameter.vidio.criteria[case_idx].lpm_en)
                vout = 0;
            else
                vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);

            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);

            if (rising_en)
            {
                // normal mode
                if ((string)test_parameter.vidio.criteria[case_idx].rise_time != "NA" && !test_parameter.vidio.criteria[case_idx].lpm_en)
                {
                    time_scale = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].rise_time) * Math.Pow(10, -6);
                }
                else
                {
                    double delta_v = vout_af - vout;
                    double parameter = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_rise) / 1000;
                    time_scale = ((delta_v / parameter)) * Math.Pow(10, -6);
                }

                InsControl._oscilloscope.SetTimeScale(time_scale / 2);
                InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
            else
            {
                if ((string)test_parameter.vidio.criteria[case_idx].fall_time != "NA")
                {
                    time_scale = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].fall_time) * Math.Pow(10, -6);
                }
                else
                {
                    double delta_v = vout_af - vout;
                    double parameter = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_fall) / 1000;
                    time_scale = ((delta_v / parameter)) * Math.Pow(10, -6);
                }

                InsControl._oscilloscope.SetTimeScale(time_scale / 2);
                InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
        }

        private bool TriggerStatus()
        {
            int cnt = 0;
            while (InsControl._oscilloscope.GetCount() == 0)
            {
                cnt++;
                MyLib.Delay1ms(50);
                if (cnt > 100) return false;
            }
            return true;
        }

        private void Scope_Task_Setting(int meas_idx, double vout, double vout_af)
        {
            //InsControl._oscilloscope.SetTimeOutTrigger();
            //InsControl._oscilloscope.SetTimeOutTriggerCHx(1);
            InsControl._oscilloscope.SetTimeOutTime(5 * Math.Pow(10, -12));
            InsControl._oscilloscope.DoCommand("TRIGger:A:EDGE:SLOpe EITher");
            InsControl._oscilloscope.DoCommand("TRIGger:A:LEVel 1.2");
            //InsControl._oscilloscope.SetTimeOutEither();

            InsControl._oscilloscope.CHx_Level(1, (vout_af - vout) / 4.7);
            InsControl._oscilloscope.CHx_Offset(1, vout);
            InsControl._oscilloscope.CHx_Position(1, -2);
            //InsControl._oscilloscope.SetTriggerLevel(1, (vout_af - vout) * 0.5 + vout);
            InsControl._oscilloscope.SetAnnotation(meas_idx);
        }

        private void GetTriggerSel(int initial, int next)
        {
            int initial_G01 = (initial & 0x06) >> 1;
            int next_G01 = (next & 0x06) >> 1;
            int res = initial_G01 ^ next_G01;
            int ch = 0;
            for (int i = 0; i < 2; i++)
            {
                if ((res & (0x01 << i)) != 0)
                {
                    ch = i;
                    break;
                }
            }
            InsControl._oscilloscope.DoCommand(string.Format("TRIGger:A:EDGE:SOUrce CH{0}", ch + 3));
        }

        private void LPMTrigger(int meas_idx)
        {
            InsControl._oscilloscope.SetTriggerLevel(3, 1.2);
            InsControl._oscilloscope.DoCommand("TRIGger:A:EDGE:SLOpe EITher");
            InsControl._oscilloscope.DoCommand("TRIGger:A:LEVel 1.2");
            InsControl._oscilloscope.SetAnnotation(meas_idx);
        }

        private bool SlewRate_Rise_Task(int case_idx, bool overshoot_en = false)
        {
            vmax_list.Clear();
            slewrate_list.Clear();
            rise_time_list.Clear();
            overshoot_list.Clear();
            delay_list.Clear();
            delay100_list.Clear();

            double vout = 0;
            double vout_af = 0;
            if (test_parameter.vidio.criteria[case_idx].lpm_en)
                vout = 0;
            else
                vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);
            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);

            int initial_state = 0;
            int next_state = 0;

            if (test_parameter.vidio.criteria[case_idx].lpm_en)
            {
                initial_state = test_parameter.vidio.vout_map[test_parameter.vidio.criteria[case_idx].vout_begin];
            }
            else
            {
                initial_state = test_parameter.vidio.vout_map[vout.ToString()];
            }

            next_state = test_parameter.vidio.vout_map[vout_af.ToString()];

            Scope_Task_Setting(meas_rising, vout, vout_af); // trigger and time scale
            IOStateSetting(initial_state);

            double hi = test_parameter.vidio.criteria[case_idx].hi;
            double lo = test_parameter.vidio.criteria[case_idx].lo;
            double deltaV = (hi - lo) * 1000;
            // setting measure level threshold
            // example : (0.9 - 0.5) / 2 + 0.5 = 0.7
            InsControl._oscilloscope.SetREFLevelMethod(meas_rising, false);
            InsControl._oscilloscope.SetREFLevel(hi, lo + ((hi * lo) / 2), lo, meas_rising, false);


            if (test_parameter.vidio.criteria[case_idx].lpm_en)
            {
                InsControl._oscilloscope.SetREFLevelMethod(meas_rising, true);
                InsControl._oscilloscope.SetREFLevel(100, 50, 1, meas_rising, true);

                InsControl._oscilloscope.DoCommand(string.Format("MEASUrement:MEAS{0}:REFLevel:PERCent:MID2 8", meas_delay));
                InsControl._oscilloscope.DoCommand(string.Format("MEASUrement:MEAS{0}:REFLevel:PERCent:MID2 100", meas_delay100));
                InsControl._oscilloscope.SetDelayTime(meas_delay, 3, 1, true, true);
                InsControl._oscilloscope.SetDelayTime(meas_delay100, 3, 1, true, true);
                InsControl._oscilloscope.SetAnnotation(meas_delay);
                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetAnnotation(meas_delay);
                InsControl._oscilloscope.SetCursorSource(1, 3); MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetCursorSource(2, 1); MyLib.Delay1ms(100);
            }

            InsControl._oscilloscope.SetCursorWaveform();
            InsControl._oscilloscope.SetCursorOn();
            Initial_TimeScale(true, case_idx);

            if (overshoot_en)
            {
                InsControl._oscilloscope.SetCursorOff();
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                InsControl._oscilloscope.DoCommand("MEASUrement:ANNOTation:STATE OFF");
                MyLib.Delay1ms(500);
            }

            if (!test_parameter.vidio.criteria[case_idx].lpm_en)
                GetTriggerSel(initial_state, next_state);
            else
                LPMTrigger(meas_rising);

            for (int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double rise_time = 0;
                double vmax = 0;
                double delay = 0;
                double delay100 = 0;
                int cnt = 0;
                bool keep_test_en = false;

            Trigger_Fail_retry:
                IOStateSetting(initial_state);
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(800);
                InsControl._oscilloscope.SetNormalTrigger();
                MyLib.Delay1ms(100);

                IOStateSetting(next_state);
                cnt++;

                if (cnt == cnt_rest) return false;
                if (!TriggerStatus())
                {
                    //InsControl._oscilloscope.SetClear();
                    InsControl._oscilloscope.SetStop();
                    MyLib.Delay1ms(200);
                    Console.WriteLine("rise cnt = {0}", cnt);
                    goto Trigger_Fail_retry;
                }

                InsControl._oscilloscope.SetStop();
                if (repeat_idx == 0) MyLib.Delay1ms(200);
                double x1 = 0;
                double x2 = 0;

                if (test_parameter.vidio.criteria[case_idx].lpm_en && repeat_idx == 0 && !overshoot_en)
                {
                    cnt = 0;
                    InsControl._oscilloscope.SetAnnotation(meas_delay);
                    MyLib.Delay1ms(500);
                DelayGetX:
                    x1 = InsControl._oscilloscope.GetAnnotationXn(1); MyLib.Delay1ms(100);
                    x2 = InsControl._oscilloscope.GetAnnotationXn(2); MyLib.Delay1ms(100);
                    if (x1 < 0 || x1 > 10 * Math.Pow(10, 9) || x2 < 0 || x2 > 10 * Math.Pow(10, 9))
                    {
                        if (cnt == cursor_fail_cnt) return false;
                        cnt++;
                        //Console.WriteLine("Get cursor fail");
                        goto DelayGetX;
                    }
                    InsControl._oscilloscope.SetCursorSource(1, 3); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorSource(2, 1); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);
                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + "_delay");
                }


                // set cursor position
                if (!overshoot_en)
                {
                    InsControl._oscilloscope.SetAnnotation(meas_rising);
                    InsControl._oscilloscope.SetAnnotation(meas_rising);
                }

                if (overshoot_en) InsControl._oscilloscope.DoCommand("MEASUrement:ANNOTation:STATE OFF");
                if (repeat_idx == 0) MyLib.Delay1ms(500);
                else MyLib.Delay1ms(50);

                if (!overshoot_en)
                {
                    cnt = 0;
                GetX:
                    x1 = InsControl._oscilloscope.GetAnnotationXn(1); MyLib.Delay1ms(100);
                    x2 = InsControl._oscilloscope.GetAnnotationXn(2); MyLib.Delay1ms(100);

                    if (x1 < 0 || x1 > 10 * Math.Pow(10, 9) || x2 < 0 || x2 > 10 * Math.Pow(10, 9))
                    {
                        if (cnt == cursor_fail_cnt) return false;
                        cnt++;
                        //Console.WriteLine("Get cursor fail");
                        goto GetX;
                    }

                    InsControl._oscilloscope.SetCursorSource(1, 1); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorSource(2, 1); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorSource(1, 1); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorSource(2, 1); MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);
                }

                vmax = InsControl._oscilloscope.MeasureMax(meas_vmax);
                vmax_list.Add(vmax);

                // get delta T
                if (!overshoot_en)
                {
                    //rise_time = InsControl._oscilloscope.GetCursorVBarDelta();
                    //// slew rate delta V / delta T
                    //slew_rate = InsControl._oscilloscope.GetCursorHBarDelta() / rise_time;

                    rise_time = InsControl._oscilloscope.MeasureMin(meas_rising);
                    slew_rate = (deltaV / rise_time) * Math.Pow(10, -3);

                    Console.WriteLine("Rise time = {0}, Rise SR = {1}", rise_time, slew_rate);


                    if (test_parameter.vidio.criteria[case_idx].lpm_en)
                    {
                        delay = InsControl._oscilloscope.MeasureMean(meas_delay);
                        delay_list.Add(delay);

                        delay100 = InsControl._oscilloscope.MeasureMean(meas_delay100);
                        delay100_list.Add(delay100);
                    }


                    slewrate_list.Add(slew_rate);
                    rise_time_list.Add(rise_time);

                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, (repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_rising");
                    phase1_name.Add((repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_rising");
                }
                else
                {
                    // measure overshoot
                    double res = (vmax - vout_af) / vout_af;
                    overshoot_list.Add(res);
                    //overshoot_list.Add(vmax);

                }
            }

            if (overshoot_en)
            {
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetPERSistenceOff();
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + "_overshoot");
            }

            return true;

        }

        private bool SlewRate_Fall_Task(int case_idx, bool undershoot_en = false)
        {
            vmin_list.Clear();
            slewrate_list.Clear();
            fall_time_list.Clear();
            undershoot_list.Clear();

            double vout = 0;
            double vout_af = 0;
            if (test_parameter.vidio.criteria[case_idx].lpm_en)
                vout = 0;
            else
                vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);

            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);

            int initial_state = 0;
            int next_state = 0;

            if (test_parameter.vidio.criteria[case_idx].lpm_en)
            {
                initial_state = test_parameter.vidio.vout_map[test_parameter.vidio.criteria[case_idx].vout_begin];
            }
            else
            {
                initial_state = test_parameter.vidio.vout_map[vout.ToString()];
            }
            next_state = test_parameter.vidio.vout_map[vout_af.ToString()];
            Scope_Task_Setting(meas_falling, vout, vout_af);

            IOStateSetting(next_state);
            //InsControl._oscilloscope.SetTriggerFall();

            double hi = test_parameter.vidio.criteria[case_idx].hi;
            double lo = test_parameter.vidio.criteria[case_idx].lo;
            double deltaV = (hi - lo) * 1000;

            // setting measure level threshold
            // example : (0.9 - 0.5) / 2 + 0.5 = 0.7
            InsControl._oscilloscope.SetREFLevelMethod(meas_falling, false);
            InsControl._oscilloscope.SetREFLevel(hi, lo + ((hi * lo) / 2), lo, meas_falling, false);
            InsControl._oscilloscope.SetCursorWaveform();
            InsControl._oscilloscope.SetCursorOn();
            Initial_TimeScale(false, case_idx);


            if (test_parameter.vidio.criteria[case_idx].lpm_en)
            {
                InsControl._oscilloscope.SetREFLevelMethod(meas_falling, true);
                InsControl._oscilloscope.SetREFLevel(100, 50, 1, meas_falling, true);
            }

            if (test_parameter.vidio.criteria[case_idx].lpm_en)
                InsControl._oscilloscope.SetTimeScale(test_parameter.vidio.discharge_time * Math.Pow(10, -3));

            if (undershoot_en)
            {
                InsControl._oscilloscope.SetCursorOff();
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                InsControl._oscilloscope.DoCommand("MEASUrement:ANNOTation:STATE OFF");
                MyLib.Delay1ms(500);
            }

            if (!test_parameter.vidio.criteria[case_idx].lpm_en)
                GetTriggerSel(initial_state, next_state);
            else
                LPMTrigger(meas_falling);

            for (int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double fall_time = 0;
                double vmin = 0;
                int cnt = 0;

            Trigger_Fail_retry:
                IOStateSetting(next_state);
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(500);
                InsControl._oscilloscope.SetNormalTrigger();
                MyLib.Delay1ms(100);
                cnt++;
                IOStateSetting(initial_state);
                if (!TriggerStatus())
                {
                    IOStateSetting(next_state);
                    //InsControl._oscilloscope.SetClear();
                    InsControl._oscilloscope.SetStop();
                    MyLib.Delay1ms(200);
                    Console.WriteLine("fall cnt = {0}", cnt);
                    if (cnt == cnt_rest) return false;
                    goto Trigger_Fail_retry;
                }
                InsControl._oscilloscope.SetStop();
                if (repeat_idx == 0) MyLib.Delay1ms(200);

                // set cursor position
                if (!undershoot_en)
                {
                    InsControl._oscilloscope.SetAnnotation(meas_falling); MyLib.Delay1ms(50);
                    InsControl._oscilloscope.SetAnnotation(meas_falling); MyLib.Delay1ms(50);
                }


                if (undershoot_en) InsControl._oscilloscope.DoCommand("MEASUrement:ANNOTation:STATE OFF");

                if (!undershoot_en)
                {
                    cnt = 0;
                GetX:
                    double x1 = InsControl._oscilloscope.GetAnnotationXn(1); MyLib.Delay1ms(100);
                    double x2 = InsControl._oscilloscope.GetAnnotationXn(2); MyLib.Delay1ms(100);
                    //InsControl._oscilloscope.SetCursorMode();
                    //InsControl._oscilloscope.SetCursorOn();

                    if (x1 < 0 || x1 > 10 * Math.Pow(10, 9) || x2 < 0 || x2 > 10 * Math.Pow(10, 9))
                    {
                        if (cnt == cursor_fail_cnt) return false;
                        cnt++;
                        //Console.WriteLine("Get cursor fail");
                        goto GetX;

                    }
                    InsControl._oscilloscope.SetCursorSource(1, 1);
                    InsControl._oscilloscope.SetCursorSource(2, 1);
                    InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);
                }
                vmin = InsControl._oscilloscope.MeasureMin(meas_vmin);
                vmin_list.Add(vmin);

                // get delta T
                if (!undershoot_en)
                {
                    fall_time = InsControl._oscilloscope.GetCursorVBarDelta();
                    // slew rate delta V / delta T
                    slew_rate = Math.Abs(InsControl._oscilloscope.GetCursorHBarDelta() / fall_time);

                    fall_time = InsControl._oscilloscope.MeasureMin(meas_falling);
                    slew_rate = (deltaV / fall_time) * Math.Pow(10, -3);

                    Console.WriteLine("Fall time = {0}, Fall SR = {1}", fall_time, slew_rate);
                    slewrate_list.Add(slew_rate);
                    fall_time_list.Add(fall_time);
                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, (repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_falling");
                    phase2_name.Add((repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_falling");
                }
                else
                {
                    // measure undershoot
                    double res = (vmin - vout) / vout;
                    //undershoot_list.Add(vmin);
                    undershoot_list.Add(res);
                }
            }

            if (undershoot_en)
            {
                //InsControl._oscilloscope.SetCursorOff();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetPERSistenceOff();
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + "_undershoot");
            }

            return true;
        }

        public override void ATETask()
        {
            progress = 0;
            updateMain.UpdateProgressBar(0);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            RTDev.GpioInit();
            OSCInit();
            int row = 10;
            int wave_row = 10;
            int wave_idx = 0;
            int idx = 0;
            string file_name = "";

            MyLib.CreateSaveWaveformFolder(test_parameter.waveform_path);

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
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver
                                            + test_parameter.vin_conditions
                                            + test_parameter.iout_conditions;

            _sheet.Cells[row, XLS_Table.C] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.D] = "超連結";
            _sheet.Cells[row, XLS_Table.E] = "Vin (V)";
            _sheet.Cells[row, XLS_Table.F] = "Vout Change (V)";
            _sheet.Cells[row, XLS_Table.G] = "VID spec (V)";
            _sheet.Cells[row, XLS_Table.H] = "Iout (A)";
            _sheet.Cells[row, XLS_Table.I] = "Rise Time spec (us)";
            _sheet.Cells[row, XLS_Table.J] = "Rise SR spec (V/us)";
            _sheet.Cells[row, XLS_Table.K] = "Rise SR (V/us)";
            _sheet.Cells[row, XLS_Table.L] = "Rise Time (us)";
            _sheet.Cells[row, XLS_Table.M] = "Fall Time spec (us)";
            _sheet.Cells[row, XLS_Table.N] = "Fall SR spec (V/us)";
            _sheet.Cells[row, XLS_Table.O] = "Fall SR (V/us)";
            _sheet.Cells[row, XLS_Table.P] = "Fall Time (us)";
            _sheet.Cells[row, XLS_Table.Q] = "Vmax spec (V)"; // overshoot vol (1.05)
            _sheet.Cells[row, XLS_Table.R] = "Vmax (V)";
            _sheet.Cells[row, XLS_Table.S] = "Vmin spec (V)"; // undershoot vol (0.95)
            _sheet.Cells[row, XLS_Table.T] = "Vmin (V)";
            _sheet.Cells[row, XLS_Table.U] = "overshoot (%)";
            _sheet.Cells[row, XLS_Table.V] = "underhoot (%)";

            if (test_parameter.vidio.criteria[0].lpm_en)
            {
                _sheet.Cells[row, XLS_Table.O] = "Fall SR (V/s)";
                _sheet.Cells[row, XLS_Table.P] = "Fall Time (ms)";
                _sheet.Cells[row, XLS_Table.W] = "Delay time (LPM->vout)";
                _sheet.Cells[row, XLS_Table.X] = "Delay time (LPM->100%)";
                _sheet.Cells[row, XLS_Table.Y] = "Result";
                _range = _sheet.Range["C" + row, "Y" + row];
                _range.Interior.Color = Color.FromArgb(124, 252, 0);
                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            else
            {
                _sheet.Cells[row, XLS_Table.W] = "Result";
                _range = _sheet.Range["C" + row, "W" + row];
                _range.Interior.Color = Color.FromArgb(124, 252, 0);
                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            }
            row++;
#endif

            for (int case_idx = 0; case_idx < test_parameter.vidio.vout_list.Count; case_idx++)
            {
                for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < test_parameter.IoutList.Count; iout_idx++)
                    {
                        InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[vin_idx]);
                        if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        {
                            // default rising to rising
                            InsControl._oscilloscope.SetDelayTime(meas_delay, 3, 1);
                        }

                        updateMain.UpdateProgressBar(++progress);
                        phase1_name.Clear();
                        phase2_name.Clear();
                        file_name = string.Format("Temp={0}_VIN={1}_IOUT={2}_Vout={3}_{4}",
                                                temp,
                                                test_parameter.VinList[vin_idx],
                                                test_parameter.IoutList[iout_idx],
                                                test_parameter.vidio.vout_list[case_idx],
                                                test_parameter.vidio.vout_list_af[case_idx]
                                                );

                        test_parameter.waveform_name = file_name;
#if Power_en
                        InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                        MyLib.Delay1ms(200);
#endif

#if Eload_en
                        MyLib.Switch_ELoadLevel(test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
#endif
                        InsControl._oscilloscope.SetAutoTrigger();

                        //double vout = test_parameter.vidio.vout_list[case_idx];
                        //double vout_af = test_parameter.vidio.vout_list_af[case_idx];

                        double vout = 0;
                        double vout_af = 0;
                        if (test_parameter.vidio.criteria[case_idx].lpm_en)
                            vout = 0;
                        else
                            vout = Convert.ToDouble(test_parameter.vidio.vout_list[case_idx]);
                        vout_af = Convert.ToDouble(test_parameter.vidio.vout_list_af[case_idx]);

                        bool rising_en = vout < vout_af ? true : false;
                        bool diff = Math.Abs(vout - vout_af) < 0.13 ? true : false;
#if Report_en
                        double vin = test_parameter.VinList[vin_idx];
                        double spec_hi = test_parameter.vidio.criteria[case_idx].spec_hi;
                        double spec_lo = test_parameter.vidio.criteria[case_idx].spec_lo;
                        double iout = test_parameter.IoutList[iout_idx];
                        //double rise_spec = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].rise_time);
                        //double sr_rise = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_rise);
                        //double fall_spec = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].fall_time);
                        //double sr_fall = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_fall);
                        double vmax = test_parameter.vidio.criteria[case_idx].overshoot;
                        double vmin = test_parameter.vidio.criteria[case_idx].undershoot;

                        _sheet.Cells[row, XLS_Table.C] = temp;
                        _sheet.Cells[row, XLS_Table.D] = "LINK";
                        _sheet.Cells[row, XLS_Table.E] = vin;
                        _sheet.Cells[row, XLS_Table.F] = vout + "->" + vout_af;
                        _sheet.Cells[row, XLS_Table.G] = vmin + "->" + vmax;
                        _sheet.Cells[row, XLS_Table.H] = iout;
                        _sheet.Cells[row, XLS_Table.I] = (string)test_parameter.vidio.criteria[case_idx].rise_time;
                        _sheet.Cells[row, XLS_Table.J] = (string)test_parameter.vidio.criteria[case_idx].sr_rise;
                        _sheet.Cells[row, XLS_Table.M] = (string)test_parameter.vidio.criteria[case_idx].fall_time;
                        _sheet.Cells[row, XLS_Table.N] = (string)test_parameter.vidio.criteria[case_idx].sr_fall;
                        _sheet.Cells[row, XLS_Table.Q] = vmax;
                        _sheet.Cells[row, XLS_Table.S] = vmin;
#endif

#if Report_en
                        // waveform 9:24
                        if (!SlewRate_Rise_Task(case_idx)) // Rise time and slew
                        {
                            _sheet.Cells[row, XLS_Table.C] = "Rise Trigger Fail";
                            _sheet.Cells[row, XLS_Table.C].font.color = Color.Red;
                            row++;
                            continue;
                        }
                        string slewrate_min = phase1_name[slewrate_list.IndexOf(slewrate_list.Min())];
                        _range = _sheet.Range["AA" + (wave_row + 2), "AI" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, slewrate_min);
                        double res = diff ? slewrate_list.Min() * Math.Pow(10, 6) : slewrate_list.Min();

                        _sheet.Cells[row, XLS_Table.K] = slewrate_list.Min() * Math.Pow(10, -3);
                        _sheet.Cells[row, XLS_Table.L] = rise_time_list.Max() * Math.Pow(10, 6);
                        if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        {
                            _sheet.Cells[row, XLS_Table.W] = delay_list.Max() * Math.Pow(10, 6);
                            _sheet.Cells[row, XLS_Table.X] = delay100_list.Max() * Math.Pow(10, 6);
                        }


                        if (!SlewRate_Rise_Task(case_idx, true))         // overshoot
                        {
                            _sheet.Cells[row, XLS_Table.C] = "Rise Persistent Trigger Fail";
                            _sheet.Cells[row, XLS_Table.C].font.color = Color.Red;
                            row++;
                            continue;
                        }
                        _sheet.Cells[row, XLS_Table.R] = vmax_list.Max();
                        string shoot_max = test_parameter.waveform_name + "_overshoot";
                        _range = _sheet.Range["AK" + (wave_row + 2), "AS" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, shoot_max);

                        _sheet.Cells[row, XLS_Table.U] = overshoot_list.Max() * 100;
                        // --------------------------------------------------------------------------------------------------------

                        if (!SlewRate_Fall_Task(case_idx))
                        {
                            _sheet.Cells[row, XLS_Table.C] = "Fall Trigger Fail";
                            _sheet.Cells[row, XLS_Table.C].font.color = Color.Red;
                            row++;
                            continue;
                        }
                        slewrate_min = phase2_name[slewrate_list.IndexOf(slewrate_list.Min())];
                        _range = _sheet.Range["AU" + (wave_row + 2), "BC" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, slewrate_min);
                        res = diff ? slewrate_list.Min() * Math.Pow(10, 6) : slewrate_list.Min();

                        _sheet.Cells[row, XLS_Table.O] = test_parameter.vidio.criteria[case_idx].lpm_en ? slewrate_list.Min() : slewrate_list.Min() * Math.Pow(10, -3);
                        _sheet.Cells[row, XLS_Table.P] = test_parameter.vidio.criteria[case_idx].lpm_en ? fall_time_list.Max() * Math.Pow(10, 3) : fall_time_list.Max() * Math.Pow(10, 6);

                        if (!SlewRate_Fall_Task(case_idx, true))
                        {
                            _sheet.Cells[row, XLS_Table.C] = "Fall Persistent Trigger Fail";
                            _sheet.Cells[row, XLS_Table.C].font.color = Color.Red;
                            row++;
                            continue;
                        }
                        _sheet.Cells[row, XLS_Table.T] = vmin_list.Min();
                        shoot_max = test_parameter.waveform_name + "_undershoot";
                        // past over/under-shoot max case
                        _range = _sheet.Range["BE" + (wave_row + 2), "BM" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, shoot_max);
                        _sheet.Cells[row, XLS_Table.V] = undershoot_list.Min() * 100;
                        //_sheet.Cells[row, XLS_Table.F] = vmin_list.Max() + "->" + vmax_list.Max();

                        if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        {
                            _range = _sheet.Range["BO" + (wave_row + 2), "BW" + (wave_row + 25)];
                            string name = test_parameter.waveform_name + "_delay";
                            MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, name);
                        }
#endif
                        //-----------------------------------------------------------------------------------------
#if Report_en

                        double vmax_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.R].Value);
                        double vmin_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.T].Value);
                        bool judge_vol = (vmax_res < vmax & vmin_res > vmin) ? true : false;
                        bool judge_sr = true;
                        bool judge_time = true;
                        bool judge_res = true;
                        bool judge_lpm = true;

                        // slew rate judge
                        double rise_sr, fall_sr, rise_res, fall_res;
                        double rise_time, fall_time, rise_time_res, fall_time_res;
                        rise_sr = fall_sr = rise_res = fall_res = 0;
                        rise_time = fall_time = rise_time_res = fall_time_res = 0;

                        if (test_parameter.vidio.criteria[case_idx].sr_time_jd)
                        {
                            rise_sr = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_rise);
                            fall_sr = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_fall);
                            rise_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.K].Value);
                            fall_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.O].Value);
                            // slew rate big is good.
                            judge_sr = (rise_res > rise_sr & fall_res > fall_sr) ? true : false;
                        }


                        if (test_parameter.vidio.criteria[case_idx].time_jd)
                        {
                            rise_time = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].rise_time);
                            fall_time = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].fall_time);

                            rise_time_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.L].Value);
                            fall_time_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.P].Value);
                            // rise and fall time small is good.
                            judge_time = (rise_time_res < rise_time & fall_time_res < fall_time) ? true : false;
                        }
                        judge_res = judge_vol & judge_sr & judge_time;

                        if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        {
                            double delay = Convert.ToDouble(_sheet.Cells[row, XLS_Table.W].Value);
                            double delay100 = Convert.ToDouble(_sheet.Cells[row, XLS_Table.X].Value);

                            judge_lpm = ((delay100 + delay) < 120) & (delay < 20);
                            judge_res = judge_res & judge_lpm;

                            _sheet.Cells[row, XLS_Table.Y] = judge_res ? "Pass" : "Fail";
                            //_range.Interior.Color = judge_res ? Color.LightGreen : Color.LightPink;

                        }
                        else
                        {
                            _sheet.Cells[row, XLS_Table.W] = judge_res ? "Pass" : "Fail";
                            //_range = _sheet.Cells[row, XLS_Table.W];
                            //_range.Interior.Color = judge_res ? Color.White : Color.LightPink;
                        }



                        //if (test_parameter.vidio.criteria[case_idx].sr_time_jd)
                        //{
                        //    // slew rate judege
                        //    //_sheet.Cells[row, XLS_Table.I] = (string)test_parameter.vidio.criteria[case_idx].rise_time;
                        //    //_sheet.Cells[row, XLS_Table.J] = (string)test_parameter.vidio.criteria[case_idx].sr_rise;
                        //    //_sheet.Cells[row, XLS_Table.M] = (string)test_parameter.vidio.criteria[case_idx].fall_time;
                        //    //_sheet.Cells[row, XLS_Table.N] = (string)test_parameter.vidio.criteria[case_idx].sr_fall;
                        //    bool judge = judge_sr & judge_vol;
                        //    if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        //    {
                        //        double delay = Convert.ToDouble(_sheet.Cells[row, XLS_Table.W].Value);
                        //        double delay100 = Convert.ToDouble(_sheet.Cells[row, XLS_Table.X].Value);
                        //        judge = judge_sr & (delay100 < 120) & (delay < 20);
                        //        _range = _sheet.Cells[row, XLS_Table.Y];
                        //        _sheet.Cells[row, XLS_Table.Y] = judge ? "Pass" : "Fail";
                        //        _range.Interior.Color = judge ? Color.LightGreen : Color.LightPink;
                        //    }
                        //    else
                        //    {
                        //        _range = _sheet.Cells[row, XLS_Table.W];
                        //        _sheet.Cells[row, XLS_Table.W] = judge ? "Pass" : "Fail";
                        //        _range.Interior.Color = judge ? Color.LightGreen : Color.LightPink;
                        //    }
                        //}
                        //else
                        //{
                        //    // rise / fall judege
                        //    double rise_time = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].rise_time);
                        //    double fall_time = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].fall_time);
                        //    double rise_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.L].Value);
                        //    double fall_res = Convert.ToDouble(_sheet.Cells[row, XLS_Table.P].Value);
                        //    bool judge_time = (rise_res > rise_time | fall_res > fall_time) ? false : true;

                        //    bool judge = judge_time & judge_vol;

                        //    if (test_parameter.vidio.criteria[case_idx].lpm_en)
                        //    {
                        //        double delay = Convert.ToDouble((string)_sheet.Cells[row, XLS_Table.W].Value);
                        //        double delay100 = Convert.ToDouble((string)_sheet.Cells[row, XLS_Table.X].Value);

                        //        judge = judge_time & (delay100 < 120) & (delay < 20);
                        //        _range = _sheet.Cells[row, XLS_Table.Y];
                        //        _sheet.Cells[row, XLS_Table.Y] = judge ? "Pass" : "Fail";
                        //        _range.Interior.Color = judge ? Color.LightGreen : Color.LightPink;
                        //    }
                        //    else
                        //    {
                        //        _range = _sheet.Cells[row, XLS_Table.W];
                        //        _sheet.Cells[row, XLS_Table.W] = judge ? "Pass" : "Fail";
                        //        _range.Interior.Color = judge ? Color.LightGreen : Color.LightPink;
                        //    }
                        //}
                        //    _sheet.Cells[row, XLS_Table.N] = (rise > 20) | (fall > 20) ? "Pass" : "Fail";
                        //    _range.Interior.Color = (rise > 20) | (fall > 20) ? Color.LightGreen : Color.LightPink;

                        Excel.Range main_range = _sheet.Range["D" + row];
                        Excel.Range hyper = _sheet.Range["AA" + (wave_row + 1)];
                        // A to B
                        _sheet.Hyperlinks.Add(main_range, "#'" + _sheet.Name + "'!AA" + (wave_row + 1));
                        _sheet.Hyperlinks.Add(hyper, "#'" + _sheet.Name + "'!D" + row);

                        _sheet.Cells[wave_row, XLS_Table.AA] = "超連結";
                        _sheet.Cells[wave_row, XLS_Table.AB] = "VIN";
                        _sheet.Cells[wave_row, XLS_Table.AC] = "Vout";
                        _sheet.Cells[wave_row, XLS_Table.AD] = "Iout";
                        _sheet.Cells[wave_row, XLS_Table.AE] = "Rise (us)";
                        _sheet.Cells[wave_row, XLS_Table.AF] = "Rise SR (us/V)";
                        _sheet.Cells[wave_row, XLS_Table.AG] = "Fall (us)";
                        _sheet.Cells[wave_row, XLS_Table.AH] = "Fall SR (us/V)";
                        _sheet.Cells[wave_row, XLS_Table.AI] = "VMax (V)";
                        _sheet.Cells[wave_row, XLS_Table.AJ] = "VMin (V)";
                        _range = _sheet.Range["AA" + wave_row, "AJ" + wave_row];
                        _range.Interior.Color = Color.FromArgb(124, 252, 0);
                        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                        _sheet.Cells[wave_row + 1, XLS_Table.AA] = "Go back";
                        _sheet.Cells[wave_row + 1, XLS_Table.AB] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[wave_row + 1, XLS_Table.AC] = test_parameter.vidio.vout_list[case_idx] + "->" + test_parameter.vidio.vout_list_af[case_idx];
                        _sheet.Cells[wave_row + 1, XLS_Table.AD] = test_parameter.IoutList[iout_idx];
                        _sheet.Cells[wave_row + 1, XLS_Table.AE] = "=L" + row; // tise time
                        _sheet.Cells[wave_row + 1, XLS_Table.AF] = "=K" + row; // rise slew rate
                        _sheet.Cells[wave_row + 1, XLS_Table.AG] = "=P" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.AH] = "=O" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.AI] = "=R" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.AJ] = "=T" + row;
#endif
                        InsControl._oscilloscope.SetAutoTrigger();
                        wave_row += 31;
                        row++;
                    }
                }
            }

            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
#if Report_en
            string conditions = (string)_sheet.Cells[2, XLS_Table.B].Value + "\r\n";
            conditions = conditions + time;
            _sheet.Cells[2, XLS_Table.B] = conditions;
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_VIDIO_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif


#if Power_en
            InsControl._power.AutoPowerOff();
#endif

#if Eload_en
            InsControl._eload.AllChannel_LoadOff();
#endif

        } // function end

    }
}

