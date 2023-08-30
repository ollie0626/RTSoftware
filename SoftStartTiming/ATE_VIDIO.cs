
#define Report_en
//#define Power_en
//#define Eload_en

using RTBBLibDotNet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.CompilerServices;
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

        List<double> vmax_list = new List<double>();
        List<double> vmin_list = new List<double>();
        List<string> phase1_name = new List<string>();
        List<string> phase2_name = new List<string>();

        int meas_rising = 1;
        int meas_falling = 2;
        int meas_vmax = 3;
        int meas_vmin = 4;


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

        private void CursorAdjust(int case_idx)
        {
            //double vout = test_parameter.vidio.vout_list[case_idx];
            //double vout_af = test_parameter.vidio.vout_list_af[case_idx];

            double vout = 0;
            double vout_af = 0;


            try
            {
                vout = Convert.ToDouble(test_parameter.vidio.vout_list[case_idx]);
            }
            catch
            {
                vout = 0;
            }

            try
            {
                vout_af = Convert.ToDouble(test_parameter.vidio.vout_list_af[case_idx]);
            }
            catch
            {
                vout_af = 0;
            }

            bool diff = Math.Abs(vout - vout_af) > 0.13 ? true : false;
            bool rising_en = vout < vout_af ? true : false;
            diff = (vout == 0 || vout_af == 0) ? false : true;

            double x1 = 0, x2 = 0;

            if (diff)
            {
                // > 130mV: 20% to 80%
                //InsControl._oscilloscope.SetREFLevelMethod(1);
                //InsControl._oscilloscope.SetREFLevel(80, 50, 20, 1);

                InsControl._oscilloscope.SetCursorMode();
                InsControl._oscilloscope.SetCursorWaveform();

                MyLib.Delay1ms(100);
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                MyLib.Delay1ms(100);
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                MyLib.Delay1ms(100);
                x2 = InsControl._oscilloscope.GetAnnotationXn(2);
                x2 = InsControl._oscilloscope.GetAnnotationXn(2);
                MyLib.Delay1ms(100);
                x2 = InsControl._oscilloscope.GetAnnotationXn(2);
                MyLib.Delay1ms(100);
            }
            else
            {
                // < 130mV: 0% to 100%
                // get 0% position
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                MyLib.Delay1ms(100);
                x1 = InsControl._oscilloscope.GetAnnotationXn(1);
                MyLib.Delay1ms(100);

                x2 = InsControl._oscilloscope.GetAnnotationXn(2);
                x2 = InsControl._oscilloscope.GetAnnotationXn(2);
                MyLib.Delay1ms(100);
                x2 = InsControl._oscilloscope.GetAnnotationXn(2);

                double high = rising_en ? vout_af : vout;
                double mid = Math.Abs(vout - vout_af) + (rising_en ? vout : vout_af);
                double low = rising_en ? vout : vout_af;
            }

            InsControl._oscilloscope.SetCursorMode();
            InsControl._oscilloscope.SetCursorOn();
            MyLib.Delay1ms(300);
            InsControl._oscilloscope.SetCursorSource(1, 1);
            InsControl._oscilloscope.SetCursorSource(2, 1);
            MyLib.Delay1ms(300);
            InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);
            MyLib.Delay1ms(100);
            InsControl._oscilloscope.SetCursorScreenYpos(diff ? vout * 0.8 : vout, diff ? vout_af * 0.2 : vout_af);
            MyLib.Delay1ms(100);
        }

        private void Initial_TimeScale(bool rising_en, bool LPM_en)
        {
            if (LPM_en && rising_en)
            {
                // LPM + rising
                InsControl._oscilloscope.SetTimeScale(100 * Math.Pow(10, -6));
                InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
            else if (LPM_en && !rising_en)
            {
                // LPM + falling
                // LPM + rising
                InsControl._oscilloscope.SetTimeScale(test_parameter.vidio.discharge_time);
                InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
            else
            {
                // normal mode
                InsControl._oscilloscope.SetTimeScale(5 * Math.Pow(10, -6));
                InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            }
        }

        private void ReferLevelSel(bool sel)
        {
            if (sel)
            {
                // > 130mV
                InsControl._oscilloscope.SetREFLevelMethod(1);
                InsControl._oscilloscope.SetREFLevel(80, 50, 20, 1);
            }
            else
            {
                // < 130mV
                InsControl._oscilloscope.SetREFLevelMethod(1);
                InsControl._oscilloscope.SetREFLevel(99, 50, 2, 1);
            }
        }


        private bool TriggerStatus()
        {
            int cnt = 0;
            while (InsControl._oscilloscope.GetCount() == 0)
            {
                cnt++;
                MyLib.Delay1ms(100);
                if (cnt > 100) return false;
            }
            return true;
        }

        private void Scope_Task_Setting(int meas_idx, double vout, double vout_af)
        {
            InsControl._oscilloscope.SetTimeOutTrigger();
            InsControl._oscilloscope.SetTimeOutTriggerCHx(1);
            InsControl._oscilloscope.SetTimeOutTime(5 * Math.Pow(10, -12));
            InsControl._oscilloscope.SetTimeOutEither();

            InsControl._oscilloscope.CHx_Level(1, (vout_af - vout) / 4.7);
            InsControl._oscilloscope.CHx_Offset(1, vout);
            InsControl._oscilloscope.CHx_Position(1, -2);
            InsControl._oscilloscope.SetTriggerLevel(1, (vout_af - vout) * 0.5 + vout);
            Initial_TimeScale(true, false);
            InsControl._oscilloscope.SetAnnotation(meas_idx);
        }

        private void SlewRate_Rise_Task(int case_idx, bool overshoot_en = false)
        {
            vmax_list.Clear();
            slewrate_list.Clear();
            rise_time_list.Clear();
            overshoot_list.Clear();

            double vout = 0;
            double vout_af = 0;
            vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);
            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);
            int intital_state = test_parameter.vidio.vout_map[vout.ToString()];
            int next_State = test_parameter.vidio.vout_map[vout_af.ToString()];
            Scope_Task_Setting(meas_rising, vout, vout_af); // trigger and time scale
            IOStateSetting(intital_state);

            double hi = test_parameter.vidio.criteria[case_idx].hi;
            double lo = test_parameter.vidio.criteria[case_idx].lo;

            // setting measure level threshold
            // example : (0.9 - 0.5) / 2 + 0.5 = 0.7
            InsControl._oscilloscope.SetREFLevelMethod(meas_rising, false);
            InsControl._oscilloscope.SetREFLevel(hi, lo + ((hi * lo) / 2), lo, meas_rising, false);
            InsControl._oscilloscope.SetTriggerRise();
            //InsControl._oscilloscope.SetCursorMode();
            InsControl._oscilloscope.SetCursorWaveform();
            InsControl._oscilloscope.SetCursorOn();
            

            if (overshoot_en)
            {
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(500);
            }

            for (int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double rise_time = 0;
                double vmax = 0;

            Trigger_Fail_retry:
                IOStateSetting(intital_state);
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetNormalTrigger();
                MyLib.Delay1ms(100);

                IOStateSetting(next_State);
                if (!TriggerStatus()) goto Trigger_Fail_retry;
                InsControl._oscilloscope.SetStop();
                if(repeat_idx == 0) MyLib.Delay1ms(200);
                // set cursor position
                InsControl._oscilloscope.SetAnnotation(meas_rising);
                MyLib.Delay1ms(50);
                double x1 = InsControl._oscilloscope.GetAnnotationXn(1); MyLib.Delay1ms(100);
                double x2 = InsControl._oscilloscope.GetAnnotationXn(2); MyLib.Delay1ms(100);

                InsControl._oscilloscope.SetCursorSource(1, 1);
                InsControl._oscilloscope.SetCursorSource(2, 1);
                InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);

                vmax = InsControl._oscilloscope.MeasureMean(meas_vmax);
                vmax_list.Add(vmax);

                // get delta T
                if(!overshoot_en)
                {
                    rise_time = InsControl._oscilloscope.GetCursorVBarDelta();
                    // slew rate delta V / delta T
                    slew_rate = rise_time / InsControl._oscilloscope.GetCursorHBarDelta();
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
                }

            }

            if (overshoot_en)
            {
                InsControl._oscilloscope.SetCursorOff();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetPERSistenceOff();
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + "_overshoot");
            }

        }

        private void SlewRate_Fall_Task(int case_idx, bool undershoot_en = false)
        {
            vmin_list.Clear();
            slewrate_list.Clear();
            fall_time_list.Clear();
            undershoot_list.Clear();

            double vout = 0;
            double vout_af = 0;
            vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);
            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);
            int inittal_state = test_parameter.vidio.vout_map[vout.ToString()];
            int next_state = test_parameter.vidio.vout_map[vout_af.ToString()];
            Scope_Task_Setting(meas_falling, vout, vout_af);

            IOStateSetting(next_state);
            InsControl._oscilloscope.SetTriggerFall();

            double hi = test_parameter.vidio.criteria[case_idx].hi;
            double lo = test_parameter.vidio.criteria[case_idx].lo;

            // setting measure level threshold
            // example : (0.9 - 0.5) / 2 + 0.5 = 0.7
            InsControl._oscilloscope.SetREFLevelMethod(meas_falling, false);
            InsControl._oscilloscope.SetREFLevel(hi, lo + ((hi * lo) / 2), lo, meas_falling, false);
            InsControl._oscilloscope.SetCursorWaveform();
            InsControl._oscilloscope.SetCursorOn();

            if (undershoot_en)
            {
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(500);
            }

            for(int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double fall_time = 0;
                double vmin = 0;

            Trigger_Fail_retry:
                IOStateSetting(next_state);
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetNormalTrigger();
                MyLib.Delay1ms(100);

                IOStateSetting(inittal_state);
                if (!TriggerStatus()) goto Trigger_Fail_retry;
                InsControl._oscilloscope.SetStop();
                if (repeat_idx == 0) MyLib.Delay1ms(200);

                // set cursor position
                InsControl._oscilloscope.SetAnnotation(meas_falling); MyLib.Delay1ms(50);
                double x1 = InsControl._oscilloscope.GetAnnotationXn(1); MyLib.Delay1ms(100);
                double x2 = InsControl._oscilloscope.GetAnnotationXn(2); MyLib.Delay1ms(100);
                //InsControl._oscilloscope.SetCursorMode();
                //InsControl._oscilloscope.SetCursorOn();
                InsControl._oscilloscope.SetCursorSource(1, 1);
                InsControl._oscilloscope.SetCursorSource(2, 1);
                InsControl._oscilloscope.SetCursorScreenXpos(x1, x2);

                vmin = InsControl._oscilloscope.MeasureMean(meas_vmax);
                vmin_list.Add(vmin);

                // get delta T
                if (!undershoot_en)
                {
                    fall_time = InsControl._oscilloscope.GetCursorVBarDelta();
                    // slew rate delta V / delta T
                    slew_rate = fall_time / InsControl._oscilloscope.GetCursorHBarDelta();
                    slewrate_list.Add(slew_rate);
                    fall_time_list.Add(fall_time);
                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, (repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_falling");
                    phase1_name.Add((repeat_idx).ToString() + "_" + test_parameter.waveform_name + "_falling");
                }
                else
                {
                    // measure undershoot
                    double res = (vmin - vout) / vout;
                    undershoot_list.Add(res);
                }
            }

            if (undershoot_en)
            {
                InsControl._oscilloscope.SetCursorOff();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetPERSistenceOff();
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + "_undershoot");
            }

        }

        private void Phase1Test(int case_idx, bool overshoot_en = false)
        {
            overshoot_list.Clear();
            undershoot_list.Clear();
            slewrate_list.Clear();
            vmin_list.Clear();
            vmax_list.Clear();

            double vout = 0;
            double vout_af = 0;

            string vout_str = "";
            string vout_af_str = "";

            vout = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_begin);
            vout_af = Convert.ToDouble(test_parameter.vidio.criteria[case_idx].vout_end);

            bool rising_en = vout < vout_af ? true : false;
            bool diff = Math.Abs(vout - vout_af) > 0.13 ? true : false;
            bool LPM_en = vout == 0 || vout_af == 0;
            ReferLevelSel(diff);

            // change edge trigger to timeOut trigger
            InsControl._oscilloscope.SetTimeOutTrigger();
            InsControl._oscilloscope.SetTimeOutTriggerCHx(1);
            InsControl._oscilloscope.SetTimeOutTime(5 * Math.Pow(10, -12));
            InsControl._oscilloscope.SetTimeOutEither();

            if (overshoot_en)
            {
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(1000);
            }

            if (rising_en)
            {
                InsControl._oscilloscope.CHx_Level(1, (vout_af - vout) / 4.7);
                InsControl._oscilloscope.CHx_Offset(1, vout);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout_af - vout) * 0.5 + vout);
                Initial_TimeScale(rising_en, LPM_en);
                InsControl._oscilloscope.SetAnnotation(meas_rising);
            }
            else
            {
                InsControl._oscilloscope.CHx_Level(1, (vout - vout_af) / 4.7);
                InsControl._oscilloscope.CHx_Offset(1, vout_af);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout - vout_af) * 0.5 + vout_af);
                Initial_TimeScale(rising_en, LPM_en);
                InsControl._oscilloscope.SetAnnotation(meas_falling);
            }

            for (int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double over_shoot = 0;
                double under_shoot = 0;
                double vmax = 0, vmin = 0;

                if (rising_en && LPM_en)
                {
                    // initial sate setting
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_str.ToString()]
                                    );
                    MyLib.Delay1ms(1000);
                }
                else
                {
                    // initial sate setting
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout.ToString()]
                                    );
                }

#if Eload_en
                if (LPM_en && !rising_en) InsControl._eload.CH1_Loading(test_parameter.vidio.discharge_load);
#endif
                Trigger_Fail_retry:
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetNormalTrigger();
                if (!overshoot_en) InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(100);

                if (rising_en && LPM_en)
                {
                    // transfer condition
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_af.ToString()]
                                );
                }
                else if (LPM_en)
                {
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_af_str.ToString()]
                                );
                }
                else
                {
                    // transfer condition
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_af.ToString()]
                                );
                }

                int cnt = 0;
                while (InsControl._oscilloscope.GetCount() == 0)
                {
                    cnt++;
                    MyLib.Delay1ms(100);
                    if (cnt > 100) goto Trigger_Fail_retry;
                }

                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetStop();
                MyLib.Delay1ms(100);

                if (rising_en)
                {
                    if (overshoot_en)
                    {
                        vmax = InsControl._oscilloscope.CHx_Meas_Max(1, meas_vmax);
                        over_shoot = (vmax - vout_af) / vout_af;
                        overshoot_list.Add(over_shoot * 100);
                        vmax_list.Add(vmax);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, meas_rising);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, meas_rising);
                    }

                    InsControl._oscilloscope.SetAnnotation(meas_rising);
                    MyLib.Delay1ms(100);
                    CursorAdjust(case_idx);
                    CursorAdjust(case_idx);

                    if (!diff)
                    {
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                    }
                    else
                    {
                        double time = InsControl._oscilloscope.GetCursorVBarDelta();
                        time = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        time = InsControl._oscilloscope.GetCursorVBarDelta();

                        double vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        MyLib.Delay1ms(100);
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();

                        slew_rate = vol / time;
                    }
                    if (!overshoot_en) slewrate_list.Add(!diff ? slew_rate : slew_rate * Math.Pow(10, -3));
                }
                else
                {
                    if (overshoot_en)
                    {
                        vmin = InsControl._oscilloscope.CHx_Meas_Min(1, meas_vmin);
                        under_shoot = (vout_af - vmin) / vout_af;
                        vmin_list.Add(vmin);
                        undershoot_list.Add(under_shoot);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, meas_falling);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, meas_falling);
                    }

                    InsControl._oscilloscope.SetAnnotation(meas_falling);
                    MyLib.Delay1ms(100);
                    CursorAdjust(case_idx);
                    CursorAdjust(case_idx);

                    if (!diff)
                    {
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                    }
                    else
                    {
                        double time = InsControl._oscilloscope.GetCursorVBarDelta();
                        time = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        time = InsControl._oscilloscope.GetCursorVBarDelta();

                        double vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        MyLib.Delay1ms(100);
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        slew_rate = vol / time;
                    }
                    if (!overshoot_en) slewrate_list.Add(!diff ? slew_rate : slew_rate * Math.Pow(10, -3));
                }

                // save every times wavefrom
                if (!overshoot_en)
                {
                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, (repeat_idx).ToString() + "_" + test_parameter.waveform_name + (rising_en ? "_rising" : "_falling"));
                    phase1_name.Add((repeat_idx).ToString() + "_" + test_parameter.waveform_name + (rising_en ? "_rising" : "_falling"));
                }

                if (LPM_en && !rising_en)
                {
#if Eload_en
                    InsControl._eload.LoadOFF(1);
#endif
                    //break;
                }
            }

            if (overshoot_en)
            {
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + (rising_en ? "_overshoot" : "_undershoot"));
                InsControl._oscilloscope.SetPERSistenceOff();
            }

        }

        private void Phase2Test(int case_idx, bool undershoot_en = false)
        {
            overshoot_list.Clear();
            undershoot_list.Clear();
            slewrate_list.Clear();
            vmin_list.Clear();
            vmax_list.Clear();

            double vout = 0;
            double vout_af = 0;

            string vout_str = "";
            string vout_af_str = "";

            try
            {
                vout = Convert.ToDouble(test_parameter.vidio.vout_list[case_idx]);
            }
            catch
            {
                vout_str = (string)test_parameter.vidio.vout_list[case_idx];
                vout = 0;
            }

            try
            {
                vout_af = Convert.ToDouble(test_parameter.vidio.vout_list_af[case_idx]);
            }
            catch
            {
                vout_af_str = (string)test_parameter.vidio.vout_list_af[case_idx];
                vout_af = 0;
            }

            bool rising_en = vout_af < vout ? true : false;
            bool diff = Math.Abs(vout - vout_af) > 0.13 ? true : false;
            bool LPM_en = vout == 0 || vout_af == 0;
            ReferLevelSel(diff);

            InsControl._oscilloscope.SetTimeOutTrigger();
            InsControl._oscilloscope.SetTimeOutTriggerCHx(1);
            InsControl._oscilloscope.SetTimeOutTime(5 * Math.Pow(10, -12));
            InsControl._oscilloscope.SetTimeOutEither();
            Initial_TimeScale(rising_en, LPM_en);

            if (undershoot_en)
            {
                InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(1000);
            }

            if (rising_en)
            {
                InsControl._oscilloscope.SetAnnotation(meas_rising);
            }
            else
            {
                InsControl._oscilloscope.SetAnnotation(meas_falling);
            }
            MyLib.Delay1ms(200);



            for (int repeat_idx = 0; repeat_idx < test_parameter.vidio.test_cnt; repeat_idx++)
            {
                double slew_rate = 0;
                double over_shoot = 0;
                double under_shoot = 0;
                double vmax = 0, vmin = 0;

                if (rising_en && LPM_en)
                {
                    // initial sate setting
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_af_str.ToString()]
                                    );
                    MyLib.Delay1ms(1000);
                }
                else
                {
                    // initial sate setting
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_af.ToString()]
                                    );
                    if (LPM_en) MyLib.Delay1ms(5000);
                }

#if Eload_en
                if (LPM_en && !rising_en) InsControl._eload.CH1_Loading(test_parameter.vidio.discharge_load);
#endif
                Trigger_Fail_retry:
                InsControl._oscilloscope.SetRun();
                if (LPM_en) MyLib.Delay1ms(1000);
                MyLib.Delay1ms(200);
                InsControl._oscilloscope.SetNormalTrigger();
                if (!undershoot_en) InsControl._oscilloscope.SetClear();
                MyLib.Delay1ms(200);
                if (LPM_en) MyLib.Delay1ms(1000);
                if (rising_en && LPM_en)
                {
                    // transfer condition
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_str.ToString()]
                                );
                }
                else if (LPM_en)
                {
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout_str.ToString()]
                            );
                    if (LPM_en) MyLib.Delay1s(4);
                }
                else
                {
                    // transfer condition
                    IOStateSetting(
                                    test_parameter.vidio.vout_map[vout.ToString()]
                                );
                }

                int cnt = 0;
                while (InsControl._oscilloscope.GetCount() == 0)
                {
                    cnt++;
                    MyLib.Delay1ms(100);
                    if (cnt > 100) goto Trigger_Fail_retry;
                }

                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetStop();
                MyLib.Delay1ms(100);

                if (rising_en)
                {
                    if (undershoot_en)
                    {
                        vmax = InsControl._oscilloscope.CHx_Meas_Max(1, meas_vmax);
                        vmax_list.Add(vmax);
                        over_shoot = (vmax - vout_af) / vout_af;
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, meas_rising);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Rise(1, meas_rising);
                    }

                    InsControl._oscilloscope.SetAnnotation(meas_rising);
                    MyLib.Delay1ms(100);
                    CursorAdjust(case_idx);
                    CursorAdjust(case_idx);

                    if (!diff)
                    {
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                    }
                    else
                    {
                        double time = InsControl._oscilloscope.GetCursorVBarDelta();
                        time = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        time = InsControl._oscilloscope.GetCursorVBarDelta();

                        double vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        MyLib.Delay1ms(100);
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();

                        slew_rate = vol / time;
                    }
                    if (!undershoot_en) slewrate_list.Add(!diff ? slew_rate : slew_rate * Math.Pow(10, -3));
                }
                else
                {
                    if (undershoot_en)
                    {
                        vmin = InsControl._oscilloscope.CHx_Meas_Min(1, meas_vmin);
                        vmin_list.Add(vmin);
                        under_shoot = (vout_af - vmin) / vout_af;
                        undershoot_list.Add(under_shoot);
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, meas_falling);
                    }
                    else
                    {
                        slew_rate = InsControl._oscilloscope.CHx_Meas_Fall(1, meas_falling);
                    }

                    InsControl._oscilloscope.SetAnnotation(meas_falling);
                    MyLib.Delay1ms(100);
                    CursorAdjust(case_idx);
                    CursorAdjust(case_idx);
                    if (!diff)
                    {
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        slew_rate = InsControl._oscilloscope.GetCursorVBarDelta();
                    }
                    else
                    {
                        double time = InsControl._oscilloscope.GetCursorVBarDelta();
                        time = InsControl._oscilloscope.GetCursorVBarDelta();
                        MyLib.Delay1ms(100);
                        time = InsControl._oscilloscope.GetCursorVBarDelta();

                        double vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();
                        MyLib.Delay1ms(100);
                        vol = InsControl._oscilloscope.GetCursorHBarDelta();

                        slew_rate = vol / time;
                    }

                    if (!undershoot_en) slewrate_list.Add(!diff ? slew_rate : slew_rate * Math.Pow(10, -3));
                }

                if (!undershoot_en)
                {
                    // save every times wavefrom
                    InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, (repeat_idx).ToString() + "_" + test_parameter.waveform_name + (rising_en ? "_rising" : "_falling"));
                    phase2_name.Add((repeat_idx).ToString() + "_" + test_parameter.waveform_name + (rising_en ? "_rising" : "_falling"));
                }

                if (LPM_en && !rising_en)
                {
#if Eload_en
                    InsControl._eload.LoadOFF(1);
#endif
                    //break;
                }
            }

            if (undershoot_en)
            {
                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, test_parameter.waveform_name + (rising_en ? "_overshoot" : "_undershoot"));
                InsControl._oscilloscope.SetPERSistenceOff();
            }
        }

        public override void ATETask()
        {
            progress = 0;
            updateMain.UpdateProgressBar(0);

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
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
            _sheet.Cells[row, XLS_Table.I] = "Rise SR spec (us/V)";
            _sheet.Cells[row, XLS_Table.J] = "Rise Time spec (us)";
            _sheet.Cells[row, XLS_Table.K] = "Rise SR (us/V)";
            _sheet.Cells[row, XLS_Table.L] = "Rise Time (us)";
            _sheet.Cells[row, XLS_Table.M] = "Fall SR spec (us/V)";
            _sheet.Cells[row, XLS_Table.N] = "Fall Time spec (us)";
            _sheet.Cells[row, XLS_Table.O] = "Fall SR (us/V)";
            _sheet.Cells[row, XLS_Table.P] = "Fall Time (us)";
            _sheet.Cells[row, XLS_Table.Q] = "Vmax spec (V)";
            _sheet.Cells[row, XLS_Table.R] = "Vmax (V)";
            _sheet.Cells[row, XLS_Table.S] = "Vmin spec (V)";
            _sheet.Cells[row, XLS_Table.T] = "Vmin (V)";
            _sheet.Cells[row, XLS_Table.U] = "overshoot (%)";
            _sheet.Cells[row, XLS_Table.V] = "underhoot (%)";
            _sheet.Cells[row, XLS_Table.W] = "Result";

            _range = _sheet.Range["C" + row, "W" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            row++;
#endif

            for (int case_idx = 0; case_idx < test_parameter.vidio.vout_list.Count; case_idx++)
            {
                for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < test_parameter.IoutList.Count; iout_idx++)
                    {
                        InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[vin_idx]);
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

                        vout = Convert.ToDouble(test_parameter.vidio.vout_list[case_idx]);
                        vout_af = Convert.ToDouble(test_parameter.vidio.vout_list_af[case_idx]);

                        bool rising_en = vout < vout_af ? true : false;
                        bool diff = Math.Abs(vout - vout_af) < 0.13 ? true : false;
#if Report_en
                        //_sheet.Cells[row, XLS_Table.C] = "Temp(C)";
                        //_sheet.Cells[row, XLS_Table.D] = "超連結";
                        //_sheet.Cells[row, XLS_Table.E] = "Vin (V)";
                        //_sheet.Cells[row, XLS_Table.F] = "Vout Change (V)";
                        //_sheet.Cells[row, XLS_Table.G] = "VID spec (V)";
                        //_sheet.Cells[row, XLS_Table.H] = "Iout (A)";
                        //_sheet.Cells[row, XLS_Table.I] = "Rise SR spec (us/V)";
                        //_sheet.Cells[row, XLS_Table.J] = "Rise Time spec (us)";
                        //_sheet.Cells[row, XLS_Table.K] = "Rise SR (us/V)";
                        //_sheet.Cells[row, XLS_Table.L] = "Rise Time (us)";
                        //_sheet.Cells[row, XLS_Table.M] = "Fall SR spec (us/V)";
                        //_sheet.Cells[row, XLS_Table.N] = "Fall Time spec (us)";
                        //_sheet.Cells[row, XLS_Table.O] = "Fall SR (us/V)";
                        //_sheet.Cells[row, XLS_Table.P] = "Fall Time (us)";
                        //_sheet.Cells[row, XLS_Table.Q] = "Vmax spec (V)";
                        //_sheet.Cells[row, XLS_Table.R] = "Vmax (V)";
                        //_sheet.Cells[row, XLS_Table.S] = "Vmin spec (V)";
                        //_sheet.Cells[row, XLS_Table.T] = "Vmin (V)";
                        //_sheet.Cells[row, XLS_Table.U] = "overshoot (%)";
                        //_sheet.Cells[row, XLS_Table.V] = "underhoot (%)";
                        //_sheet.Cells[row, XLS_Table.W] = "Result";

                        double vin = test_parameter.VinList[vin_idx];
                        double spec_hi = test_parameter.vidio.criteria[case_idx].spec_hi;
                        double spec_lo = test_parameter.vidio.criteria[case_idx].spec_lo;
                        double iout = test_parameter.IoutList[iout_idx];
                        double rise_spec = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].rise_time);
                        double sr_rise = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_rise);
                        double fall_spec = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].fall_time);
                        double sr_fall = Convert.ToDouble((string)test_parameter.vidio.criteria[case_idx].sr_fall);
                        double vmax = test_parameter.vidio.criteria[case_idx].overshoot;
                        double vmin = test_parameter.vidio.criteria[case_idx].undershoot;

                        _sheet.Cells[row, XLS_Table.C] = temp;
                        _sheet.Cells[row, XLS_Table.D] = "LINK";
                        _sheet.Cells[row, XLS_Table.E] = vin;
                        _sheet.Cells[row, XLS_Table.F] = vout + "->" + vout_af;
                        _sheet.Cells[row, XLS_Table.G] = spec_hi + "->" + spec_lo;
                        _sheet.Cells[row, XLS_Table.H] = iout;
                        _sheet.Cells[row, XLS_Table.I] = rise_spec;
                        _sheet.Cells[row, XLS_Table.J] = sr_rise;
                        _sheet.Cells[row, XLS_Table.I] = fall_spec;
                        _sheet.Cells[row, XLS_Table.J] = sr_fall;
                        _sheet.Cells[row, XLS_Table.Q] = vmax;
                        _sheet.Cells[row, XLS_Table.S] = vmin;
#endif

#if Report_en
                        // waveform 9:24
                        SlewRate_Rise_Task(case_idx);               // Rise time and slew
                        string slewrate_min = phase1_name[slewrate_list.IndexOf(slewrate_list.Min())];
                        _range = _sheet.Range["AA" + (wave_row + 2), "AI" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, slewrate_min);
                        double res = diff ? slewrate_list.Min() * Math.Pow(10, 6) : slewrate_list.Min();

                        _sheet.Cells[row, XLS_Table.K] = rise_time_list.Min();
                        _sheet.Cells[row, XLS_Table.L] = slewrate_list.Min();
                        _sheet.Cells[row, XLS_Table.R] = vmax_list.Max();

                        SlewRate_Rise_Task(case_idx, true);         // overshoot
                        string shoot_max = test_parameter.waveform_name + "_overshoot";
                        _range = _sheet.Range["AK" + (wave_row + 2), "AS" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, shoot_max);

                        _sheet.Cells[row, XLS_Table.U] = overshoot_list.Max();
                        // --------------------------------------------------------------------------------------------------------
                        
                        SlewRate_Fall_Task(case_idx);
                        slewrate_min = phase2_name[slewrate_list.IndexOf(slewrate_list.Min())];
                        _range = _sheet.Range["AU" + (wave_row + 2), "BC" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, slewrate_min);
                        res = diff ? slewrate_list.Min() * Math.Pow(10, 6) : slewrate_list.Min();

                        _sheet.Cells[row, XLS_Table.O] = fall_time_list.Min();
                        _sheet.Cells[row, XLS_Table.P] = slewrate_list.Min();
                        _sheet.Cells[row, XLS_Table.T] = vmin_list.Max();


                        SlewRate_Fall_Task(case_idx, true);
                        shoot_max = test_parameter.waveform_name + "_undershoot";
                        // past over/under-shoot max case
                        _range = _sheet.Range["BE" + (wave_row + 2), "BM" + (wave_row + 25)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, shoot_max);

                        _sheet.Cells[row, XLS_Table.V] = undershoot_list.Min();

#endif
                        //-----------------------------------------------------------------------------------------
#if Report_en

                        //    _sheet.Cells[row, XLS_Table.N] = (rise > 20) | (fall > 20) ? "Pass" : "Fail";
                        //    _range.Interior.Color = (rise > 20) | (fall > 20) ? Color.LightGreen : Color.LightPink;

                        Excel.Range main_range = _sheet.Range["D" + row];
                        Excel.Range hyper = _sheet.Range["AA" + (wave_row + 1)];
                        // A to B
                        _sheet.Hyperlinks.Add(main_range, "#'" + _sheet.Name + "'!Q" + (wave_row + 1));
                        _sheet.Hyperlinks.Add(hyper, "#'" + _sheet.Name + "'!D" + row);

                        _sheet.Cells[wave_row, XLS_Table.AA] = "超連結";
                        _sheet.Cells[wave_row, XLS_Table.AB] = "VIN";
                        _sheet.Cells[wave_row, XLS_Table.AC] = "Vout";
                        _sheet.Cells[wave_row, XLS_Table.AD] = "Iout";
                        _sheet.Cells[wave_row, XLS_Table.AE] = "Rise (us)";
                        _sheet.Cells[wave_row, XLS_Table.AF] = "Rise SR (us/V)";
                        _sheet.Cells[wave_row, XLS_Table.AG] = "Fall (us)";
                        _sheet.Cells[wave_row, XLS_Table.AH] = "Fall SR (us/V)";
                        _sheet.Cells[wave_row, XLS_Table.AI] = "Overshoot(%)";
                        _sheet.Cells[wave_row, XLS_Table.AJ] = "Undershoot(%)";
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
                        _sheet.Cells[wave_row + 1, XLS_Table.AI] = "=U" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.AJ] = "=V" + row;
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
