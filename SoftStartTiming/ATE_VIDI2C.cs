
//#define Report_en

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoftStartTiming
{
    public class ATE_VIDI2C : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public new double temp;
        RTBBControl RTDev = new RTBBControl();

        const double level_scale_div = 5;
        const double time_scale_div = 4;

        const int EN = 0;
        //const int Reset = 1;

        private void OSCInit()
        {
            InsControl._oscilloscope.CHx_On(1); // vout
            InsControl._oscilloscope.CHx_On(2); // Lx
            InsControl._oscilloscope.CHx_On(3); // vin
            InsControl._oscilloscope.CHx_On(4); // ILx

            InsControl._oscilloscope.CHx_Level(1, test_parameter.vidi2c.vout_des[0] / 4);
            InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[0] / 2);
            InsControl._oscilloscope.CHx_Level(3, test_parameter.VinList[0] / 2);

            InsControl._oscilloscope.CHx_Position(1, -2); // vout
            InsControl._oscilloscope.CHx_Position(2, -3); // Lx
            InsControl._oscilloscope.CHx_Position(3, 2);  // vin
            InsControl._oscilloscope.CHx_Position(4, -3); // iLx

            InsControl._oscilloscope.SetTimeBasePosition(27);

            // initial time scale
            InsControl._oscilloscope.SetTimeScale(500 * Math.Pow(10, -6));
            InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
            InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
            InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
        }

        private void I2CSetting(int data, int vout_idx)
        {
            byte data_msb = (byte)((data & 0xff00) >> 8);
            byte data_lsb = (byte)(data & 0xff);

            // i2c change vout
            RTDev.I2C_Write(
                test_parameter.slave,
                test_parameter.vidi2c.addr[vout_idx],
                test_parameter.vidi2c._2byte_en ? new byte[] { data_msb, data_lsb } : new byte[] { (byte)data }
                );


            if(test_parameter.vidi2c.addr_update == test_parameter.vidi2c.addr[vout_idx])
            {
                RTDev.I2C_Write(
                    test_parameter.slave,
                    test_parameter.vidi2c.addr_update,
                    new byte[] { (byte)(test_parameter.vidi2c.data_update | data_msb) }
                    );
            }
            else if(test_parameter.vidi2c.addr_update == test_parameter.vidi2c.addr[vout_idx] + 1)
            {
                RTDev.I2C_Write(
                    test_parameter.slave,
                    test_parameter.vidi2c.addr_update,
                    new byte[] { (byte)(test_parameter.vidi2c.data_update | data_lsb) }
                    );
            }
            else
            {
                // i2c update vout register
                RTDev.I2C_Write(
                    test_parameter.slave,
                    test_parameter.vidi2c.addr_update,
                    new byte[] { (byte)(test_parameter.vidi2c.data_update) }
                    );
            }
        }

        private void IOStateSetting(int en)
        {
            int value = (en << 0);
            int mask = 1 << EN;
            RTDev.GPIOnState((uint)mask, (uint)value);
        }

        private void CursorAdjust(bool rising_en)
        {
            InsControl._oscilloscope.SetREFLevelMethod(1);
            InsControl._oscilloscope.SetREFLevel(100, 50, 2, 1);
            MyLib.Delay1ms(100);
            double us_unit = Math.Pow(10, -6);
            double[] time_table = new double[] { 500 * us_unit, 400 * us_unit, 250 * us_unit, 200 * us_unit, 100 * us_unit, 40 * us_unit, 20 * us_unit};
            double x1 = 0, x2 = 0;
            List<double> min_list = new List<double>();
            if (rising_en)
            {
                InsControl._oscilloscope.CHx_Meas_Rise(1);
                double rise_time = InsControl._oscilloscope.CHx_Meas_Rise(1);
                rise_time = InsControl._oscilloscope.CHx_Meas_Rise(1);
                MyLib.Delay1ms(100);
                rise_time = InsControl._oscilloscope.CHx_Meas_Rise(1);
                double time_scale = rise_time / time_scale_div;
                for(int idx = 0; idx < time_table.Length; idx++)
                {
                    min_list.Add(Math.Abs(time_table[idx] - time_scale));
                }
                double min = min_list.Min();
                int min_idx = min_list.IndexOf(min);
                InsControl._oscilloscope.SetTimeScale(time_table[min_idx]);
                MyLib.Delay1ms(100);
            }
            else
            {
                InsControl._oscilloscope.CHx_Meas_Fall(1);
                double fall_time = InsControl._oscilloscope.CHx_Meas_Fall(1);
                fall_time = InsControl._oscilloscope.CHx_Meas_Fall(1);
                MyLib.Delay1ms(100);
                fall_time = InsControl._oscilloscope.CHx_Meas_Fall(1);
                double time_scale = fall_time / time_scale_div;
                for (int idx = 0; idx < time_table.Length; idx++)
                {
                    min_list.Add(Math.Abs(time_table[idx] - time_scale));
                }
                double min = min_list.Min();
                int min_idx = min_list.IndexOf(min);
                InsControl._oscilloscope.SetTimeScale(time_table[min_idx]);
                MyLib.Delay1ms(100);
            }

            x1 = InsControl._oscilloscope.GetAnnotationXn(1);
            x2 = InsControl._oscilloscope.GetAnnotationXn(2);
            InsControl._oscilloscope.SetCursorVPos(x1, x2);
            MyLib.Delay1ms(200);

        }

        private void PhaseTest(int vout_idx, bool rising_en)
        {
            double vout = test_parameter.vidi2c.vout_des[vout_idx];
            double vout_af = test_parameter.vidi2c.vout_des_af[vout_idx];
            int vout_data = test_parameter.vidi2c.vout_data[vout_idx];
            int vout_data_af = test_parameter.vidi2c.vout_data_af[vout_idx];
            double trigger_level = (vout_af > vout) ? vout_af - (vout_af - vout) * 0.5 : vout - (vout - vout_af) * 0.5;
            double ch_offset = (vout > vout_af) ? vout_af : vout;
            double ch_level = Math.Abs(vout - vout_af) / level_scale_div;

            if (rising_en)
            {
                // do rising event
                InsControl._oscilloscope.SetTriggerRise();
                InsControl._oscilloscope.CHx_Level(1, ch_level);
                InsControl._oscilloscope.CHx_Offset(1, ch_offset);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, trigger_level);
                MyLib.Delay1ms(500);
                for (int idx = 0; idx < 3; idx++)
                {
                    // initial state setting
                    IOStateSetting(1); // en
                    I2CSetting(vout_data, vout_idx);
                    MyLib.Delay1ms(500);
                    IOStateSetting(0); // en
                    MyLib.Delay1ms(100);
                    IOStateSetting(1); // en
                    InsControl._oscilloscope.SetRun();
                    MyLib.Delay1ms(300);
                    InsControl._oscilloscope.SetNormalTrigger();
                    InsControl._oscilloscope.SetClear();
                    MyLib.Delay1ms(300);
                    I2CSetting(vout_data_af, vout_idx);
                    MyLib.Delay1ms(500);
                    CursorAdjust(rising_en);
                    MyLib.Delay1ms(300);
                }
            }
            else
            {
                // do falling event
                InsControl._oscilloscope.SetTriggerFall();
                InsControl._oscilloscope.CHx_Level(1, ch_level);
                InsControl._oscilloscope.CHx_Offset(1, ch_offset);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, trigger_level);
                for (int idx = 0; idx < 3; idx++)
                {
                    // initial state setting
                    IOStateSetting(1); // en
                    I2CSetting(vout_data_af, vout_idx);
                    IOStateSetting(0); // en
                    IOStateSetting(1); // en
                    InsControl._oscilloscope.SetRun();
                    MyLib.Delay1ms(100);
                    InsControl._oscilloscope.SetNormalTrigger();
                    InsControl._oscilloscope.SetClear();
                    I2CSetting(vout_data, vout_idx);
                    CursorAdjust(rising_en);
                    MyLib.Delay1ms(300);
                }
            }

        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            OSCInit();
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Cells.Font.Name = "Calibri";
            _sheet.Cells.Font.Size = 11;
#endif

            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                for (int iout_idx = 0; iout_idx < test_parameter.IoutList.Count; iout_idx++)
                {
                    for(int freq_idx = 0; freq_idx < test_parameter.vidi2c.freq_data.Count; freq_idx++)
                    {
                        for(int vout_idx = 0; vout_idx < test_parameter.vidi2c.vout_data.Count; vout_idx++)
                        {
                            double vout = 0, vout_af = 0;
                            vout = test_parameter.vidi2c.vout_des[vout_idx];
                            vout_af = test_parameter.vidi2c.vout_des_af[vout_idx];

                            InsControl._oscilloscope.SetTimeScale(500 * Math.Pow(10, -6));
                            InsControl._oscilloscope.DoCommand("HORizontal:ROLL OFF");
                            InsControl._oscilloscope.DoCommand("HORizontal:MODE AUTO");
                            InsControl._oscilloscope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");

                            InsControl._oscilloscope.SetAutoTrigger();
                            InsControl._oscilloscope.CHx_Level(2, test_parameter.VinList[vin_idx] / 3);
                            InsControl._oscilloscope.CHx_Level(3, test_parameter.VinList[vin_idx] / 3);


                            RTDev.I2C_Write(test_parameter.slave, 
                                test_parameter.vidi2c.freq_addr, 
                                new byte[] { test_parameter.vidi2c.freq_data[freq_idx] });

                            //string freq = test_parameter.vidi2c.freq_list[freq_idx];
                            bool rising_en = vout_af > vout ? true : false;

                            // phase 1 test
                            PhaseTest(vout_idx, rising_en);

                            // phase 2 test
                            PhaseTest(vout_idx, !rising_en);

                        } // vout loop
                    } // freq loop
                } // iout loop
            } // vin loop

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

        }

    }
}
