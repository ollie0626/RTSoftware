
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

        const int EN = 0;
        const int Reset = 1;


        private void OSCInit()
        {
            InsControl._oscilloscope.CHx_On(1); // vout
            InsControl._oscilloscope.CHx_On(2); // Lx
            InsControl._oscilloscope.CHx_On(3); // vin
            InsControl._oscilloscope.CHx_On(4); // ILx
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

            // i2c update vout register
            RTDev.I2C_Write(
                test_parameter.slave,
                test_parameter.vidi2c.addr_update,
                new byte[] { test_parameter.vidi2c.data_update }
                );
        }

        private void IOStateSetting(int en, int reset)
        {
            int value = (en << 0 | reset << 1);
            int mask = 1 << EN | 1 << Reset;
            RTDev.GPIOnState((uint)mask, (uint)value);
        }

        private void PhaseTest(int vout_idx, bool rising_en)
        {

            double vout = test_parameter.vidi2c.vout_des[vout_idx];
            double vout_af = test_parameter.vidi2c.vout_des_af[vout_idx];
            int vout_data = test_parameter.vidi2c.vout_data[vout_idx];
            int vout_data_af = test_parameter.vidi2c.vout_data_af[vout_idx];

            if (rising_en)
            {
                InsControl._oscilloscope.SetTriggerRise();
                InsControl._oscilloscope.CHx_Level(1, (vout_af - vout) / 4.5);
                InsControl._oscilloscope.CHx_Offset(1, vout);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout_af - vout) * 0.3 + vout);

                // initial state setting
                IOStateSetting(1, 1); // en, reset
                I2CSetting(vout_data, vout_idx);
                IOStateSetting(1, 0); // en, reset
                IOStateSetting(1, 1); // en, reset
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                I2CSetting(vout_data_af, vout_idx);
            }
            else
            {
                InsControl._oscilloscope.SetTriggerFall();
                InsControl._oscilloscope.CHx_Level(1, (vout - vout_af) / 4.5);
                InsControl._oscilloscope.CHx_Offset(1, vout_af);
                InsControl._oscilloscope.CHx_Position(1, -2);
                InsControl._oscilloscope.SetTriggerLevel(1, (vout - vout_af) * 0.3 + vout_af);


                // initial state setting
                IOStateSetting(1, 1); // en, reset
                I2CSetting(vout_data_af, vout_idx);
                IOStateSetting(1, 0); // en, reset
                IOStateSetting(1, 1); // en, reset
                InsControl._oscilloscope.SetRun();
                MyLib.Delay1ms(100);
                InsControl._oscilloscope.SetNormalTrigger();
                InsControl._oscilloscope.SetClear();
                I2CSetting(vout_data, vout_idx);
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

                            InsControl._oscilloscope.SetAutoTrigger();
                            RTDev.I2C_Write(test_parameter.slave, 
                                test_parameter.vidi2c.freq_addr, 
                                new byte[] { test_parameter.vidi2c.freq_data[freq_idx] });

                            string freq = test_parameter.vidi2c.freq_list[freq_idx];
                            bool rising_en = vout_af > vout ? true : false;

                            PhaseTest(vout_idx, rising_en);
                            //PhaseTest(vout_idx, !rising_en);

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
