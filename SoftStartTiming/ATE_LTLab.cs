#define Report_en
#define Power_en
#define Eload_en

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;

namespace SoftStartTiming
{
    public class ATE_LTLab : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        RTBBControl RTDev = new RTBBControl();

        private bool sel;
        private int temp_meas = 8;

        private int meas_vout = 1;
        private int meas_vpp = 2;
        private int meas_vmax = 3;
        private int meas_vmin = 4;
        private int meas_imax = 5;
        private int meas_imin = 6;

        private void OSCInit()
        {
            InsControl._oscilloscope.CHx_On(1);                 // Vout
            InsControl._oscilloscope.CHx_On(2);                 // Lx
            InsControl._oscilloscope.CHx_On(4);                 // Iout

            InsControl._oscilloscope.CHx_BWLimitOn(1);
            InsControl._oscilloscope.CHx_BWLimitOn(2);
            InsControl._oscilloscope.CHx_BWLimitOn(3);
            InsControl._oscilloscope.CHx_BWLimitOn(4);

            InsControl._oscilloscope.SetTimeScale(Math.Pow(10, -6));
            InsControl._oscilloscope.SetTimeBasePosition(35);

            // channel position
            InsControl._oscilloscope.CHx_Level(1, 1);
            InsControl._oscilloscope.CHx_Level(2, 1);
            InsControl._oscilloscope.CHx_Level(4, 1);
            InsControl._oscilloscope.CHx_Position(1, 2.5);
            InsControl._oscilloscope.CHx_Position(2, -1.5);
            InsControl._oscilloscope.CHx_Position(4, -3);

            InsControl._oscilloscope.SetMeasureSource(1, meas_vpp, "PK2Pk");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmax, "MAXimum");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmin, "MINImum");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vout, "AMPlitude");

            InsControl._oscilloscope.SetMeasureSource(4, meas_imax, "MAXimum");
            InsControl._oscilloscope.SetMeasureSource(4, meas_imin, "MINImum");


            InsControl._funcgen.CH1_Off();
            InsControl._eload.AllChannel_LoadOff();
        }

        private bool I2C_Check(byte match)
        {
            byte addr = test_parameter.lt_lab.addr_list[0];
            byte data = test_parameter.lt_lab.data_list[0];
            byte[] buf = new byte[1];
            RTDev.I2C_Read(test_parameter.slave, addr, ref buf);
            return (buf[0] == match);
        }


        public override void ATETask()
        {
            RTDev.BoadInit();
            OSCInit();
            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                for(int i2c_idx = 0; i2c_idx < test_parameter.lt_lab.data_list.Count; i2c_idx++)
                {
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                    MyLib.Delay1ms(200);


                    while (I2C_Check(test_parameter.lt_lab.data_list[i2c_idx])) ;






                }
            }
        }
    }



}
