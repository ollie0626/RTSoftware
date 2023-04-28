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

namespace SoftStartTiming
{
    public class ATE_VIDIO : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public double temp;
        RTBBControl RTDev = new RTBBControl();

        const int LPM = 1;
        const int G1 = 2;
        const int G2 = 4;


        private void IOStateSetting(int lpm, int g1, int g2)
        {
            int value = (lpm << 0 | g1 << 1 | g2 << 2);
            int mask = LPM << 0 | G1 << 1 | G2 << 2;
            RTDev.GPIOnState((uint)mask, (uint)value);
        }

        private void OSCInit()
        {
            InsControl._oscilloscope.CHx_On(1); // G1 or G2
            InsControl._oscilloscope.CHx_On(2); // Vout
            InsControl._oscilloscope.CHx_On(3); // Lx
            InsControl._oscilloscope.CHx_Off(4); // un-use channel
            
            // initial time scale
            InsControl._oscilloscope.SetTimeScale(4 * Math.Pow(10, -6));

            InsControl._oscilloscope.CHx_Level(1, 2);
            InsControl._oscilloscope.CHx_Position(1, 2.5);

            double max = test_parameter.vidio.vout_list[0] > test_parameter.vidio.vout_list_af[0] ?
                         test_parameter.vidio.vout_list[0] : test_parameter.vidio.vout_list_af[0];
            double min = test_parameter.vidio.vout_list[0] < test_parameter.vidio.vout_list_af[0] ?
                         test_parameter.vidio.vout_list[0] : test_parameter.vidio.vout_list_af[0];
            InsControl._oscilloscope.CHx_Level(2, max - min / 3);
            InsControl._oscilloscope.CHx_Offset(2, min);
            InsControl._oscilloscope.CHx_Position(2, -2);


            InsControl._oscilloscope.CHx_Level(3, test_parameter.VinList[0] / 1.5);
            InsControl._oscilloscope.CHx_Position(3, -3);

            InsControl._oscilloscope.SetAutoTrigger();
            InsControl._oscilloscope.SetTriggerLevel(2, max - min);
        }

        public override void ATETask()
        {
            RTDev.BoadInit();
            OSCInit();

            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                for(int iout_idx = 0; iout_idx < test_parameter.IoutList.Count; iout_idx++)
                {
                    for(int case_idx = 0; case_idx < test_parameter.vidio.g1_sel.Count; case_idx++)
                    {
                        InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                        MyLib.Delay1ms(200);
                        MyLib.Switch_ELoadLevel(test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);


                        double vout = test_parameter.voutList[case_idx];
                        double vout_af = test_parameter.voutList[case_idx];
                        bool rising_en = vout < vout_af ? true : false;


                        if (rising_en) InsControl._oscilloscope.SetTriggerRise();
                        else InsControl._oscilloscope.SetTriggerFall();
                        

                        // initial sate setting
                        IOStateSetting(
                                        test_parameter.vidio.lpm_sel[case_idx],
                                        test_parameter.vidio.g1_sel[case_idx],
                                        test_parameter.vidio.g2_sel[case_idx]
                                        );

                        MyLib.Delay1ms(500);

                        





                        // transfer condition
                        IOStateSetting(
                                        test_parameter.vidio.lpm_sel_af[case_idx],
                                        test_parameter.vidio.g1_sel_af[case_idx],
                                        test_parameter.vidio.g2_sel_af[case_idx]
                                        );


                        InsControl._oscilloscope.SetAutoTrigger();
                    }
                }
            }
        } // function end

    }
}
