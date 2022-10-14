using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.IO;

namespace OLEDLite
{
    public class OutputRipple : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        //Excel.Range _range;
        string eLoadInfo = "";
        string SwireInfo = "";
        RTBBControl RTDev = new RTBBControl();


        private void OSCInint()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();

            InsControl._scope.CH1_On(); // Lx
            InsControl._scope.CH2_On(); // output

            InsControl._scope.CH2_ACoupling();
            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();

            InsControl._scope.CH1_Level(6);
            InsControl._scope.CH2_Level(1);

        }

        public override void ATETask()
        {

            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            int bin_cnt = 1;
            int wave_idx = 0;
            int row = 11;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length == 0 ? 1 : binList.Length;
            InsControl._power.AutoSelPowerOn(test_parameter.vinList[0]);

            OSCInint();
            for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
            {
                for(int load_idx = 0; load_idx < test_parameter.ioutList.Count; load_idx++)
                {
                    for(int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++)
                    {
                        //if (test_parameter.run_stop == true) goto Stop;


                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        MyLib.EloadFixChannel();
                        MyLib.Switch_ELoadLevel(test_parameter.eload_iout[load_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.eload_iout[load_idx]);

                        // set trigger

                        // resize channel

                        // measure result & catch waveform


                    }
                }
            }

        }
    }
}
