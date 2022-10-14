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
    public class ATE_OutputRipple : TaskRun
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
            int idx = 0;
            int bin_cnt = 1;
            int wave_idx = 0;
            int row = 11;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length == 0 ? 1 : binList.Length;
            double[] ori_vinTable = new double[test_parameter.vinList.Count];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, test_parameter.vinList.Count);
            InsControl._power.AutoSelPowerOn(test_parameter.vinList[0]);

            OSCInint();
            for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
            {
                for (int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++)
                {
                    for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    {
                        //if (test_parameter.run_stop == true) goto Stop;
                        if ((interface_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        string res = binList[interface_idx];
                        string file_name = string.Format("{0}_{1}_Temp={2}C_Vin={3:0.0#}V_{4:0.0#}A",
                                                        idx,
                                                        Path.GetFileNameWithoutExtension(res),
                                                        temp,
                                                        test_parameter.vinList[vin_idx],
                                                        test_parameter.ioutList[iout_idx]);

                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                        MyLib.EloadFixChannel();
                        MyLib.Switch_ELoadLevel(test_parameter.eload_iout[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.eload_iout[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];
                        if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                        {
                            System.Windows.Forms.MessageBox.Show("34970 沒有連結 !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                            return;
                        }
                        // set trigger
                        // adjust ch1 level
                        InsControl._scope.CH1_Level(1);
                        System.Threading.Thread.Sleep(500);
                        //InsControl._scope.CH1_Level(0.05);
                        //MyLib.Channel_LevelSetting(1);
                        System.Threading.Thread.Sleep(1000);
                        // scope open rgb color function
                        InsControl._scope.DoCommand(":DISPlay:PERSistence 5");
                        System.Threading.Thread.Sleep(5000);
                        double max, min, vpp, vin, vout, iin, iout;
                        // save waveform
                        InsControl._scope.Root_STOP();
                        // measure data
                        max = InsControl._scope.Meas_CH1MAX() * 1000;
                        min = InsControl._scope.Meas_CH1MIN() * 1000;
                        vpp = InsControl._scope.Meas_CH1VPP() * 1000;
                        vin = InsControl._34970A.Get_100Vol(1);
                        vout = InsControl._34970A.Get_100Vol(2);
                        iin = InsControl._power.GetCurrent();
                        iout = InsControl._eload.GetIout();



                    }
                }
            }

        }
    }
}
