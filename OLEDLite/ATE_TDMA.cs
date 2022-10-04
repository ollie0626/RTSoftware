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

    public interface ITask
    {
        void ATETask();
    }

    public class TaskRun : ITask
    {
        public double temp = 25;
        virtual public void ATETask()
        { }
    }

    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
    };

    public class ATE_TDMA : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        RTBBControl RTDev = new RTBBControl();

        private void OSCInint()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();

            InsControl._scope.TimeScale((1 / (test_parameter.Freq * 1000)) / 10);

            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);

            InsControl._scope.CH1_BWLimitOn();
            InsControl._scope.CH2_BWLimitOn();

            InsControl._scope.DoCommand(":MEASure:VAVG CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMIN CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VMAX CHANnel2");
            InsControl._scope.DoCommand(":MEASure:VBASE CHANnel1");
            InsControl._scope.DoCommand(":MEASure:VTOP CHANnel1");
        }

        private void OSCRest()
        {
            InsControl._scope.CH1_Level(5);
            InsControl._scope.CH2_Level(5);
            InsControl._scope.CH1_Offset(0);
            InsControl._scope.CH2_Offset(0);
            MyLib.WaveformCheck();
        }

        // VinH : Hi target
        // VinL : Lo target
        private void ViResize(double VinH, double VinL)
        {
            double margin = 0.02;
            double Offset = 0.005;
            double meas_VinHi, meas_VinLo;
            int VinH_cnt = 0, VinL_cnt = 0;

            InsControl._scope.Root_RUN();
            InsControl._scope.Root_Clear();
            MyLib.Delay1s(1);
            InsControl._scope.Root_RUN();
            InsControl._scope.CH1_Level((VinH - VinL) / 3);
            InsControl._scope.CH1_Offset(VinH);
            InsControl._scope.DoCommand(":MEASure:STATistics MEAN");
            string[] res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
            InsControl._scope.Trigger_CH1();
            InsControl._scope.TriggerLevel_CH1((VinH + VinL) / 2);
            InsControl._scope.SetTriggerMode("TIMeout");
            InsControl._scope.SetTimeoutSource(1);
            InsControl._scope.SetTimeoutTime(((1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100) * 0.8) * Math.Pow(10, 9));

            // measure real VinHi, VinLo
            meas_VinHi = Convert.ToDouble(res[0]);
            meas_VinLo = Convert.ToDouble(res[1]);
            double VinH_in = VinH, VinL_in = VinL;


            while (meas_VinHi < (VinH - margin) || meas_VinHi > (VinH + margin))
            {
                InsControl._scope.Root_Clear();
                InsControl._scope.Root_RUN();
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                meas_VinHi = Convert.ToDouble(res[0]);

                if (meas_VinHi < (VinH - margin))
                {
                    VinH_in += Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL);
                }

                if (meas_VinHi > (VinH + margin))
                {
                    VinH_in -= Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL);
                }
                // out of loop
                VinH_cnt++;
                if (VinH_cnt > 40) break;
            }


            while (meas_VinLo < (VinL - margin) || meas_VinLo > (VinL + margin))
            {
                InsControl._scope.Root_Clear();
                InsControl._scope.Root_RUN();
                res = InsControl._scope.doQeury(":MEASure:RESults?").Split(',');
                meas_VinLo = Convert.ToDouble(res[1]);

                if (meas_VinLo < (VinL - margin))
                {
                    VinL_in += Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL_in);
                }

                if (meas_VinLo > (VinL + margin))
                {
                    VinL_in -= Offset;
                    MyLib.FuncGen_loopparameter(VinH_in, VinL_in);
                }
                // out of loop
                VinL_cnt++;
                if (VinL_cnt > 40) break;
            }
        }

        private void VoResize()
        {
            double Vo_offset = InsControl._scope.Meas_CH2AVG();

            InsControl._scope.CH2_Offset(Vo_offset);
            InsControl._scope.CH2_Level(0.1); // 100mV
            MyLib.WaveformCheck();

            for (int i = 0; i < 3; i++)
            {
                Vo_offset = InsControl._scope.Meas_CH2AVG();
                InsControl._scope.CH2_Offset(Vo_offset);
            }

            double abs_vo = Math.Abs(Vo_offset);
            if (abs_vo < 5) InsControl._scope.CH2_Level(0.01);
            else if (5 < abs_vo && abs_vo < 10) InsControl._scope.CH2_Level(0.02);
            else InsControl._scope.CH2_Level(0.1);
        }

        private void ExcelInit()
        {
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Name = "TDMA";

            _sheet.Cells[10, XLS_Table.K] = "No.";
            _sheet.Cells[10, XLS_Table.L] = "Temp(C)";
            _sheet.Cells[10, XLS_Table.M] = "Vin(V)";
            _sheet.Cells[10, XLS_Table.N] = "Iout(mA)";
            _sheet.Cells[10, XLS_Table.O] = "Overshoot(mV)";
            _sheet.Cells[10, XLS_Table.P] = "Undershoot(mV)";

        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();

            RTDev.BoadInit();
            int bin_cnt = 1;
            int row = 11;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[0].Highlevel + 0.5);
            MyLib.FuncGen_Fixedparameter(test_parameter.Freq * 1000,
                                         test_parameter.duty,
                                         test_parameter.tr,
                                         test_parameter.tf);
#if Report
            ExcelInit();
            _sheet.Cells[1, 1] = "Vin:";
            _sheet.Cells[2, 1] = "Iout:";
            _sheet.Cells[3, 1] = "setting conditions:";
            string res = "";
            for (int i = 0; i < test_parameter.HiLo_table.Count; i++)
                res += test_parameter.HiLo_table[i].Highlevel + "->" + test_parameter.HiLo_table[i].LowLevel + ", ";
            _sheet.Cells[1, 2] = res;
            res = "";
            for (int i = 0; i < test_parameter.ioutList.Count; i++)
                res += test_parameter.ioutList[i] + ", ";
            _sheet.Cells[2, 2] = res;
            _sheet.Cells[3, 2] = test_parameter.i2c_enable ? binList.Length : test_parameter.swireList.Count;
#endif
            OSCInint();
            for (int func_idx = 0; func_idx < test_parameter.HiLo_table.Count; func_idx++) // functino gen vin 
            {
                for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                {
                    for (int interface_idx = 0; interface_idx < (test_parameter.i2c_enable ? bin_cnt : test_parameter.swireList.Count); interface_idx++) // interface
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        res = Path.GetFileNameWithoutExtension(binList[interface_idx]);
                        string file_name = string.Format("{0}_{1}_Temp={2}C_Line={3:0.##}V->{4:0.##}_iout={5:0.##}A",
                                                        row - 11, res, temp,
                                                        test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel,
                                                        test_parameter.ioutList[iout_idx]);
                        if ((func_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();


                        OSCRest();
                        InsControl._power.AutoSelPowerOn(test_parameter.HiLo_table[func_idx].Highlevel + 0.5);
                        MyLib.FuncGen_loopparameter(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);

                        if (test_parameter.i2c_enable)
                        {
                            // i2c interface
                            RTDev.I2C_WriteBin(test_parameter.slave, 0x00, binList[interface_idx]);
                        }
                        else
                        {
                            // swire
                            int[] pulse_tmp = test_parameter.swireList[interface_idx].Split(',').Select(int.Parse).ToArray();
                            for (int pulse_idx = 0; pulse_idx < pulse_tmp.Length; pulse_idx++) RTDev.SwirePulse(pulse_tmp[pulse_idx]);
                        }

                        MyLib.EloadFixChannel();
                        ViResize(test_parameter.HiLo_table[func_idx].Highlevel, test_parameter.HiLo_table[func_idx].LowLevel);
                        VoResize();

                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);

                        // measure part
                        double zoomout_peak = InsControl._scope.Meas_CH2MAX();
                        double zoomout_neg_peak = InsControl._scope.Meas_CH2MIN();
                        double on_time = (1 / (test_parameter.Freq * 1000)) * (test_parameter.duty / 100);
                        double off_time = (1 / (test_parameter.Freq * 1000)) * ((100 - test_parameter.duty) / 100);

                        InsControl._scope.TimeScale(on_time / 20);
                        MyLib.Delay1ms(250);
                        double hi_peak = InsControl._scope.Meas_CH2MAX();
                        double hi_neg_peak = InsControl._scope.Meas_CH2MIN();
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_Rising");

                        InsControl._scope.TimeScale(on_time / 10);
                        InsControl._scope.TimeBasePosition(on_time);
                        MyLib.Delay1ms(250);
                        double lo_peak = InsControl._scope.Meas_CH2MAX();
                        double lo_neg_peak = InsControl._scope.Meas_CH2MIN();
                        InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name + "_Falling");

                        // power off
                        InsControl._funcgen.CH1_Off();

                        // report
                        double[] overshoot_list = new double[] { hi_peak, lo_peak };
                        double[] undershoot_list = new double[] { hi_neg_peak, lo_neg_peak };
#if Report
                        _sheet.Cells[row, XLS_Table.K] = row - 11;
                        _sheet.Cells[row, XLS_Table.L] = temp;
                        _sheet.Cells[row, XLS_Table.M] = test_parameter.HiLo_table[func_idx].Highlevel.ToString() + "->" + test_parameter.HiLo_table[func_idx].LowLevel.ToString();
                        _sheet.Cells[row, XLS_Table.N] = test_parameter.ioutList[iout_idx];
                        _sheet.Cells[row, XLS_Table.O] = Math.Abs(zoomout_peak - overshoot_list.Max());
                        _sheet.Cells[row, XLS_Table.P] = Math.Abs(zoomout_neg_peak - undershoot_list.Min());
#endif
                        row++;
                    }
                }
            
            }
        Stop:
            stopWatch.Stop();

#if Report
            TimeSpan timeSpan = stopWatch.Elapsed;
            MyLib.SaveExcelReport(test_parameter.wave_path, temp + "C_TDMA" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
            System.Windows.Forms.MessageBox.Show("Test finished!!!", "OLED Lite", System.Windows.Forms.MessageBoxButtons.OK);
        }



    }
}
