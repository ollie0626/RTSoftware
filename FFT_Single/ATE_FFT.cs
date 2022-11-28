using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using InsLibDotNet;
using RTBBLibDotNet;
using System.Drawing;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Globalization;

namespace FFT_Single
{
    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
    };


    public class ATE_FFT
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;


        public double temp;
        RTBBControl RTDev = new RTBBControl();

        private void func_gen_fixed_parameter(double freq, double duty)
        {
            InsControl._funcgen.CH1_ContinuousMode();
            InsControl._funcgen.CH1_PulseMode();
            MyLib.Delay1ms(500);
            InsControl._funcgen.CH1_Frequency(freq);
            MyLib.Delay1ms(500);
            InsControl._funcgen.CH1_DutyCycle(duty);
            InsControl._funcgen.CHl1_HiLevel(1.6);
        }


        private void FFT_Task(int row)
        {
            int channel = test_parameter.channel;
            InsControl._scope.Measure_Clear();
            InsControl._scope.Measure_Freq(channel);
            InsControl._scope.DoCommand(":MARKer:MODE OFF");
            System.Threading.Thread.Sleep(1000);
            InsControl._scope.Bandwidth_Limit_On(channel);
            InsControl._scope.Ch_On(1);
            InsControl._scope.TimeScaleUs(20);
            MyLib.Delay1ms(100);
            /* time scale setting */
            double freq = InsControl._scope.Measure_Freq(channel);
            double period = 1 / freq;
            double time_scale = period * 30;
            double ch_level = 20;
            InsControl._scope.TriggerLevel_CH1(15);
            InsControl._scope.Trigger_CH1();
            InsControl._scope.CHx_Level(channel, ch_level);
            InsControl._scope.Ch_Offset(channel, ch_level * 2);
            InsControl._scope.Trigger(channel);
            InsControl._scope.TimeScaleUs(time_scale);
            InsControl._scope.TimeBasePosition(0);
            MyLib.Delay1ms(500);
            double Vmax = InsControl._scope.Meas_CH1MAX();
            InsControl._scope.CH1_Level(Vmax / 4);

            InsControl._scope.DoCommand(":FUNCtion1:FFTMagnitude CHANnel" + channel);
            // _scope.DoCommand(":FUNCtion1:FFT:DETector:TYPE NORMal");
            // _scope.DoCommand(":FUNCtion1:FFT:DETector:POINts 5");
            InsControl._scope.DoCommand(":FUNCtion:FFT:PEAK:SORT IFRequency");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:VUNits DBUV");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:HSCale LOG");
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY ON");
            // need to input parameter by user
            // Start 150K, Stop 30M. RBW 9K
            InsControl._scope.DoCommand(":FUNCtion1:FFT:STOP 30E6");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:START 150E3");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:RESolution 9000");
            // _scope.DoCommand(":FUNCtion1:FFT:PEAK:STATe ON"); 

            //:MEASure
            InsControl._scope.DoCommand(":FUNC1:SCALe 50");
            InsControl._scope.DoCommand(":FUNC1:OFFSet 200");
            double max = InsControl._scope.doQueryNumber(":MEASure:VMAX? FUNC1");

            InsControl._scope.DoCommand(":FUNC1:SCALe " + (max / 10));
            InsControl._scope.DoCommand(":FUNC1:OFFSet " + (max + 30));


            double peak1 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 1, " + test_parameter.peak_level1);
            double peak2 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 2, " + test_parameter.peak_level1);
            double peak3 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 3, " + test_parameter.peak_level1);
            double peak4 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 4, " + test_parameter.peak_level1);
            double peak5 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 5, " + test_parameter.peak_level1);

            double magn1 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 1, " + test_parameter.peak_level1);
            double magn2 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 2, " + test_parameter.peak_level1);
            double magn3 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 3, " + test_parameter.peak_level1);
            double magn4 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 4, " + test_parameter.peak_level1);
            double magn5 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 5, " + test_parameter.peak_level1);

            _sheet.Cells[row, XLS_Table.G] = peak1;
            _sheet.Cells[row, XLS_Table.H] = peak2;
            _sheet.Cells[row, XLS_Table.I] = peak3;
            _sheet.Cells[row, XLS_Table.J] = peak4;
            _sheet.Cells[row, XLS_Table.K] = peak5;
            _sheet.Cells[row, XLS_Table.L] = magn1;
            _sheet.Cells[row, XLS_Table.M] = magn2;
            _sheet.Cells[row, XLS_Table.N] = magn3;
            _sheet.Cells[row, XLS_Table.O] = magn4;
            _sheet.Cells[row, XLS_Table.P] = magn5;

            MyLib.Delay1ms(500);
            // Start 30M, Stop 200M. RBW 120K
            InsControl._scope.DoCommand(":FUNCtion1:FFT:STOP 200E6");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:START 30E6");
            InsControl._scope.DoCommand(":FUNCtion1:FFT:RESolution 120E3");

            peak1 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 1," + test_parameter.peak_level2);
            peak2 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 2," + test_parameter.peak_level2);
            peak3 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 3," + test_parameter.peak_level2);
            peak4 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 4," + test_parameter.peak_level2);
            peak5 = InsControl._scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 5," + test_parameter.peak_level2);

            magn1 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 1," + test_parameter.peak_level2);
            magn2 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 2," + test_parameter.peak_level2);
            magn3 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 3," + test_parameter.peak_level2);
            magn4 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 4," + test_parameter.peak_level2);
            magn5 = InsControl._scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 5," + test_parameter.peak_level2);

            _sheet.Cells[row, XLS_Table.Q] = peak1;
            _sheet.Cells[row, XLS_Table.R] = peak2;
            _sheet.Cells[row, XLS_Table.S] = peak3;
            _sheet.Cells[row, XLS_Table.T] = peak4;
            _sheet.Cells[row, XLS_Table.U] = peak5;
            _sheet.Cells[row, XLS_Table.V] = magn1;
            _sheet.Cells[row, XLS_Table.W] = magn2;
            _sheet.Cells[row, XLS_Table.X] = magn3;
            _sheet.Cells[row, XLS_Table.Y] = magn4;
            _sheet.Cells[row, XLS_Table.Z] = magn5;
            
        }

        private void OSCInint()
        {
            InsControl._scope.AgilentOSC_RST();
            MyLib.WaveformCheck();

            InsControl._scope.CH1_On(); // LX
            InsControl._scope.CH1_Level(20);
            InsControl._scope.CH1_Offset(40);
            InsControl._scope.CH1_BWLimitOn();
        }

        public void ATE_Task()
        {
            
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            int bin_cnt = 1;
            bin_cnt = binList.Length;
            double[] ori_vinTable = new double[test_parameter.vinList.Count];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, test_parameter.vinList.Count);
            RTDev.BoadInit();
            int row = 6;
            int idx = 0;
#if true
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;

            _sheet.Cells[1, XLS_Table.A] = "Vin";
            _sheet.Cells[2, XLS_Table.A] = "Iout";
            _sheet.Cells[3, XLS_Table.A] = "Date";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _sheet.Cells[5, XLS_Table.A] = "Version";

            _sheet.Cells[1, XLS_Table.B] = test_parameter.vin_info;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.eload_info;
            _sheet.Cells[3, XLS_Table.B] = test_parameter.date_info;
            _sheet.Cells[5, XLS_Table.B] = test_parameter.ver_info;


            _sheet.Cells[row, XLS_Table.A] = "Temp (C)";
            _sheet.Cells[row, XLS_Table.B] = "VIN (V)";
            _sheet.Cells[row, XLS_Table.C] = "Iin (mA)";
            _sheet.Cells[row, XLS_Table.D] = "Iout (mA)";
            _sheet.Cells[row, XLS_Table.E] = "Bin";
            _sheet.Cells[row, XLS_Table.F] = "Code / Duty";

            _sheet.Cells[row, XLS_Table.G] = "Meas1 peak1 Freq";
            _sheet.Cells[row, XLS_Table.H] = "Meas1 peak2 Freq";
            _sheet.Cells[row, XLS_Table.I] = "Meas1 peak3 Freq";
            _sheet.Cells[row, XLS_Table.J] = "Meas1 peak4 Freq";
            _sheet.Cells[row, XLS_Table.K] = "Meas1 peak5 Freq";
            _sheet.Cells[row, XLS_Table.L] = "Meas1 peak1 Magn";
            _sheet.Cells[row, XLS_Table.M] = "Meas1 peak2 Magn";
            _sheet.Cells[row, XLS_Table.N] = "Meas1 peak3 Magn";
            _sheet.Cells[row, XLS_Table.O] = "Meas1 peak4 Magn";
            _sheet.Cells[row, XLS_Table.P] = "Meas1 peak5 Magn";

            _sheet.Cells[row, XLS_Table.Q] = "Meas2 peak1 Freq";
            _sheet.Cells[row, XLS_Table.R] = "Meas2 peak2 Freq";
            _sheet.Cells[row, XLS_Table.S] = "Meas2 peak3 Freq";
            _sheet.Cells[row, XLS_Table.T] = "Meas2 peak4 Freq";
            _sheet.Cells[row, XLS_Table.U] = "Meas2 peak5 Freq";
            _sheet.Cells[row, XLS_Table.V] = "Meas2 peak1 Magn";
            _sheet.Cells[row, XLS_Table.W] = "Meas2 peak2 Magn";
            _sheet.Cells[row, XLS_Table.X] = "Meas2 peak3 Magn";
            _sheet.Cells[row, XLS_Table.Y] = "Meas2 peak4 Magn";
            _sheet.Cells[row, XLS_Table.Z] = "Meas2 peak5 Magn";
            _range = _sheet.Range["A" + row, "P" + row];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
#endif
            
            OSCInint();
            MessageBox.Show("Please setting FFT Function !!");
            for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
            {
                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    string file_name = string.Format("{0}_Temp={1}_Vin={2}_{3}_{4}",
                                        idx,
                                        temp,
                                        vin_idx,
                                        test_parameter.brightness_sel ? 
                                        "I2C_Code=" + test_parameter.i2c_code :
                                        "Duty=" + test_parameter.duty,
                                        Path.GetFileNameWithoutExtension(binList[bin_idx])
                                        );
                    //for(int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    //{
                    if (test_parameter.run_stop == true) goto Stop;

                    InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);
                    //MyLib.Switch_ELoadLevel(test_parameter.ioutList[iout_idx]);
                    //InsControl._eload.CH1_Loading(test_parameter.ioutList[iout_idx]);

                    if (test_parameter.brightness_sel)
                    {
                        byte MSB = (byte)((test_parameter.i2c_code & 0xFF00) >> 8);
                        byte LSB = (byte)(test_parameter.i2c_code & 0xFF);
                        // brightness code
                    }
                    else
                    {
                        func_gen_fixed_parameter(test_parameter.freq, test_parameter.duty);
                    }


                    double tempVin = ori_vinTable[vin_idx];
                    if (!MyLib.Vincompensation(ori_vinTable[vin_idx], ref tempVin))
                    {
                        System.Windows.Forms.MessageBox.Show("Please connect DAQ !!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
                        return;
                    }
#if true
                    double vin, iin, iout;
                    vin = InsControl._power.GetVoltage();
                    iin = InsControl._power.GetCurrent() * 1000;
                    iout = InsControl._eload.GetIout() * 1000;
                    _sheet.Cells[row, XLS_Table.A] = vin;
                    _sheet.Cells[row, XLS_Table.B] = iin;
                    _sheet.Cells[row, XLS_Table.C] = iout;
                    _sheet.Cells[row, XLS_Table.D] = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                    _sheet.Cells[row, XLS_Table.E] = test_parameter.brightness_sel ? "Code" : "Duty";
#endif
                    FFT_Task(row);
                    InsControl._scope.SaveWaveform(test_parameter.wave_path, file_name);
                    row++; idx++;
                    //} // iout loop
                } // bin loop
            } // vin loop

        Stop:
            InsControl._scope.DoCommand(":FUNCTION1:DISPLAY OFF");
            System.Windows.Forms.MessageBox.Show("Test finished!!!", "FFT Item", System.Windows.Forms.MessageBoxButtons.OK);

        }
    }
}
