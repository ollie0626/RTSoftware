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
using MetroFramework.Controls;
using System.Globalization;

namespace DP_ATE_Tool
{
    class ATE_Lx : ITask
    {
        public int Mode;
        
        public string PSU_Ori;
        public List<double> PSU_List = new List<double>();

        public int topology; // 0: Buck,Inverting 1:Boost 2:Buck-Boost

        public string Bin_Path;
        public byte Bin_SlaveID;

        public string Freq_Ori;
        public List<double> Freq_List = new List<double>();

        public string Duty_Ori;
        public List<double> Duty_List = new List<double>();

        public string Code_Ori;
        public byte Code_DataAddr;
        List<two_byte> Code_List = new List<two_byte>();

        public bool inverting;
        public bool buck = false;
        public bool boost;
        public bool buck_boost;
        public bool[] ItemEn = new bool[3]; // 0=Freq, 1=Slew Rate, 2=Jitter

        // Excel variable
        Excel.Application app;
        Excel.Workbook book;
        Excel.Worksheet sheet, sheet_2;
        Excel.Range range;

        // DEVICE
        public PowerModule _PSU;
        public AgilentOSC _scope;
        public MultiChannelModule _34970A;
        public FuncGenModule _fun_gen;
        public I2CModule _i2cModule;
        BridgeBoardEnum hEnum;
        BridgeBoard hDevice;

        string[] BinFileList = null;

        int X_Yindex = 0;
        //string[] XRange = new string[] { "W", "AE", "AM", "AU" };
        //string[] YRange = new string[] { "AB", "AK", "AS", "BA" };

        string[] XRange = new string[] { "Y", "AH", "AQ", "AZ" };
        string[] YRange = new string[] { "AE", "AN", "AW", "BF" };

        MyLib MyLib = new MyLib();
        DataHandle datahandle = new DataHandle();
        CultureInfo culture = new CultureInfo("en-US");

        //INSTRUMENT
        private void RTBBConnect()
        {
            hEnum = BridgeBoardEnum.GetBoardEnum();
            hDevice = BridgeBoard.ConnectByDefault(hEnum);
            if (hDevice != null)
            {
                _i2cModule = hDevice.GetI2CModule();
                _i2cModule.RTBB_I2CSetFrequency(GlobalVariable.ERTI2CFrequency.eRTI2CFreq50KHz, 50);
            }
        }
        private void device_init()
        {
            _scope.AgilentOSC_RST();
            _PSU.AutoSetOCP(7);
            _PSU.AutoSelPowerOn(PSU_List[0]);
        }



        private void CreateExcel()
        {
            app = new Excel.Application();
            app.Visible = true;
            book = app.Workbooks.Add();
        }

        private void func_gen_fixed_parameter(double freq, double duty)
        {
            //_fun_gen.CH1_Off();
            _fun_gen.CH1_ContinuousMode();
            _fun_gen.CH1_PulseMode();
            MyLib.DelayMs(500);

            _fun_gen.CH1_Frequency(freq);
            MyLib.DelayMs(500);

            _fun_gen.CH1_DutyCycle(duty);
            _fun_gen.CHl1_HiLevel(1.6);
        }

        private void Time_ReScale(ref double time_scale, ref double freq, int channel)
        {
            while (freq > (100000 * 1000))
            {
                time_scale += 1;
                _scope.TimeScaleUs(time_scale);
                freq = _scope.Measure_Freq(channel);
                if (time_scale >= 100) break;
            }
        }

        private void Channel_LevelInit(int channel)
        {
            if (inverting)
                _scope.Ch_Offset(channel, 0.5);
            else if (buck)
                _scope.CHx_Level(channel, 4);
            else if (boost)
            {
                _scope.CHx_Level(channel, 10);
                _scope.Ch_Offset(channel, 5);
            }

        }

        private void Channel_LevelSetting(int channel)
        {
            double Vmax, Vmin, avg = 0;
            if (buck || boost)
            {
                for (int i = 0; i <= 1; i++)
                {
                    Vmax = _scope.Measure_Ch_Max(channel);
                    Vmin = _scope.Measure_Ch_min(channel);
                    avg = Vmax - Vmin;
                    _scope.CHx_Level(channel, avg / 5);
                    System.Threading.Thread.Sleep(10);
                    _scope.Ch_Offset(channel, avg / 2);
                }
            }
            else if (inverting)
            {
                for (int i = 0; i <= 1; i++)
                {
                    Vmax = Math.Abs(_scope.Measure_Ch_Max(channel));
                    Vmin = Math.Abs(_scope.Measure_Ch_min(channel));
                    if (Vmax > Vmin)
                        avg = Vmax;
                    else if (Vmax < Vmin)
                        avg = Vmin;
                    double value = avg / 3;
                    _scope.CHx_Level(channel, value);
                }
            }
        }

        private void Channel_TriggerLevel(int channel)
        {
            double Vtop, Vbase;
            double trigger_level;

            if (buck || inverting)
            {
                _scope.SetTrigModeEdge(false);
                Vtop = _scope.Measure_Top(channel);
                Vbase = _scope.Measure_Base(channel);
                trigger_level = 0.65 * Vtop + 0.35 * Vbase;
                _scope.Trigger_Level(channel, trigger_level);
            }
            else if (boost)
            {
                _scope.SetTrigModeEdge(true);
                Vtop = _scope.Measure_Top(channel);
                Vbase = _scope.Measure_Base(channel);
                trigger_level = 0.45 * Vtop + 0.65 * Vbase;
                _scope.Trigger_Level(channel, trigger_level);
            }
        }


        private void FFT_Task()
        {
            MessageBox.Show("Please setting FFT Function !!");

            int channel = 1;
            double meas_dm_level = 100;

            _scope.Measure_Clear();
            _scope.Measure_Freq(channel);
            _scope.DoCommand(":MARKer:MODE OFF");
            System.Threading.Thread.Sleep(1000);
            _scope.Bandwidth_Limit_On(channel);
            _scope.Ch_On(1);
            _scope.TimeScaleUs(20);
            MyLib.DelayMs(100);
            /* time scale setting */
            double time_scale = 0.02;
            double freq = _scope.Measure_Freq(channel);
            Time_ReScale(ref time_scale, ref freq, channel);
            time_scale = ((1 / freq) * 1000000 * 3) / 10;
            _scope.CHx_Level(channel, 20);
            _scope.Ch_Offset(channel, 40);
            _scope.Trigger(channel);
            _scope.TimeScaleUs(time_scale);
            _scope.TimeBasePosition(0);
            MyLib.DelayMs(500);
            Channel_LevelInit(channel);
            Channel_LevelSetting(channel);
            Channel_TriggerLevel(channel);
            Channel_LevelSetting(channel);
            // _scope.Measurement_Threshold_Percent_Mode(channel);

            _scope.DoCommand(":FUNCtion1:FFTMagnitude CHANnel1");
            // _scope.DoCommand(":FUNCtion1:FFT:DETector:TYPE NORMal");
            // _scope.DoCommand(":FUNCtion1:FFT:DETector:POINts 5");
            _scope.DoCommand(":FUNCtion:FFT:PEAK:SORT IFRequency");
            _scope.DoCommand(":FUNCtion1:FFT:VUNits DBUV");
            _scope.DoCommand(":FUNCtion1:FFT:HSCale LOG");
            _scope.DoCommand(":FUNCTION1:DISPLAY ON");
            // need to input parameter by user
            // Start 150K, Stop 30M. RBW 9K
            _scope.DoCommand(":FUNCtion1:FFT:STOP 30E6");
            _scope.DoCommand(":FUNCtion1:FFT:START 150E3");
            _scope.DoCommand(":FUNCtion1:FFT:RESolution 9000");
            // _scope.DoCommand(":FUNCtion1:FFT:PEAK:STATe ON"); 

            //:MEASure
            _scope.DoCommand(":FUNC1:SCALe 50");
            _scope.DoCommand(":FUNC1:OFFSet 200");
            double max = _scope.doQueryNumber(":MEASure:VMAX? FUNC1");

            _scope.DoCommand(":FUNC1:SCALe " + (max / 10));
            _scope.DoCommand(":FUNC1:OFFSet " + (max + 30));


            double peak1 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 1, " + meas_dm_level);
            double peak2 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 2, " + meas_dm_level);
            double peak3 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 3, " + meas_dm_level);
            double peak4 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 4, " + meas_dm_level);
            double peak5 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 5, " + meas_dm_level);

            double magn1 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 1, " + meas_dm_level);
            double magn2 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 2, " + meas_dm_level);
            double magn3 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 3, " + meas_dm_level);
            double magn4 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 4, " + meas_dm_level);
            double magn5 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 5, " + meas_dm_level);

            MyLib.DelayMs(500);
            // Start 30M, Stop 200M. RBW 120K
            _scope.DoCommand(":FUNCtion1:FFT:STOP 200E6");
            _scope.DoCommand(":FUNCtion1:FFT:START 30E6");
            _scope.DoCommand(":FUNCtion1:FFT:RESolution 120E3");

            peak1 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 1," + meas_dm_level);
            peak2 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 2," + meas_dm_level);
            peak3 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 3," + meas_dm_level);
            peak4 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 4," + meas_dm_level);
            peak5 = _scope.doQueryNumber(":MEASure:FFT:FREQuency? FUNC1, 5," + meas_dm_level);
            
            magn1 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 1," + meas_dm_level);
            magn2 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 2," + meas_dm_level);
            magn3 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 3," + meas_dm_level);
            magn4 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 4," + meas_dm_level);
            magn5 = _scope.doQueryNumber(":MEASure:FFT:MAGNitude? FUNC1, 5," + meas_dm_level);

            _scope.DoCommand(":FUNCTION1:DISPLAY OFF");
        }



        //BIN & FUNCTION GEN
        private void DealwithBins()
        {
            if (Bin_Path != null && Bin_Path != "")
            {
                DirectoryInfo di = new DirectoryInfo(Bin_Path);
                BinFileList = Directory.GetFiles(Bin_Path, "*.bin");
            }
        }
        private void dealWithByte(string str, ref List<two_byte> list)
        {
            list.Clear();
            string[] tmp = str.Split(',');

            two_byte tb = new two_byte();
            for (int i = 0; i < tmp.Length; i++)
            {
                byte[] res = datahandle.to2Byte(tmp[i]);
                tb.LSB = res[0];
                tb.MSB = res[1];

                list.Add(tb);
            }
        }

        private void main_loop(Excel.Worksheet sheet, int channel)
        {
            int row = 28;
            int waveRow = 28;

            for (int layer_psu = 0; layer_psu < PSU_List.Count; layer_psu++)
            {
                _PSU.AutoSelPowerOn(PSU_List[layer_psu]);

                int BinCount = BinFileList != null ? BinFileList.Length : 1;
                for (int layer_bin = 0; layer_bin < BinCount; layer_bin++)
                {
                    int CodeNum = Code_List.Count == 0 ? 1 : Code_List.Count;
                    for (int layer_code = 0; layer_code < CodeNum; layer_code++)
                    {
                        int FreqNum = Freq_List.Count == 0 ? 1 : Freq_List.Count;
                        for (int layer_freq = 0; layer_freq < FreqNum; layer_freq++)
                        {
                            int DutyNum = Duty_List.Count == 0 ? 1 : Duty_List.Count;
                            for(int layer_duty = 0; layer_duty < DutyNum; layer_duty++)
                            {
                                if(Mode == 0)
                                    func_gen_fixed_parameter(Freq_List[layer_freq] * 1000, Duty_List[layer_duty]);
                                else
                                {
                                    byte[] c1 = new byte[] { Code_List[layer_code].LSB, Code_List[layer_code].MSB };
                                    if (_i2cModule != null)
                                        _i2cModule.RTBB_I2CWrite(Bin_SlaveID >> 1, 0x01, Code_DataAddr, 0x02, c1);
                                }    
                                printTitle(sheet, ref row);

                                if (BinFileList != null)
                                    MyLib.runBin(ref _i2cModule, BinFileList[layer_bin], Bin_SlaveID);

                                X_Yindex = 0;
                                _PSU.AutoSelPowerOn(PSU_List[layer_psu]);

                                if (layer_bin == 0)
                                    MyLib.DelayMs(5000);

                                double vin = _34970A.Get_100Vol(1); //_PSU.GetVoltage();
                                double iin = _PSU.GetCurrent() * 1000;
                                double vout = _34970A.Get_100Vol(2);

                                sheet.Cells[row, XLS_Table.B] = main.ChamberEn ? main.temperature.ToString() : "";
                                sheet.Cells[row, XLS_Table.C] = "LINK";
                                sheet.Cells[row, XLS_Table.D] = vin;
                                sheet.Cells[row, XLS_Table.E] = iin;

                                if (BinFileList != null)
                                    sheet.Cells[row, XLS_Table.F] = Path.GetFileName(BinFileList[layer_bin]);



                                if (Mode == 0)
                                {
                                    sheet.Cells[row, XLS_Table.G] = string.Format("{0:#.##}", Freq_List[layer_freq]);
                                    sheet.Cells[row, XLS_Table.H] = string.Format("{0:#.##}", Duty_List[layer_duty]);
                                }
                                else
                                    sheet.Cells[row, XLS_Table.H] = string.Format("{0:X02}{1:X02}", Code_List[layer_code].MSB, Code_List[layer_code].LSB);

                                sheet.Cells[row, XLS_Table.I] = vout;
                                range = sheet.Range["A" + row, "W" + row];
                                range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                bool check_freq_valid = true;

                                X_Yindex = 0;
                                if (ItemEn[0])
                                    check_freq_valid = FreqTask(row, waveRow, sheet, PSU_List[layer_psu], channel);

                                X_Yindex = 1;
                                if (check_freq_valid && ItemEn[1])
                                    SlewRateTask(row, waveRow, sheet, channel);

                                if (check_freq_valid && ItemEn[2])
                                    JitterTask(row, waveRow, sheet, channel);

                                _scope.SetTrigModeEdge(false);
                                waveRow += 20;
                                row++;
                            }
                            
                        }
                    }
                }

                _PSU.AutoPowerOff();
                MyLib.DelayMs(2000);
            }
        }

        public void ATE_Task()
        {
            RTBBConnect();
            device_init();

            CreateExcel();
            FillSheet(sheet);
            ReportInit(topology);

            DealwithBins();
            if (Mode == 1)
                datahandle.str_to_TwoByteList(Code_Ori, ref Code_List);

            InputTitleAndUserInput();

            if (topology == 0 || topology == 1)
                main_loop(sheet, 1);
            else
            {
                inverting = true;
                boost = false;
                main_loop(sheet, 1);

                inverting = false;
                boost = true;
                main_loop(sheet_2, 3);
            }

            var culture = new CultureInfo("en-US");
            sheet.Cells[3, XLS_Table.J] = string.Format("{0}", DateTime.Now.ToString(culture));

            for (int i = 1; i <= 38; i++)
                sheet.Columns[i].AutoFit();

            if (topology == 3)
            {
                sheet_2.Cells[3, XLS_Table.J] = string.Format("{0}", DateTime.Now.ToString(culture));
                for (int i = 1; i <= 38; i++)
                    sheet_2.Columns[i].AutoFit();
            }

            MyLib.SaveExcelReport(main.Report_Path, "Lx_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), book);
            book.Close(false);
            app.Quit();
        }
    }
}
