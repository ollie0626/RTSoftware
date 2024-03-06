﻿
#define Report_en

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Sunny.UI;

namespace IN528ATE_tool
{
    public class ATE_PowerOn : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        //Excel.Chart _chart;

        public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        public ATE_PowerOn()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void OSCInit()
        {
            InsControl._scope.DoCommand("SYSTem:CONTrol \"ExpandAbout - 1 xpandGnd\"");
            InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
            InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);
            MyLib.WaveformCheck();

            InsControl._scope.CH1_On();
            InsControl._scope.CH2_On();
            InsControl._scope.CH3_On();
            InsControl._scope.CH4_On();

            if(test_parameter.bw_en)
            {
                InsControl._scope.CH1_BWLimitOn();
                InsControl._scope.CH2_BWLimitOn();
                InsControl._scope.CH3_BWLimitOn();
                InsControl._scope.CH4_BWLimitOn();
            }

            InsControl._scope.CH2_Level(test_parameter.ch2_level);
            InsControl._scope.Trigger_CH1();
            InsControl._scope.TriggerLevel_CH1(test_parameter.trigger_level); // gui trigger level
            InsControl._scope.AutoTrigger();
            RTDev.GpEn_Disable();
            InsControl._scope.Root_RUN();
            MyLib.WaveformCheck();
        }


        private void Scope_Channel_Resize(int idx, string path)
        {
            InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
            MyLib.Delay1ms(250);
            // write default bin file
            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
            MyLib.Delay1ms(100);
            RTDev.I2C_WriteBin((byte)(test_parameter.slave >> 1), 0x00, path); // test conditions
            MyLib.Delay1ms(250);

            // program test conditons 
            if (test_parameter.mtp_enable)
            {
                byte[] buf = new byte[] { test_parameter.mtp_data };
                RTDev.I2C_Write((byte)(test_parameter.mtp_slave >> 1), test_parameter.mtp_addr, buf);
            }
            MyLib.Delay1ms(250);
            InsControl._eload.CH1_Loading(0.01);
            InsControl._scope.AutoTrigger();
            InsControl._scope.Trigger_CH1();

            // inital channel level setting
            if (test_parameter.trigger_vin_en)
            {
                InsControl._scope.TriggerLevel_CH1(test_parameter.trigger_level);
                InsControl._scope.CH1_Level(test_parameter.VinList[idx] / 3);
            }
            else if(test_parameter.trigger_en)
            {
                InsControl._scope.CH1_Level(1);
                RTDev.GpEn_Enable();
            }

            InsControl._scope.CH2_Level(test_parameter.ch2_level);

            if (test_parameter.dt_rising_en) InsControl._scope.CH2_Offset(test_parameter.ch2_level);
            else InsControl._scope.CH2_Offset(test_parameter.ch2_level * -1);


            if(!test_parameter.ch2_user_define)
            {
                for (int i = 0; i < 3; i++)
                {
                    double Vo;
                    Vo = Math.Abs(InsControl._scope.Meas_CH2MAX());

                    InsControl._scope.CH3_Level(Vo / 3);
                    InsControl._scope.CH2_Level(Vo / 3);
                    MyLib.WaveformCheck();
                }
            }


            RTDev.GpEn_Disable();
            InsControl._power.AutoPowerOff();
            InsControl._eload.AllChannel_LoadOff();
            MyLib.Delay1ms(300);
        }



        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            RTDev.BoadInit();
            RTDev.GpioInit();

            int vin_cnt = test_parameter.VinList.Count;
            int iout_cnt = test_parameter.IoutList.Count;
            int row = 22;
            int wave_row = 22;
            string[] binList;
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;
            binList = Mylib.ListBinFile(test_parameter.binFolder);
            bin_cnt = binList.Length;
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if Report_en
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            Mylib.ExcelReportInit(_sheet);
            Mylib.testCondition(_sheet, "Delay Time & Soft-start Time", bin_cnt, temp);

            _sheet.Cells[row, XLS_Table.A] = "No.";
            _sheet.Cells[row, XLS_Table.B] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.C] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.D] = "Iout(mA)";
            _sheet.Cells[row, XLS_Table.E] = "Bin file";
            _sheet.Cells[row, XLS_Table.F] = "delay time(ms)";
            _sheet.Cells[row, XLS_Table.G] = "Soft Start(ms)";
            _sheet.Cells[row, XLS_Table.H] = "Vmax(V)";
            _sheet.Cells[row, XLS_Table.I] = "Power on Inrush(A)";
            _sheet.Cells[row, XLS_Table.J] = "Power off Inrush(A)";
            _range = _sheet.Range["A" + row, "E" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range = _sheet.Range["F" + row, "J" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            row++;
#endif
            InsControl._power.AutoPowerOff();
            OSCInit();
            //MyLib.Delay1s(1);
            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {

                InsControl._scope.TimeScaleMs(test_parameter.ontime_scale_ms);
                InsControl._scope.TimeBasePositionMs(test_parameter.ontime_scale_ms * 3);

                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;

                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();
                        //Scope_Channel_Resize(vin_idx);

                        /* test initial setting */
                        InsControl._scope.DoCommand(":MARKer:MODE OFF");
                        string file_name;
                        string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        file_name = string.Format("{0}{1}_Temp={2}C_vin={3:0.##}V_iout={4:0.##}A",
                                                    "",res, temp,
                                                    test_parameter.VinList[vin_idx],
                                                    test_parameter.IoutList[iout_idx]
                                                    );
                        // inside has auto trigger
                        Scope_Channel_Resize(vin_idx, binList[bin_idx]);

                        //:MARKer:MEASurement:MEASurement {MEASurement<N>}
                        Mylib.eLoadLevelSwich(InsControl._eload, test_parameter.IoutList[iout_idx]);
                        InsControl._eload.CH1_Loading(test_parameter.IoutList[iout_idx]);
                        double tempVin = ori_vinTable[vin_idx];

                        MyLib.WaveformCheck();
                        
                        Change_scale:
                        InsControl._scope.NormalTrigger();
                        if (test_parameter.trigger_vin_en)
                        {
                            // vin trigger
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            // rising edge trigger
                            InsControl._scope.SetTrigModeEdge(false);
                            MyLib.Delay1s(1);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1s(1);
                            
                        }
                        else if (test_parameter.trigger_en)
                        {
                            //Gpio 2.0 trigger
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1s(1);
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            InsControl._scope.SetTrigModeEdge(false);
                            InsControl._scope.TriggerLevel_CH1(1.5);
                            MyLib.Delay1ms(500);
                            RTDev.GpEn_Enable();
                            MyLib.Delay1s(1);
                        }
                        else
                        {
                            // I2c run and GPIO trigger
                            InsControl._scope.DoCommand(":TRIGger:MODE EDGE");
                            InsControl._scope.Trigger(1);
                            InsControl._scope.TriggerLevel_CH1(1.65);
                            InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                            MyLib.Delay1ms(250);
                            InsControl._scope.Root_STOP();
                            MyLib.Delay1ms(100);
                            if (test_parameter.specify_bin != "") RTDev.I2C_WriteBin((byte)(test_parameter.specify_id >> 1), 0x00, test_parameter.specify_bin);
                            InsControl._scope.NormalTrigger();
                            InsControl._scope.Root_RUN();
                            if (binList[0] != "") RTDev.I2C_WriteBinAndGPIO((byte)(test_parameter.slave), 0x00, binList[bin_idx]);
                            MyLib.Delay1ms(250);
                            InsControl._scope.Measure_Clear();
                        }
                        InsControl._scope.Root_STOP();
                        double delay_time, ss_time, Vmax, Inrush;

                        // Delay time Measure
                        InsControl._scope.CH4_On();
                        InsControl._scope.DoCommand(":MEASure:THResholds:METHod CHANnel1,ABSolute");
                        InsControl._scope.DoCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute CHANnel1,{0},{1},{2}",
                                                        test_parameter.hivol,
                                                        test_parameter.midvol,
                                                        test_parameter.lovol));
                        MyLib.Delay1ms(100);
                        double vbase, vtop, vmid;
                        InsControl._scope.DoCommand(":MEASure:THResholds:METHod CHANnel2,ABSolute");
                        if (test_parameter.dt_rising_en)
                        {
                            vtop = InsControl._scope.Meas_CH2Top();
                            vbase = InsControl._scope.Meas_CH2Base();
                            vmid = vtop * 0.5;
                            InsControl._scope.DoCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute CHANnel2,{0},{1},{2}",
                                                        vmid,
                                                        vmid * 0.5,
                                                        vbase));
                        }
                        else
                        {
                            vtop = InsControl._scope.Meas_CH2Top();
                            vbase = InsControl._scope.Meas_CH2Base();
                            vmid = vbase * 0.5;
                            InsControl._scope.DoCommand(string.Format(":MEASure:THResholds:GENeral:ABSolute CHANnel2,{0},{1},{2}",
                                                        vtop,
                                                        vmid * 0.5,
                                                        vbase));
                        }


                        // Delay time
                        if (test_parameter.dt_rising_en)
                            // set rising to rising
                            InsControl._scope.SetDeltaTime(true, 1, 2, true, 1, 0); // rising to rising (low)
                        else
                            // set rising to falling
                            InsControl._scope.SetDeltaTime(true, 1, 2, false, 1, 2); // rising to falling (up)
                        
                        InsControl._scope.DoCommand(":MEASure:DELTatime CHANnel1, CHANnel2");
                        InsControl._scope.DoCommand(":MARKer:MODE MEASurement");
                        delay_time = InsControl._scope.doQueryNumber(":MEASure:DELTatime? CHANnel1, CHANnel2");
                        Vmax = InsControl._scope.Meas_CH1MAX();
                        Inrush = InsControl._scope.Meas_CH4MAX();
                        // Delay time waveform
                        double scope_time_scale = InsControl._scope.doQueryNumber(":TIMEBASE:SCALE?");
                        if(delay_time > (scope_time_scale * 4))
                        {
                            if (delay_time > Math.Pow(10, 7))
                            {
                                InsControl._scope.TimeScale(scope_time_scale * 2);
                                InsControl._scope.TimeBasePosition(scope_time_scale * 2 * 3);
                                InsControl._scope.Root_RUN();
                                Power_Off();
                                goto Change_scale;
                            }
                            InsControl._scope.TimeScale(delay_time / 4);
                            InsControl._scope.TimeBasePosition((delay_time / 4) * 3);
                            InsControl._scope.Root_RUN();
                            Power_Off();
                            goto Change_scale;
                        }
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_DT");


                        // Soft-Start times
                        switch(test_parameter.sst_sel)
                        {
                            case 0:
                                InsControl._scope.DoCommand(":MEASure:THResholds:METHod CHANnel2,PERCent");
                                InsControl._scope.DoCommand(":MEASure:THResholds:RFALl:PERCent CHANnel2,100,50,0");
                                break;
                            case 1:
                                InsControl._scope.DoCommand(":MEASure:THResholds:METHod CHANnel2,ABSolute");
                                string cmd = string.Format(":MEASure:THResholds:RFALl:ABSolute CHANnel2,{0},{1},{2}", test_parameter.hivout, test_parameter.midvout, test_parameter.lovout);
                                InsControl._scope.DoCommand(cmd);
                                break;
                        }

                        if(test_parameter.dt_rising_en)
                            InsControl._scope.DoCommand(":MEASure:RISetime CHANnel2");
                        else
                            InsControl._scope.DoCommand(":MEASure:FALLtime CHANnel2");
                        
                        InsControl._scope.DoCommand(":MARKer:MEASurement:MEASurement MEASurement1");
                        if (test_parameter.dt_rising_en)
                            ss_time = InsControl._scope.Meas_CH2Rise();
                        else
                            ss_time = InsControl._scope.Meas_CH2Fall();

                        // Soft-Start time waveform
                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_SST");

                        InsControl._scope.Measure_Clear();
                        MyLib.Delay1s(1);
                        InsControl._scope.Root_Clear();
                        InsControl._scope.Root_RUN();
#if Report_en
                        // gpio control for relay 
                        _sheet.Cells[row, XLS_Table.A] = row - 22;
                        _sheet.Cells[row, XLS_Table.B] = temp;
                        _sheet.Cells[row, XLS_Table.C] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.IoutList[iout_idx];
                        _sheet.Cells[row, XLS_Table.E] = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        _sheet.Cells[row, XLS_Table.F] = delay_time * 1000;
                        _sheet.Cells[row, XLS_Table.G] = ss_time * 1000;
                        _sheet.Cells[row, XLS_Table.H] = Vmax;
                        _sheet.Cells[row, XLS_Table.I] = Inrush;
#endif
                        scope_time_scale = InsControl._scope.doQueryNumber(":TIMEBASE:SCALE?");
                        InsControl._scope.TimeScaleMs(test_parameter.offtime_scale_ms);
                        InsControl._scope.TimeBasePositionMs(test_parameter.offtime_scale_ms * 1);
                        System.Threading.Thread.Sleep(1000);
                        InsControl._scope.NormalTrigger();
                        InsControl._scope.Trigger_CH2();
                        

                        if(test_parameter.dt_rising_en)
                        {
                            InsControl._scope.TriggerLevel_CH2(InsControl._scope.doQueryNumber(":CHANnel2:SCALe?"));
                            InsControl._scope.SetTrigModeEdge(true);
                        }
                        else
                        {
                            InsControl._scope.TriggerLevel_CH2(InsControl._scope.doQueryNumber(":CHANnel2:SCALe?") * -1);
                            InsControl._scope.SetTrigModeEdge(false);
                        }

                        InsControl._scope.Root_RUN();
                        System.Threading.Thread.Sleep(1000);


                        // power off section
                        Power_Off();
                        InsControl._scope.DoCommand(":MEASure:VMAX CHANnel4");
                        InsControl._scope.DoCommand(":MEASure:VMIN CHANnel4");
                        RTDev.GpEn_Disable();
                        System.Threading.Thread.Sleep(800);
                        Inrush = InsControl._scope.Meas_CH4MAX();
                        InsControl._scope.CH4_On();
#if Report_en
                        _sheet.Cells[row, XLS_Table.J] = Inrush;

                        InsControl._scope.SaveWaveform(test_parameter.waveform_path, file_name + "_OFF");

                        // past waveform
                        _sheet.Cells[wave_row, XLS_Table.P] = "超連結";
                        _sheet.Cells[wave_row, XLS_Table.Q] = "Temp(C)";
                        _sheet.Cells[wave_row, XLS_Table.R] = "Vin(V)";
                        _sheet.Cells[wave_row, XLS_Table.S] = "Iout(A)";
                        _sheet.Cells[wave_row, XLS_Table.T] = "Bin file";
                        _range = _sheet.Range["P" + wave_row, "T" + wave_row];
                        _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        _range.Interior.Color = Color.FromArgb(124, 252, 0);

                        _sheet.Cells[wave_row + 1, XLS_Table.P] = "LINK";
                        _sheet.Cells[wave_row + 1, XLS_Table.Q] = "=B" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.R] = "=C" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.S] = "=D" + row;
                        _sheet.Cells[wave_row + 1, XLS_Table.T] = "=E" + row;

                        Excel.Range main_range = _sheet.Range["A" + row];
                        Excel.Range hyper = _sheet.Range["P" + (wave_row + 1)];
                        // A to B
                        _sheet.Hyperlinks.Add(main_range, "#'" + _sheet.Name + "'!P" + (wave_row + 1));
                        _sheet.Hyperlinks.Add(hyper, "#'" + _sheet.Name + "'!A" + row);

                        // Past Delay time waveform
                        _range = _sheet.Range["P" + (wave_row + 2), "X" + (wave_row + 15)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_DT");

                        _range = _sheet.Range["Z" + (wave_row + 2), "AH" + (wave_row + 15)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_SST");

                        _range = _sheet.Range["AJ" + (wave_row + 2), "AR" + (wave_row + 15)];
                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_OFF");
#endif

                        InsControl._scope.Trigger_CH1();
                        InsControl._scope.TimeBasePosition(scope_time_scale * 3);
                        InsControl._scope.TimeScale(scope_time_scale);
                        MyLib.Delay1s(1);
                        row++;
                        wave_row += 20;
                    }
                }
            }

        Stop:
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
#if Report_en
            string str_temp = _sheet.Cells[2, 2].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;

            Mylib.SaveExcelReport(test_parameter.waveform_path, temp + "C_DT_SST_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
            if (!test_parameter.all_en && !test_parameter.chamber_en) delegate_mess.Invoke();
#endif

        }




        public void Power_Off()
        {
            if (test_parameter.trigger_vin_en)
            {
                InsControl._power.AutoPowerOff();
                System.Threading.Thread.Sleep(250);
                RTDev.GpEn_Disable();
            }
            else if (test_parameter.trigger_en)
            {
                RTDev.GpEn_Disable();
                System.Threading.Thread.Sleep(250);
                InsControl._power.AutoPowerOff();
            }
        }
    }
}
