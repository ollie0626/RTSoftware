

#define Report
#define Power_en
#define Eload_en
#define Scope_en


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

    public class ATE_SoftStartTime : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        Excel.Range _range_fall;
        //Excel.Chart _chart;

        //public double temp;
        MyLib Mylib = new MyLib();
        RTBBControl RTDev = new RTBBControl();
        //TestClass tsClass = new TestClass();
        public delegate void FinishNotification();
        FinishNotification delegate_mess;

        //_sheet.Cells[row, XLS_Table.I] = "SST (us)";
        //_sheet.Cells[row, XLS_Table.J] = "V1 Max (V)";
        //_sheet.Cells[row, XLS_Table.K] = "V1 Min (V)";
        //_sheet.Cells[row, XLS_Table.L] = "ILx Max (mA)";
        //_sheet.Cells[row, XLS_Table.M] = "ILx Min (mA)";

        int meas_rising = 1;
        int meas_vmax = 2;
        int meas_vmin = 3;
        int meas_imax = 4;
        int meas_imin = 5;
        //int meas_level = 6;
        int meas_falling = 7;


        public ATE_SoftStartTime()
        {
            delegate_mess = new FinishNotification(MessageNotify);
        }

        private void MessageNotify()
        {
            System.Windows.Forms.MessageBox.Show("Delay time/Soft start time test finished!!!", "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
        }

        private void GpioOnSelect(int num)
        {
            //RTDev.GPIOnState((uint)1 << num, (uint)1 << num);
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Enable();
                    break;
                case 1:
                    RTDev.Gp2En_Enable();
                    break;
                case 2:
                    RTDev.Gp3En_Enable();
                    break;
            }
        }

        private void GpioOffSelect(int num)
        {

            //RTDev.GPIOnState((uint)1 << num, (uint)(~(1 << num)));
            switch (num)
            {
                case 0:
                    RTDev.Gp1En_Disable();
                    break;
                case 1:
                    RTDev.Gp2En_Disable();
                    break;
                case 2:
                    RTDev.Gp3En_Disable();
                    break;
            }
        }

        private void I2C_DG_Write(DataGridView dg)
        {
            for (int i = 0; i < dg.RowCount; i++)
            {
                byte addr = Convert.ToByte(dg[0, i].Value.ToString(), 16);
                byte data = Convert.ToByte(dg[1, i].Value.ToString(), 16);
                RTDev.I2C_Write((byte)(test_parameter.slave), addr, new byte[] { data });
                MyLib.Delay1ms(50);
            }
        }

        private void OSCInit()
        {
            InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
            InsControl._tek_scope.DoCommand("HORizontal:ROLL OFF");
            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");

            InsControl._tek_scope.SetTimeBasePosition(35);
            InsControl._tek_scope.SetRun();
            InsControl._tek_scope.SetTriggerMode();
            InsControl._tek_scope.SetTriggerSource(2);
            InsControl._tek_scope.SetTriggerLevel(1.5);

            InsControl._tek_scope.CHx_On(1);
            InsControl._tek_scope.CHx_On(2);
            InsControl._tek_scope.CHx_On(3);
            InsControl._tek_scope.CHx_On(4);

            InsControl._tek_scope.CHx_Level(1, 1.65);
            InsControl._tek_scope.CHx_Level(2, test_parameter.VinList[0]);
            InsControl._tek_scope.CHx_Level(3, test_parameter.LX_Level);
            InsControl._tek_scope.CHx_Level(4, test_parameter.ILX_Level);

            InsControl._tek_scope.CHx_Position(1, 0);
            InsControl._tek_scope.CHx_Position(3, -1);
            InsControl._tek_scope.CHx_Position(4, -3);
            InsControl._tek_scope.CHx_Position(2, -1);

            InsControl._tek_scope.CHx_BWlimitOn(1);
            InsControl._tek_scope.CHx_BWlimitOn(2);
            InsControl._tek_scope.CHx_BWlimitOn(3);
            InsControl._tek_scope.CHx_BWlimitOn(4);

            InsControl._tek_scope.DoCommand("MEASUrement:MEAS1:REFLevel:METHod PERCent");
            InsControl._tek_scope.DoCommand("MEASUrement:MEAS1:REFLevel:PERCent:HIGH 90");
            InsControl._tek_scope.DoCommand("MEASUrement:MEAS1:REFLevel:PERCent:MID 50");
            InsControl._tek_scope.DoCommand("MEASUrement:MEAS1:REFLevel:PERCent:LOW 10");

            InsControl._tek_scope.SetMeasureSource(2, meas_rising, "RISe");
            InsControl._tek_scope.SetMeasureSource(2, meas_vmax, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(2, meas_vmin, "MINImum");
            InsControl._tek_scope.SetMeasureSource(4, meas_imax, "MAXimum");
            InsControl._tek_scope.SetMeasureSource(4, meas_imin, "MINImum");
            InsControl._tek_scope.SetMeasureSource(2, meas_falling, "FALL");
            InsControl._tek_scope.PersistenceDisable();
            MyLib.Delay1ms(500);
        }

        private void Scope_Channel_Resize(int idx, string path)
        {

            InsControl._tek_scope.SetRun();
            InsControl._tek_scope.SetTriggerMode();

#if Power_en
            InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
            MyLib.Delay1ms(800);
            RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_dis });
#endif

            double time_scale = 0;
            time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");
            InsControl._tek_scope.SetTimeScale(Math.Pow(10, -9) * 40);



            switch (test_parameter.trigger_event)
            {
                case 0: // gpio
                    InsControl._tek_scope.SetTriggerLevel(1);
                    InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                    InsControl._tek_scope.CHx_Position(1, 2.5);

                    if (test_parameter.sleep_mode)
                        GpioOnSelect(test_parameter.gpio_pin);
                    else
                        GpioOffSelect(test_parameter.gpio_pin);
                    break;
                case 1: // i2c trigger

                    InsControl._tek_scope.SetTriggerLevel(1);
                    InsControl._tek_scope.CHx_Level(1, 3.3 / 2);
                    InsControl._tek_scope.CHx_Position(1, 2.5);

                    I2C_DG_Write(test_parameter.i2c_init_dg); // write initial code
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                    break;
                case 2: // vin trigger
                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[idx]);
                    InsControl._tek_scope.SetTriggerLevel(test_parameter.VinList[idx] * 0.35);
                    break;
            }
            
            MyLib.Delay1s(1);
            RTDev.I2C_WriteBin((byte)(test_parameter.slave), 0x00, path); // test conditions
            I2C_DG_Write(test_parameter.i2c_mtp_dg); // mtp program
            MyLib.Delay1s(1); // wait for program time
            if (test_parameter.trigger_event == 1)
            {
                // i2c trigger
                I2C_DG_Write(test_parameter.i2c_init_dg);
                RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
            }
            MyLib.Delay1ms(800);
            InsControl._tek_scope.CHx_Level(2, 1.5);
            InsControl._tek_scope.CHx_Level(3, test_parameter.LX_Level);
            InsControl._tek_scope.CHx_Level(4, test_parameter.ILX_Level);
            MyLib.Delay1s(4);
            int re_cnt = 0;
            for (int ch_idx = 0; ch_idx < 3; ch_idx++)
            {
            re_scale:;
                if (re_cnt > 3)
                {
                    re_cnt = 0;
                    continue;
                }

                double vmax = 0;
                vmax = InsControl._tek_scope.CHx_Meas_MAX(ch_idx + 2, meas_vmax);
                vmax = InsControl._tek_scope.CHx_Meas_MAX(ch_idx + 2, meas_vmax);
                MyLib.Delay1ms(100);
                vmax = InsControl._tek_scope.CHx_Meas_MAX(ch_idx + 2, meas_vmax);

                // catch wrong data, reset initial condition
                if (vmax > Math.Pow(10, 9))
                {
                    re_cnt++;
                    InsControl._tek_scope.CHx_Level(ch_idx + 2, test_parameter.VinList[0]);
                    MyLib.Delay1ms(800);
                    goto re_scale;
                }

                int ch = ch_idx + 2;
                switch (ch)
                {
                    case 2:
                        // vout setting
                        InsControl._tek_scope.CHx_Level(ch, vmax / 5.5);                    
                        InsControl._tek_scope.SetTriggerSource(2);
                        InsControl._tek_scope.SetTriggerLevel(vmax / 3);
                        InsControl._tek_scope.CHx_Position(ch, -2);
                        break;
                    case 3:
                        // lx setting
                        InsControl._tek_scope.CHx_Level(ch, test_parameter.VinList[0] / 1.5);
                        InsControl._tek_scope.CHx_Position(ch, -3);
                        break;
                    case 4:
                        // ilx setting
                        InsControl._tek_scope.CHx_Level(ch, test_parameter.ILX_Level);
                        break;
                }

                MyLib.Delay1ms(800);
            }
            //// set trigger 50%
            //InsControl._tek_scope.DoCommand("TRIGger:A");
            MyLib.Delay1ms(800);
            PowerOffEvent();
            InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
            InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
            InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
            InsControl._tek_scope.SetTimeBasePosition(35);
            MyLib.Delay1ms(250);
        }

        private bool TriggerStatus()
        {
            int cnt = 0;
            while (InsControl._tek_scope.GetCount() == 0)
            {
                cnt++;
                MyLib.Delay1ms(50);
                if (cnt > 100) return false;
            }
            return true;
        }

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            RTDev.BoadInit();
            RTDev.GpioInit();

            int vin_cnt = test_parameter.VinList.Count;
            int iout_cnt = test_parameter.IoutList.Count;
            int row = 8;
            int wave_row = 8;
            int wave_pos = 0;
            string[] binList;
            double[] ori_vinTable = new double[vin_cnt];
            int bin_cnt = 1;
            Array.Copy(test_parameter.VinList.ToArray(), ori_vinTable, vin_cnt);

#if Report
            // Excel initial
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
#endif


#if Power_en
            InsControl._power.AutoPowerOff();
#endif
            OSCInit();
            MyLib.Delay1s(1);
            int cnt = 0;
            #region "Report initial"
#if Report
            _sheet = _book.Worksheets.Add();
            _sheet.Name = "SST Test";
            row = 8;
            wave_row = 8;
            wave_pos = 0;
            _sheet.Cells[1, XLS_Table.A] = "Item";
            _sheet.Cells[2, XLS_Table.A] = "Test Conditions";
            _sheet.Cells[3, XLS_Table.A] = "Result";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _range = _sheet.Range["A1", "A4"];
            _range.Font.Bold = true;
            _range.Interior.Color = Color.FromArgb(255, 178, 102);
            _range = _sheet.Range["A2"];
            _range.RowHeight = 150;
            _range = _sheet.Range["B1"];
            _range.ColumnWidth = 60;
            _range = _sheet.Range["A1", "B4"];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // print test conditions
            _sheet.Cells[1, XLS_Table.B] = "Soft-Start time";
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions + test_parameter.bin_file_cnt;

            _sheet.Cells[row, XLS_Table.D] = "No.";
            _sheet.Cells[row, XLS_Table.E] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.F] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.G] = "Iout(A)";
            _sheet.Cells[row, XLS_Table.H] = "Bin file";
            _range = _sheet.Range["D" + row, "H" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            // major measure timing
            _sheet.Cells[row, XLS_Table.I] = "SST_Rise (us)";
            _sheet.Cells[row, XLS_Table.J] = "V1 Max_Rise (V)";
            _sheet.Cells[row, XLS_Table.K] = "V1 Min_Rise (V)";
            _sheet.Cells[row, XLS_Table.L] = "ILx Max_Rise (mA)";
            _sheet.Cells[row, XLS_Table.M] = "ILx Min_Rise (mA)";
            _sheet.Cells[row, XLS_Table.N] = "Pass/Fail";

            _sheet.Cells[row, XLS_Table.O] = "SST_Fall (us)";
            _sheet.Cells[row, XLS_Table.P] = "V1 Max_Fall (V)";
            _sheet.Cells[row, XLS_Table.Q] = "V1 Min_Fall (V)";
            _sheet.Cells[row, XLS_Table.R] = "ILx Max_Fall (mA)";
            _sheet.Cells[row, XLS_Table.S] = "ILx Min_Fall (mA)";
            _sheet.Cells[row, XLS_Table.T] = "Pass/Fail";

            _range = _sheet.Range["I" + row, "M" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["O" + row, "S" + row];
            _range.Interior.Color = Color.FromArgb(30, 144, 255);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["N" + row, "N" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            _range = _sheet.Range["T" + row, "T" + row];
            _range.Interior.Color = Color.FromArgb(124, 252, 0);
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            row++;
#endif
            #endregion

            stopWatch.Start();
            binList = MyLib.ListBinFile(test_parameter.bin_path[0]);
            bin_cnt = binList.Length;
            cnt = 0;

            for (int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
            {

                
                InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
                InsControl._tek_scope.SetTimeBasePosition(35);

                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {

                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {

                        int retry_cnt = 0;
                        double iout = test_parameter.IoutList[iout_idx];

                        if (test_parameter.run_stop == true) goto Stop;
                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                        /* test initial setting */
                        string file_name;
                        string res = Path.GetFileNameWithoutExtension(binList[bin_idx]);
                        //test_parameter.sleep_mode = (res.IndexOf("sleep_en") == -1) ? false : true;

#if Eload_en
                        if (test_parameter.eload_cr)
                        {
                            if (iout > 80) InsControl._eload.CRL_Mode();
                            else if (iout > 80 && iout < 2900) InsControl._eload.CRM_Mode();
                            else if (iout > 2900) InsControl._eload.CRH_Mode();

                            InsControl._eload.DoCommand("CHAN 1");
                            InsControl._eload.SetCR(iout);
                        }
                        else
                        {
                            MyLib.Switch_ELoadLevel(iout);
                            InsControl._eload.CH1_Loading(iout);
                        }
#endif
                        MyLib.Delay1s(1);


                        InsControl._tek_scope.SetMeasureSource(2, meas_rising, "RISe");
                        InsControl._tek_scope.SetMeasureSource(2, meas_vmax, "MAXimum");
                        InsControl._tek_scope.SetMeasureSource(2, meas_vmin, "MINImum");
                        InsControl._tek_scope.SetMeasureSource(4, meas_imax, "MAXimum");
                        InsControl._tek_scope.SetMeasureSource(4, meas_imin, "MINImum");
                        InsControl._tek_scope.SetMeasureSource(2, meas_falling, "FALL");


                        MyLib.Delay1ms(500);
                        Console.WriteLine(res);
                        file_name = string.Format("{0}_Temp={2}C_vin={3:0.##}V_{1}",
                                                    cnt, res, temp,
                                                    test_parameter.VinList[vin_idx]
                                                    );

                        double time_scale = 0;
                        time_scale = InsControl._tek_scope.doQueryNumber("HORizontal:SCAle?");


                        // include test condition
                        Scope_Channel_Resize(vin_idx, binList[bin_idx]);
                        double tempVin = ori_vinTable[vin_idx];

                        MyLib.WaveformCheck();
                        MyLib.Delay1ms(800);
                        InsControl._tek_scope.SetTriggerMode(false);
                        InsControl._tek_scope.SetTriggerRise();
                        InsControl._tek_scope.SetClear();
                        InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                        MyLib.Delay1ms(1500);

                        // power on trigger
                        PowerOnEvent();
                        while (!TriggerStatus()) ;
                        InsControl._tek_scope.SetStop();
                        MyLib.Delay1ms(1000);

                        double delay_time = 0;

                        delay_time = InsControl._tek_scope.CHx_Meas_Rise(2, 1);
                        delay_time = InsControl._tek_scope.CHx_Meas_Rise(2, 1);
                        MyLib.Delay1ms(100);
                        delay_time = InsControl._tek_scope.CHx_Meas_Rise(2, 1);

                        double temp_time = 0;
                        temp_time = (delay_time / 4);
                        InsControl._tek_scope.SetTimeScale(temp_time);
                        InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                        InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");


                        //PowerOffEvent();
                        //RTDev.I2C_WriteBin((byte)(test_parameter.slave), 0x00, binList[bin_idx]); // test conditions
                        MyLib.Delay1ms(1800);

                        InsControl._tek_scope.SetRun();
                        //InsControl._tek_scope.SetClear();

                        //MyLib.Delay1ms(1500);
                        //PowerOnEvent();
                        //MyLib.Delay1s(1);
                        InsControl._tek_scope.SetStop();


#if Scope_en
                        double vin = 0;
                        double sst = 0, vmax = 0, vmin = 0, ilx_max = 0, ilx_min = 0;
#if Power_en
                        vin = InsControl._power.GetVoltage();
#endif
                        sst = InsControl._tek_scope.CHx_Meas_Rise(2, meas_rising) * Math.Pow(10, 6);
                        vmax = InsControl._tek_scope.CHx_Meas_MAX(2, meas_vmax);
                        vmin = InsControl._tek_scope.CHx_Meas_MIN(2, meas_vmin);
                        ilx_max = InsControl._tek_scope.CHx_Meas_MAX(4, meas_imax);
                        ilx_min = InsControl._tek_scope.CHx_Meas_MIN(4, meas_imin);

                        InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                        MyLib.Delay1ms(100);

                        InsControl._tek_scope.DoCommand("MEASUrement:ANNOTation:STATE MEAS1"); MyLib.Delay1ms(200);
                        MyLib.Delay1ms(1000);

                        for (int i = 0; i < 2; i++)
                        {
                            double x1 = InsControl._tek_scope.doQueryNumber("MEASUrement:ANNOTation:X1?"); MyLib.Delay1ms(250);
                            double x2 = InsControl._tek_scope.doQueryNumber("MEASUrement:ANNOTation:X2?"); MyLib.Delay1ms(250);

                            InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 " + x1);
                            MyLib.Delay1ms(250);
                            double data = InsControl._tek_scope.CHx_Meas_Rise(2, 1) * 0.9;
                            MyLib.Delay1ms(250);
                            InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + x2);
                            MyLib.Delay1ms(250);
                        }


                        MyLib.Delay1s(1);
                        InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_rise");

#if Report
                        _sheet.Cells[row, XLS_Table.D] = cnt++;
                        _sheet.Cells[row, XLS_Table.E] = temp;
                        _sheet.Cells[row, XLS_Table.F] = vin;
                        _sheet.Cells[row, XLS_Table.G] = iout;
                        _sheet.Cells[row, XLS_Table.H] = res;

                        _sheet.Cells[row, XLS_Table.I] = sst;
                        _sheet.Cells[row, XLS_Table.J] = vmax;
                        _sheet.Cells[row, XLS_Table.K] = vmin;
                        _sheet.Cells[row, XLS_Table.L] = ilx_max;
                        _sheet.Cells[row, XLS_Table.M] = ilx_min;
#endif
                        double criteria = MyLib.GetCriteria_time(res);
                        criteria = criteria * Math.Pow(10, 6);
                        double criteria_up = (test_parameter.judge_percent * criteria) + criteria;
                        double criteria_down = criteria - (test_parameter.judge_percent * criteria);
                        Console.WriteLine(criteria);

#if Report
                        if (sst > criteria_up || sst < criteria_down)
                        {
                            _sheet.Cells[row, XLS_Table.N] = "Fail";
                            _range = _sheet.Range["N" + row];
                            _range.Interior.Color = Color.Red;
                        }
                        else
                        {
                            _sheet.Cells[row, XLS_Table.N] = "Pass";
                            _range = _sheet.Range["N" + row];
                            _range.Interior.Color = Color.LightGreen;
                        }
#endif
#endif

#if Scope_en
                        // SST Fall test
                        //PowerOnEvent();



                        InsControl._tek_scope.SetRun();
                        InsControl._tek_scope.SetTriggerMode(false);
                        InsControl._tek_scope.SetTriggerFall();
                        InsControl._tek_scope.SetClear();
                        InsControl._tek_scope.SetTimeScale(test_parameter.ontime_scale_ms / 1000);
                        MyLib.Delay1ms(1000);

                        PowerOffEvent();
                        while (!TriggerStatus());

                        MyLib.Delay1ms(100);
                        delay_time = InsControl._tek_scope.MeasureMax(meas_falling);
                        temp_time = (delay_time / 4);
                        InsControl._tek_scope.SetTimeScale(temp_time);
                        InsControl._tek_scope.DoCommand("HORizontal:MODE AUTO");
                        InsControl._tek_scope.DoCommand("HORizontal:MODE:SAMPLERate 500E6");
#if Power_en
                        vin = InsControl._power.GetVoltage();
#endif
                        sst = InsControl._tek_scope.MeasureMax(meas_falling) * Math.Pow(10, 6);
                        vmax = InsControl._tek_scope.MeasureMax(meas_vmax);
                        vmin = InsControl._tek_scope.MeasureMin(meas_vmin);
                        ilx_max = InsControl._tek_scope.MeasureMax(meas_imax);
                        ilx_min = InsControl._tek_scope.MeasureMin(meas_imin);

                        InsControl._tek_scope.DoCommand("CURSor:FUNCtion WAVEform");
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce1 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:SOUrce2 CH2");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:MODe TRACk");
                        MyLib.Delay1ms(100);
                        InsControl._tek_scope.DoCommand("CURSor:STATE ON");
                        MyLib.Delay1ms(100);

                        InsControl._tek_scope.DoCommand("MEASUrement:ANNOTation:STATE MEAS7"); MyLib.Delay1ms(200);
                        MyLib.Delay1ms(1000);

                        for (int i = 0; i < 2; i++)
                        {
                            double x1 = InsControl._tek_scope.doQueryNumber("MEASUrement:ANNOTation:X1?"); MyLib.Delay1ms(250);
                            double x2 = InsControl._tek_scope.doQueryNumber("MEASUrement:ANNOTation:X2?"); MyLib.Delay1ms(250);

                            InsControl._tek_scope.DoCommand("CURSor:VBArs:POS1 " + x1);
                            MyLib.Delay1ms(250);
                            double data = InsControl._tek_scope.CHx_Meas_Rise(2, 1) * 0.9;
                            MyLib.Delay1ms(250);
                            InsControl._tek_scope.DoCommand("CURSor:VBArs:POS2 " + x2);
                            MyLib.Delay1ms(250);
                        }

#if Report
                        _sheet.Cells[row, XLS_Table.O] = sst;
                        _sheet.Cells[row, XLS_Table.P] = vmax;
                        _sheet.Cells[row, XLS_Table.Q] = vmin;
                        _sheet.Cells[row, XLS_Table.R] = ilx_max;
                        _sheet.Cells[row, XLS_Table.S] = ilx_min;
                        //_sheet.Cells[row, XLS_Table.T] = "Pass/Fail";
#endif

                        MyLib.Delay1s(1);
                        InsControl._tek_scope.SaveWaveform(test_parameter.waveform_path, file_name + "_fall");

#endif
#if Report
                        switch (wave_pos)
                        {
                            case 0:
                                _sheet.Cells[wave_row, XLS_Table.AA] = "No.";
                                _sheet.Cells[wave_row, XLS_Table.AB] = "Temp(C)";
                                _sheet.Cells[wave_row, XLS_Table.AC] = "Vin(V)";
                                _sheet.Cells[wave_row, XLS_Table.AD] = "Bin file";
                                _sheet.Cells[wave_row, XLS_Table.AE] = "Iout(A)";
                                _range = _sheet.Range["AA" + wave_row, "AE" + wave_row];
                                _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                _sheet.Cells[wave_row + 1, XLS_Table.AA] = "=D" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AB] = "=E" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AC] = "=F" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AD] = "=H" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AE] = "=G" + row;
                                _range = _sheet.Range["AA" + (wave_row + 2).ToString(), "AI" + (wave_row + 16).ToString()];
                                _range_fall = _sheet.Range["AK" + (wave_row + 2).ToString(), "AS" + (wave_row + 16).ToString()];
                                wave_pos++;
                                break;

                            case 1:
                                _sheet.Cells[wave_row, XLS_Table.AW] = "No.";
                                _sheet.Cells[wave_row, XLS_Table.AX] = "Temp(C)";
                                _sheet.Cells[wave_row, XLS_Table.AY] = "Vin(V)";
                                _sheet.Cells[wave_row, XLS_Table.AZ] = "Bin file";
                                _sheet.Cells[wave_row, XLS_Table.BA] = "Iout(A)";
                                _range = _sheet.Range["AW" + wave_row, "AZ" + wave_row];
                                _range.Interior.Color = Color.FromArgb(124, 252, 0);
                                _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                                _sheet.Cells[wave_row + 1, XLS_Table.AW] = "=D" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AX] = "=E" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AY] = "=F" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.AZ] = "=H" + row;
                                _sheet.Cells[wave_row + 1, XLS_Table.BA] = "=G" + row;
                                _range = _sheet.Range["AW" + (wave_row + 2).ToString(), "BE" + (wave_row + 16).ToString()];
                                _range_fall = _sheet.Range["BG" + (wave_row + 2).ToString(), "BO" + (wave_row + 16).ToString()];
                                wave_pos = 0;
                                wave_row = wave_row + 19;
                                break;
                        }

                        MyLib.PastWaveform(_sheet, _range, test_parameter.waveform_path, file_name + "_rise");
                        MyLib.PastWaveform(_sheet, _range_fall, test_parameter.waveform_path, file_name + "_fall");
#endif

                        row++;
#if Eload_en
                        InsControl._eload.CH1_Loading(0);
                        InsControl._eload.LoadOFF(1);
#endif
                    } // iout loop
                } // bin loop
            } // vin loop
        Stop:
            stopWatch.Stop();
            // record test finish time
            stopWatch.Stop();
            TimeSpan timeSpan = stopWatch.Elapsed;
#if Report
            string str_temp = _sheet.Cells[2, XLS_Table.B].Value;
            string time = string.Format("{0}h_{1}min_{2}sec", timeSpan.Hours, timeSpan.Minutes, timeSpan.Seconds);
            str_temp += "\r\n" + time;
            _sheet.Cells[2, 2] = str_temp;
            //TimeSpan timeSpan = stopWatch.Elapsed;

            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_SST_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }

        private void PowerOnEvent()
        {
            switch (test_parameter.trigger_event)
            {
                case 0:
                    // GPIO trigger event
                    if (test_parameter.sleep_mode)
                    {
                        GpioOnSelect(test_parameter.gpio_pin);
                        MyLib.Delay1ms(1000);
                    }
                    else
                    {
                        GpioOffSelect(test_parameter.gpio_pin);
                        MyLib.Delay1ms(1000);
                    }
                    break;
                case 1:
                    // I2C trigger event
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_en });
                    break;
                case 2:
#if Power_en
                    // Power supply trigger event
                    InsControl._power.AutoPowerOff();
#endif
                    break;
            }
        }

        private void PowerOffEvent()
        {
            switch (test_parameter.trigger_event)
            {
                case 0: // gpio power disable
                    if (test_parameter.sleep_mode)
                        GpioOffSelect(test_parameter.gpio_pin);
                    else
                        GpioOnSelect(test_parameter.gpio_pin);
                    break;
                case 1:
                    // rails disable
                    RTDev.I2C_Write((byte)(test_parameter.slave), test_parameter.Rail_addr, new byte[] { test_parameter.Rail_dis });
                    //I2C_DG_Write(test_parameter.i2c_init_dg);
                    break;
                case 2: // vin trigger
#if Power_en
                    InsControl._power.AutoPowerOff();
#endif
                    break;
            }
        }

    }
}





