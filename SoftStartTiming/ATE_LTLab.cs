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

        private bool I2C_Check(int match_idx)
        {
            byte addr = test_parameter.lt_lab.addr_list[match_idx];
            byte data = test_parameter.lt_lab.data_list[match_idx];
            byte[] buf = new byte[1];
            RTDev.I2C_Read(test_parameter.slave, addr, ref buf);
            return (buf[0] == data);
        }

        private void CHxResize()
        {
            double vpp = 0;
            double vmax = 0;
            InsControl._oscilloscope.SetMeasureSource(1, temp_meas, "PK2Pk");
            MyLib.Delay1ms(100);
            vpp = InsControl._oscilloscope.MeasureMax(temp_meas);
            InsControl._oscilloscope.CHx_Level(1, vpp / 2);


            InsControl._oscilloscope.SetMeasureSource(2, temp_meas, "MAXimum");
            vmax = InsControl._oscilloscope.MeasureMax(temp_meas);
            InsControl._oscilloscope.CHx_Level(2, vmax / 2);
        }

        private void TimeScaleSetting()
        {
            InsControl._oscilloscope.CHx_Meas_Period(4, 8);
            InsControl._oscilloscope.CHx_Meas_Duty(4, 7);
            double period = InsControl._oscilloscope.CHx_Meas_Period(4, 8);
            double duty = InsControl._oscilloscope.CHx_Meas_Duty(4, 7);
            double on_time = period * duty;
            double time_scale = on_time / 4.5;
            InsControl._oscilloscope.SetTimeScale(time_scale);
        }

        private void PrintExcelTitle(ref int row)
        {
            //_sheet.Cells[row, XLS_Table.B] = "超連結";
            //_sheet.Cells[row, XLS_Table.C] = "Temp(C)";
            //_sheet.Cells[row, XLS_Table.D] = "Vin(V)";
            //_sheet.Cells[row, XLS_Table.E] = "Iload(mA)";
            //_sheet.Cells[row, XLS_Table.F] = "Data(Hex)";
            //_sheet.Cells[row, XLS_Table.G] = "Vout(V)";
            //_sheet.Cells[row, XLS_Table.H] = "Vpp(mV)";
            //_sheet.Cells[row, XLS_Table.I] = "Vmin(mV)";
            //_sheet.Cells[row, XLS_Table.J] = "Vmax(mV)";
            //_sheet.Cells[row, XLS_Table.K] = "Imax (mA)";
            //_sheet.Cells[row, XLS_Table.L] = "Imin (mA)";
            //_sheet.Cells[row, XLS_Table.M] = "OverShoot(mV)";
            //_sheet.Cells[row, XLS_Table.N] = "UnderShoot(mV)";

            _range = _sheet.Range["B" + row.ToString(), "L" + row.ToString()];
            _range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            _range.Interior.Color = Color.FromArgb(0, 204, 0);
            _range = _sheet.Range["B" + row.ToString(), "F" + row.ToString()];
            _range.Interior.Color = Color.FromArgb(0xff, 0xff, 0x66);
        }


        public override void ATETask()
        {
            RTDev.BoadInit();

            int case_num = 1;
            int row = 14;
            int wave_row = 14;
            string file_name = "";
            int idx = 0;

            #region "Report Initial"
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Cells.Font.Name = "Calibri";
            _sheet.Cells.Font.Size = 11;

            _sheet.Cells[1, XLS_Table.A] = "Iin(A)";            // power supply
            _sheet.Cells[1, XLS_Table.B] = "Vout_Max(V)";       // vout max
            _sheet.Cells[1, XLS_Table.C] = "Vout_Min(V)";       // vout min
            _sheet.Cells[1, XLS_Table.D] = "Vout_Mean(V)";      // vout mean
            _sheet.Cells[1, XLS_Table.E] = "Iin_Mean(V)";       // Iin mean
            _sheet.Cells[1, XLS_Table.F] = "Iout_Duty_avg";     // Iout duty
            _sheet.Cells[1, XLS_Table.G] = "Iout_Freq_avg";     // Iout freq

            //_sheet.Cells[1, XLS_Table.A] = "Item";
            //_sheet.Cells[2, XLS_Table.A] = "Test Conditions";
            //_sheet.Cells[3, XLS_Table.A] = "Result";
            //_sheet.Cells[4, XLS_Table.A] = "Note";
            //_range = _sheet.Range["A1", "A4"];
            //_range.Font.Bold = true;
            //_range.Interior.Color = Color.FromArgb(255, 178, 102);
            //_range = _sheet.Range["A2"];
            //_range.RowHeight = 150;
            //_range = _sheet.Range["B1"];
            //_range.ColumnWidth = 60;
            //_range = _sheet.Range["A1", "B4"];
            //_range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            //string item = "Load transient";
            //_sheet.Cells[1, XLS_Table.B] = item;
            //_sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions;
#endif
            #endregion
            OSCInit();

            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                PrintExcelTitle(ref row); row++;

                for (int i2c_idx = 0; i2c_idx < test_parameter.lt_lab.data_list.Count; i2c_idx++)
                {

                    double vpp, vmax, vmin, vout;
                    double imax, imin;

                    InsControl._power.AutoSelPowerOn(test_parameter.VinList[vin_idx]);
                    while (!I2C_Check(i2c_idx)) { MyLib.Delay1ms(50); };
                    //InsControl._oscilloscope.SetTimeScale(50 * Math.Pow(10, -9));
                    MyLib.Delay1ms(200);
                    
                    CHxResize();
                    //TimeScaleSetting();
                    InsControl._oscilloscope.SetTimeScale(test_parameter.lt_lab.time_scale * Math.Pow(10, -6));
                    MyLib.Delay1ms(200);


                    InsControl._oscilloscope.SetPERSistence();
                    InsControl._oscilloscope.SetNormalTrigger();
                    InsControl._oscilloscope.SetClear();

                    vpp = InsControl._oscilloscope.MeasureMean(meas_vpp);
                    vmax = InsControl._oscilloscope.MeasureMean(meas_vmax);
                    vmin = InsControl._oscilloscope.MeasureMean(meas_vmin);
                    vout = InsControl._oscilloscope.MeasureMean(meas_vout);
                    imax = InsControl._oscilloscope.MeasureMean(meas_imax);
                    imin = InsControl._oscilloscope.MeasureMean(meas_imin);

                    file_name = string.Format("{0}_Temp={1}_Vin={2}_Iout{3}_{4}_I2C={5:X}",
                                                idx,
                                                temp,
                                                test_parameter.VinList[vin_idx],
                                                imin,
                                                imax,
                                                test_parameter.lt_lab.data_list[i2c_idx]);
                    #region "meas data"
#if Report_en
                    _sheet.Cells[row, XLS_Table.B] = "LINK";
                    _sheet.Cells[row, XLS_Table.C] = temp;
                    _sheet.Cells[row, XLS_Table.D] = test_parameter.VinList[vin_idx];
                    _sheet.Cells[row, XLS_Table.E] = imin.ToString() + "->" + imax.ToString();
                    _sheet.Cells[row, XLS_Table.F] = string.Format("{0:X}", test_parameter.lt_lab.data_list[i2c_idx]);
                    _sheet.Cells[row, XLS_Table.G] = string.Format("{0:0.000}", vout);
                    _sheet.Cells[row, XLS_Table.H] = string.Format("{0:0.000}", vpp);
                    _sheet.Cells[row, XLS_Table.I] = string.Format("{0:0.000}", vmax);
                    _sheet.Cells[row, XLS_Table.J] = string.Format("{0:0.000}", vmin);
                    _sheet.Cells[row, XLS_Table.K] = string.Format("{0:0.000}", imax);
                    _sheet.Cells[row, XLS_Table.L] = string.Format("{0:0.000}", imin);
                    row++;

                    _sheet.Cells[wave_row, XLS_Table.T] = "Go To Table";
                    _sheet.Cells[wave_row, XLS_Table.U] = "=C" + row;
                    _sheet.Cells[wave_row, XLS_Table.V] = "=D" + row;
                    _sheet.Cells[wave_row, XLS_Table.W] = "=E" + row;
                    _sheet.Cells[wave_row, XLS_Table.X] = "=F" + row;
                    _sheet.Cells[wave_row, XLS_Table.Y] = "=G" + row;
                    _sheet.Cells[wave_row, XLS_Table.Z] = "=H" + row;
                    _sheet.Cells[wave_row, XLS_Table.AA] = "=I" + row;
                    _sheet.Cells[wave_row, XLS_Table.AB] = "=J" + row;
                    _sheet.Cells[wave_row, XLS_Table.AC] = "=K" + row;
                    _sheet.Cells[wave_row, XLS_Table.AD] = "=L" + row;
#endif
                    #endregion


                    double time_scale = 0;
                    int cnt = 0;
                    for (int run_idx = 0; run_idx < case_num; run_idx++)
                    {
                        switch (run_idx)
                        {
                            case 0:
                                // past waveform cell range
                                _range = _sheet.Range["T" + (wave_row + 2), "AZ" + (wave_row + 16)];
                                InsControl._oscilloscope.SetClear();
                                InsControl._oscilloscope.SetTriggerRise();

                                while (cnt < test_parameter.meas_cnt)
                                {
                                    cnt = InsControl._oscilloscope.GetCount();
                                }
                                InsControl._oscilloscope.SetStop();
                                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, file_name);
                                break;
                            case 1:
                                InsControl._oscilloscope.SetClear();
                                InsControl._oscilloscope.SetTriggerRise();
                                _range = _sheet.Range["AC" + (wave_row + 2), "AI" + (wave_row + 16)];
                                //time_scale = (tr * Math.Pow(10, -9)) / 3;
                                InsControl._oscilloscope.CHx_Meas_Rise(4, 7);
                                time_scale = InsControl._oscilloscope.CHx_Meas_Rise(4, 7);
                                InsControl._oscilloscope.SetTimeScale(time_scale / 3);
                                while (cnt < test_parameter.meas_cnt)
                                {
                                    cnt = InsControl._oscilloscope.GetCount();
                                }
                                InsControl._oscilloscope.SetStop();
                                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, file_name + "_Rising");
                                break;
                            case 2:
                                InsControl._oscilloscope.SetClear();
                                InsControl._oscilloscope.SetTriggerFall();
                                //time_scale = (tf * Math.Pow(10, -9)) / 3;
                                InsControl._oscilloscope.CHx_Meas_Fall(4, 7);
                                time_scale = InsControl._oscilloscope.CHx_Meas_Fall(4, 7);
                                InsControl._oscilloscope.SetTimeScale(time_scale / 3);
                                _range = _sheet.Range["AL" + (wave_row + 2), "AR" + (wave_row + 16)];

                                while (cnt < test_parameter.meas_cnt)
                                {
                                    cnt = InsControl._oscilloscope.GetCount();
                                }
                                InsControl._oscilloscope.SetStop();
                                InsControl._oscilloscope.SaveWaveform(test_parameter.waveform_path, file_name + "_Falling");
                                break;
                        }

                        wave_row += 20;
                        InsControl._oscilloscope.SetPERSistenceOff();
                        InsControl._oscilloscope.SetRun();
                    }

                }
            }
        }
    }



}
