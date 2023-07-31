#define Report_en
#define Power_en
#define Eload_en

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
using System.Threading.Tasks;
using System.Threading;
using System.Runtime.InteropServices;

namespace SoftStartTiming
{
    public class ATE_LoadTransient : TaskRun
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

        private void Funcgen_HiLo( double hi, double lo)
        {
            InsControl._funcgen.CHl1_HiLevel(hi);
            InsControl._funcgen.CH1_LoLevel(lo);
        }

        private void AdjustCurrent(double hi, double lo, bool eload = false)
        {
            if(eload)
            {
                double t1 = test_parameter.loadtransient.T1;
                double t2 = test_parameter.loadtransient.T2;
                InsControl._eload.DymanicCH1(hi, lo, t1, t2);
            }
            else
            {
                double gain = test_parameter.loadtransient.gain;
                double hi_cal = hi * gain;
                double lo_cal = lo * gain;
                Funcgen_HiLo(hi_cal, lo_cal);
            }
        }

        private void AdjustTrigger(double hi, double lo)
        {
            double trigger_level = lo + (hi - lo) * 0.7;
            InsControl._oscilloscope.SetTriggerRise();
            InsControl._oscilloscope.SetTimeOutTriggerCHx(4);
            InsControl._oscilloscope.SetTriggerLevel(4, trigger_level);
        }

        private void CHxResize()
        {
            // channel1 measure vpp
            // channel2 measure vmax
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


        private void ZoomIn(bool rising = true)
        {
            if (rising)
            {

            }
            else
            {

            }
        }

        private void PrintExcelTitle(ref int row)
        {
            _sheet.Cells[row, XLS_Table.B] = "超連結";
            _sheet.Cells[row, XLS_Table.C] = "Temp(C)";
            _sheet.Cells[row, XLS_Table.D] = "Vin(V)";
            _sheet.Cells[row, XLS_Table.E] = "Iload(mA)";
            _sheet.Cells[row, XLS_Table.F] = "Data(Hex)";
            _sheet.Cells[row, XLS_Table.G] = "Vout(V)";
            _sheet.Cells[row, XLS_Table.H] = "Vpp(mV)";
            _sheet.Cells[row, XLS_Table.I] = "Vmin(mV)";
            _sheet.Cells[row, XLS_Table.J] = "Vmax(mV)";
            _sheet.Cells[row, XLS_Table.K] = "Imax (mA)";
            _sheet.Cells[row, XLS_Table.L] = "Imin (mA)";
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
            // eload: true, evb: false
            sel = test_parameter.loadtransient.eload_dev_sel;
            RTDev.BoadInit();
            double hi_current = 0;
            double lo_current = 0;
            double vin = 0;
            double freq = test_parameter.loadtransient.freq;
            double duty = test_parameter.loadtransient.duty / 100;
            double period = 1 / freq;
            double on_time = period * duty;
            int case_num = 1;
            int row = 14;
            int wave_row = 14;


            #region "Report Initial"
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            _sheet.Cells.Font.Name = "Calibri";
            _sheet.Cells.Font.Size = 11;

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

            string item = "Load transient";

            _sheet.Cells[1, XLS_Table.B] = item;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.tool_ver + test_parameter.vin_conditions;
#endif
            #endregion

            OSCInit();

            for (int vin_idx = 0; vin_idx < test_parameter.VinList.Count; vin_idx++)
            {
                PrintExcelTitle(ref row); row++;

                for (int hi_idx = 0; hi_idx < test_parameter.loadtransient.hi_current.Count; hi_idx++)
                {
                    for(int lo_idx = 0; lo_idx < test_parameter.loadtransient.lo_current.Count; lo_idx++)
                    {
                        // initial time scale
                        InsControl._oscilloscope.SetTimeScale(on_time / 5);

                        string file_name = ""; 
                        double vpp, vmax, vmin, vout;
                        double imax, imin;
                        vin = test_parameter.VinList[vin_idx];
                        hi_current = test_parameter.loadtransient.hi_current[hi_idx];
                        lo_current = test_parameter.loadtransient.lo_current[lo_idx];

                        file_name = string.Format("{0}_Temp={1}_Vin={2}_Iout={3}_{4}", temp, 14 - row, test_parameter.VinList[vin_idx], hi_current, lo_current);

                        // eload: true, evb: false
#if Power_en
                        InsControl._power.AutoSelPowerOn(vin);
#endif

#if Eload_en
                        AdjustTrigger(hi_current, lo_current);
                        AdjustCurrent(hi_current, lo_current, sel);
#endif
                        MyLib.Delay1ms(500);
                        CHxResize();
                        InsControl._oscilloscope.SetPERSistence();
                        InsControl._oscilloscope.SetNormalTrigger();
                        InsControl._oscilloscope.SetClear();

                        vpp = InsControl._oscilloscope.MeasureMean(meas_vpp);
                        vmax = InsControl._oscilloscope.MeasureMean(meas_vmax);
                        vmin = InsControl._oscilloscope.MeasureMean(meas_vmin);
                        vout = InsControl._oscilloscope.MeasureMean(meas_vout);
                        imax = InsControl._oscilloscope.MeasureMean(meas_imax);
                        imin = InsControl._oscilloscope.MeasureMean(meas_imin);

                        #region "meas data"
#if Report_en
                        _sheet.Cells[row, XLS_Table.B] = "LINK";
                        _sheet.Cells[row, XLS_Table.C] = temp;
                        _sheet.Cells[row, XLS_Table.D] = test_parameter.VinList[vin_idx];
                        _sheet.Cells[row, XLS_Table.E] = lo_current.ToString() + "->" + hi_current.ToString();
                        _sheet.Cells[row, XLS_Table.F] = "";
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

                        double tr = test_parameter.loadtransient.Tr;
                        double tf = test_parameter.loadtransient.Tf;
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
                                    
                                    while(cnt < test_parameter.meas_cnt)
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
                                    time_scale = (tr * Math.Pow(10, -9)) / 3;
                                    InsControl._oscilloscope.SetTimeScale(time_scale);
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
                                    time_scale = (tf * Math.Pow(10, -9)) / 3;
                                    InsControl._oscilloscope.SetTimeScale(time_scale);
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
}
