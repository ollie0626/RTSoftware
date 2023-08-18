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
//using System.Threading;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

using System;
using System.Timers;

namespace SoftStartTiming
{
    public class ATE_LTLab : TaskRun
    {
        System.Timers.Timer timer;
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        RTBBControl RTDev = new RTBBControl();

        //private Thread clear_cnt;
        private bool sel;
        //private int temp_meas = 8;

        private int meas_vmean = 1;
        private int meas_vmax = 2;
        private int meas_vmin = 3;

        private int meas_imean = 4;
        private int meas_iduty = 5;
        private int meas_ifreq = 6;

        private void OSCInit()
        {

            InsControl._oscilloscope.SetMeasureSource(1, meas_vmean, "MEAN");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmax, "MAXimum");
            InsControl._oscilloscope.SetMeasureSource(1, meas_vmin, "MINImum");

            InsControl._oscilloscope.SetMeasureSource(4, meas_imean, "MEAN");
            InsControl._oscilloscope.SetMeasureSource(4, meas_iduty, "PDUty");
            InsControl._oscilloscope.SetMeasureSource(4, meas_ifreq, "FREQuency");

            InsControl._oscilloscope.CHx_Level(1, 0.01);
            InsControl._oscilloscope.CHx_Position(1, 0);

        }

        private void RefleshMeasure(Object source, ElapsedEventArgs e)
        {
            //InsControl._oscilloscope.DoCommand("MEASUrement:STATIstics:COUNt RESET");
            Console.WriteLine("1s send clear counter command");
        }

        private bool I2C_Check(int match_idx)
        {
            byte addr = test_parameter.lt_lab.addr_list[match_idx];
            byte data = test_parameter.lt_lab.data_list[match_idx];
            byte[] buf = new byte[1];
            RTDev.I2C_Read(test_parameter.slave, addr, ref buf);
            return (buf[0] == data);
        }

        private void CHxResize(int idx)
        {
            //InsControl._oscilloscope.CHx_Offset(1, test_parameter.lt_lab.vout_list[idx]);
            //MyLib.Delay1ms(200);
            //InsControl._oscilloscope.CHx_Meas_Max(1, meas_vmax);
            //MyLib.Delay1ms(200);
            //double max = InsControl._oscilloscope.CHx_Meas_Max(1, meas_vmax);
            //InsControl._oscilloscope.CHx_Level(1, max / 3);
        }

        private void TimerInit()
        {
            // Create a timer with a 1 second interval
            timer = new System.Timers.Timer(1000);
            timer.Elapsed += RefleshMeasure;
            timer.AutoReset = true;
        }



        public override void ATETask()
        {
            RTDev.BoadInit();
            TimerInit();
            int row = 1;

            string path = Application.StartupPath + "\\example.xlsm";
            #region "Report Initial"
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = _app.Workbooks.Open(path);                      // open example excel
            _sheet = (Excel.Worksheet)_book.ActiveSheet;              // raw data sheet

            // Excel initial
            //_app = new Excel.Application();
            //_app.Visible = true;
            //_book = (Excel.Workbook)_app.Workbooks.Add();
            //_sheet = (Excel.Worksheet)_book.ActiveSheet;
#endif
            #endregion
            //OSCInit();

            // data cnt & vin cnt as same         
            for (int i2c_idx = 0; i2c_idx < test_parameter.lt_lab.data_list.Count; i2c_idx++)
            {

#if Report_en
                row = 1;
                string sheet_name = string.Format("Vin={0:0}_Vout={1:##.##}_{2:X2}={3:X2}",
                                            test_parameter.VinList[i2c_idx],
                                            test_parameter.lt_lab.vout_list[i2c_idx],
                                            test_parameter.lt_lab.addr_list[i2c_idx],
                                            test_parameter.lt_lab.data_list[i2c_idx]);
                if(i2c_idx != 0)  _sheet = (Excel.Worksheet)_book.Worksheets.Add();
                _sheet.Name = sheet_name;
                _sheet.Cells[row, XLS_Table.A] = "Vin (V)";
                _sheet.Cells[row, XLS_Table.B] = "Iin (A)";
                _sheet.Cells[row, XLS_Table.C] = "VMax (V)";
                _sheet.Cells[row, XLS_Table.D] = "VMin (V)";
                _sheet.Cells[row, XLS_Table.E] = "VMean (V)";
                _sheet.Cells[row, XLS_Table.F] = "IMean (A)";
                _sheet.Cells[row, XLS_Table.G] = "IDuty (%)";
                _sheet.Cells[row, XLS_Table.H] = "IFreq (Hz)";
                row++;

#endif
                InsControl._oscilloscope.SetAutoTrigger();
                InsControl._oscilloscope.CHx_Position(1, 0);
                InsControl._oscilloscope.CHx_Offset(1, test_parameter.lt_lab.vout_list[i2c_idx]); // vout offset
                InsControl._oscilloscope.CHx_Level(1, 0.01); // set level 10mV
                InsControl._power.AutoSelPowerOn(test_parameter.VinList[i2c_idx]);
                while (!I2C_Check(i2c_idx)) { MyLib.Delay1ms(50); }

                while (I2C_Check(i2c_idx))
                {
                    //InsControl._oscilloscope.SetTimeScale(50 * Math.Pow(10, -9));
                    //CHxResize(i2c_idx);
                    //InsControl._oscilloscope.SetPERSistence();

                    timer.Enabled = true;
                    //InsControl._oscilloscope.SetClear();
                    MyLib.Delay1ms(200);

                    double vin = 0;
                    double Iin = 0, vmax = 0, vmin = 0, vmean = 0;
                    double imean = 0, iduty = 0, ifreq = 0;

                    vin = InsControl._power.GetVoltage();
                    Iin = InsControl._power.GetCurrent();
                    vmax = InsControl._oscilloscope.CHx_Meas_Max(1, meas_vmax);
                    vmin = InsControl._oscilloscope.CHx_Meas_Min(1, meas_vmin);
                    vmean = InsControl._oscilloscope.CHx_Meas_Mean(1, meas_vmean);
                    imean = InsControl._oscilloscope.CHx_Meas_Mean(4, meas_imean);
                    iduty = InsControl._oscilloscope.CHx_Meas_Duty(4, meas_iduty);
                    ifreq = InsControl._oscilloscope.CHx_Meas_Freq(4, meas_ifreq);

                    #region "meas data"
#if Report_en
                    _sheet.Cells[row, XLS_Table.A] = vin;
                    _sheet.Cells[row, XLS_Table.B] = Iin;           // power supply
                    _sheet.Cells[row, XLS_Table.C] = vmax;          // vout max
                    _sheet.Cells[row, XLS_Table.D] = vmin;          // vout min
                    _sheet.Cells[row, XLS_Table.E] = vmean;         // vout mean
                    _sheet.Cells[row, XLS_Table.F] = imean;         // Iin mean
                    _sheet.Cells[row, XLS_Table.G] = iduty;         // Iout duty
                    _sheet.Cells[row, XLS_Table.H] = ifreq;         // Iout freq
                    row++;
#endif
                    #endregion
                };
                //"MEASUrement:STATIstics:COUNt RESET"
                //MyLib.Delay1ms(300);
            }

#if Report_en
            timer.Enabled = false;
            MyLib.SaveExcelReport(test_parameter.waveform_path, temp + "C_LTLab_" + DateTime.Now.ToString("yyyyMMdd_hhmm"), _book);
            _book.Close(false);
            _book = null;
            _app.Quit();
            _app = null;
            GC.Collect();
#endif
        }
    }
}

