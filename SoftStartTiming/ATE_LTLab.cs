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
using System.Threading;
using System.Runtime.InteropServices;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;

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

        public override void ATETask()
        {
            RTDev.BoadInit();
            int row = 2;

            string path = Application.StartupPath + "\\example.xlsm";
            #region "Report Initial"
#if Report_en
            _app = new Excel.Application();
            _app.Visible = true;
            _book = _app.Workbooks.Open(path);                      // open example excel
            _sheet = (Excel.Worksheet)_book.ActiveSheet;              // raw data sheet
#endif
            #endregion
            OSCInit();

            // data cnt & vin cnt as same         
            for (int i2c_idx = 0; i2c_idx < test_parameter.lt_lab.data_list.Count; i2c_idx++)
            {
                InsControl._power.AutoSelPowerOn(test_parameter.VinList[i2c_idx]);
                while (!I2C_Check(i2c_idx)) { MyLib.Delay1ms(50); };
                //InsControl._oscilloscope.SetTimeScale(50 * Math.Pow(10, -9));
                MyLib.Delay1ms(200);

                CHxResize(i2c_idx);
                //InsControl._oscilloscope.SetPERSistence();
                InsControl._oscilloscope.SetAutoTrigger();
                InsControl._oscilloscope.SetClear();

                //TimeScaleSetting();
                InsControl._oscilloscope.SetTimeScale(test_parameter.lt_lab.time_scale * Math.Pow(10, -3));
                MyLib.Delay1ms(200);

                double Iin, vmax, vmin, vmean;
                double imean, iduty, ifreq;

                Iin = InsControl._power.GetCurrent();
                vmax = InsControl._oscilloscope.CHx_Meas_Max(meas_vmax);
                vmin = InsControl._oscilloscope.CHx_Meas_Min(meas_vmin);
                vmean = InsControl._oscilloscope.CHx_Meas_Mean(meas_vmean);

                imean = InsControl._oscilloscope.CHx_Meas_Mean(meas_imean);
                iduty = InsControl._oscilloscope.CHx_Meas_Duty(meas_iduty);
                ifreq = InsControl._oscilloscope.CHx_Meas_Freq(meas_ifreq);

                #region "meas data"
#if Report_en
                _sheet.Cells[row, XLS_Table.A] = Iin;           // power supply
                _sheet.Cells[row, XLS_Table.B] = vmax;          // vout max
                _sheet.Cells[row, XLS_Table.C] = vmin;          // vout min
                _sheet.Cells[row, XLS_Table.D] = vmean;         // vout mean
                _sheet.Cells[row, XLS_Table.E] = imean;         // Iin mean
                _sheet.Cells[row, XLS_Table.F] = iduty;         // Iout duty
                _sheet.Cells[row, XLS_Table.G] = ifreq;         // Iout freq
                row++;
#endif
                #endregion
            }

#if Report_en
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

