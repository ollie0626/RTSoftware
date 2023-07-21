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

        private void OSCInit()
        {
            
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

        private void ZoomIn(bool rising = true)
        {

        }

        public override void ATETask()
        {
            // eload: true, evb: false
            sel = test_parameter.loadtransient.eload_dev_sel;
            RTDev.BoadInit();
            double hi_current = 0;
            double lo_current = 0;
            double vin = 0;

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

            for (int vin_idx = 0; vin_idx < test_parameter.vin_conditions.Length; vin_idx++)
            {
                for(int hi_idx = 0; hi_idx < test_parameter.loadtransient.hi_current.Count; hi_idx++)
                {
                    for(int lo_idx = 0; lo_idx < test_parameter.loadtransient.lo_current.Count; lo_idx++)
                    {
                        vin = test_parameter.VinList[vin_idx];
                        hi_current = test_parameter.loadtransient.hi_current[hi_idx];
                        lo_current = test_parameter.loadtransient.lo_current[lo_idx];
                        // eload: true, evb: false
                        InsControl._power.AutoSelPowerOn(vin);
                        AdjustTrigger(hi_current, lo_current);
                        AdjustCurrent(hi_current, lo_current, sel);
                    }
                }
            }


        }
    }
}
