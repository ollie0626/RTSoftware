using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;

namespace SDCTool
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


    public class SDCTesting : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;

        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;

#if Report
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
#endif
            InsControl._power.AutoPowerOff();



            for (int vin_idx = 0; vin_idx < test_parameter.vinList.Count; vin_idx++)
            {
                for (int bin_idx = 0; bin_idx < bin_cnt; bin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < test_parameter.ioutList.Count; iout_idx++)
                    {
                        if (test_parameter.run_stop == true) goto Stop;
                        if ((bin_idx % 5) == 0 && test_parameter.chamber_en == true) InsControl._chamber.GetChamberTemperature();

                        InsControl._power.AutoSelPowerOn(test_parameter.vinList[vin_idx]);


                    } // eload loop
                } // bin loop 
            } // power loop


        Stop:
            stopWatch.Stop();
        }
    }
}
