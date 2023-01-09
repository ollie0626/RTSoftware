using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Drawing;


namespace OLEDLite
{
    public class ATE_IDNL : TaskRun
    {
        Excel.Application _app;
        Excel.Worksheet _sheet;
        Excel.Workbook _book;
        Excel.Range _range;
        MyLib MyLib;
        RTBBControl RTDev = new RTBBControl();


        public override void ATETask()
        {
            Stopwatch stopWatch = new Stopwatch();
            stopWatch.Start();
            MyLib = new MyLib();
            int row = 11;
            int idx = 0;
            int bin_cnt = 1;
            string[] binList = new string[1];
            binList = MyLib.ListBinFile(test_parameter.bin_path);
            bin_cnt = binList.Length;
            bool ispos = test_parameter.vol_max > test_parameter.vol_min;
            int vin_cnt = test_parameter.vinList.Count;
            int iout_cnt = test_parameter.ioutList.Count;
            double[] ori_vinTable = new double[vin_cnt];
            Array.Copy(test_parameter.vinList.ToArray(), ori_vinTable, vin_cnt);
            RTDev.BoadInit();
#if Report
            //MyLib.ExcelReportInit(_sheet);
            //MyLib.testCondition(_sheet, "Code Inrush", bin_cnt, temp);
            _app = new Excel.Application();
            _app.Visible = true;
            _book = (Excel.Workbook)_app.Workbooks.Add();
            _sheet = (Excel.Worksheet)_book.ActiveSheet;
            string eload_condition = test_parameter.ioutList[0] + " ~ " + test_parameter.ioutList[test_parameter.ioutList.Count - 1];
            string swire_condition = "Swire:" + test_parameter.code_min + "→" + test_parameter.code_max;
            string vin_condition = "";
            for (int i = 0; i < test_parameter.vinList.Count; i++)
            {
                if (i == test_parameter.vinList.Count - 1) vin_condition += test_parameter.vinList[i];
                else vin_condition += test_parameter.vinList[i] + ",";
            }
            _sheet.Cells[1, XLS_Table.A] = "Vin";
            _sheet.Cells[2, XLS_Table.A] = "Iout";
            _sheet.Cells[3, XLS_Table.A] = "Date";
            _sheet.Cells[4, XLS_Table.A] = "Note";
            _sheet.Cells[5, XLS_Table.A] = "Version";
            _sheet.Cells[6, XLS_Table.A] = "Temperatrue";
            _sheet.Cells[7, XLS_Table.A] = "test time";

            _sheet.Cells[1, XLS_Table.B] = test_parameter.vin_info;
            _sheet.Cells[2, XLS_Table.B] = test_parameter.eload_info;
            _sheet.Cells[3, XLS_Table.B] = test_parameter.date_info;
            _sheet.Cells[5, XLS_Table.B] = test_parameter.ver_info;
            _sheet.Cells[6, XLS_Table.B] = temp;
#endif
            for(int swire_idx = 0; swire_idx < test_parameter.swire_cnt; swire_idx++)
            {
                for(int vin_idx = 0; vin_idx < vin_cnt; vin_idx++)
                {
                    for (int iout_idx = 0; iout_idx < iout_cnt; iout_idx++)
                    {

                        //for(int pulse_idx = 0;)



                    }
                }
            }





            stopWatch.Stop();
        }
    }
}
