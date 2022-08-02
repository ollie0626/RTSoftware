using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using InsLibDotNet;

namespace BuckTool
{
    public class MyLib
    {
        public static List<double> DGData(DataGridView dataGrid)
        {
            List<double> data = new List<double>();

            for(int row_idx = 0; row_idx < dataGrid.RowCount; row_idx++)
            {
                double start = Convert.ToDouble(dataGrid[0, row_idx].Value);
                double step = Convert.ToDouble(dataGrid[1, row_idx].Value);
                double stop = Convert.ToDouble(dataGrid[2, row_idx].Value);
                double res = 0;
                for (int idx = 0; res < stop; idx++)
                {
                    res = start + step * idx;
                    data.Add(res);
                }
            }
            return data;
        }

        public static List<double> TBData(TextBox tb)
        {
            List<double> data = new List<double>();
            string[] str_data = tb.Text.Split(',');
            double start = Convert.ToDouble(str_data[0]);
            double stop = Convert.ToDouble(str_data[1]);
            double step = 0.1;
            double res = 0;

            for (int idx = 0; res < stop; idx++)
            {
                res = start + step * idx;
                data.Add(res);
            }
            return data;
        }



        public string[] ListBinFile(string path)
        {
            string[] binList = new string[1];
            try
            {
                if (Directory.Exists(test_parameter.binFolder))
                {
                    binList = Directory.GetFiles(test_parameter.binFolder, "*.bin");
                    List<int> numList = new List<int>();
                    // full map
                    for (int i = 0; i < binList.Length; i++)
                    {
                        //map.Add(i, binList[i]);
                        string res = Path.GetFileNameWithoutExtension(binList[i]);
                        int idx_of = res.IndexOf("_");
                        numList.Add(Convert.ToInt16(res.Substring(0, idx_of)));
                    }

                    for (int i = 1; i < binList.Length; i++)
                    {
                        string res = binList[i];
                        int temp = numList[i];
                        int j = i - 1;

                        while (j > -1 && temp < numList[j])
                        {
                            numList[j + 1] = numList[j];
                            binList[j + 1] = binList[j];
                            j--;
                        }
                        numList[j + 1] = temp;
                        binList[j + 1] = res;
                    }
                }
                else
                {
                    return null;
                }
            }
            catch
            {
                if (Directory.Exists(test_parameter.binFolder))
                {
                    binList = Directory.GetFiles(test_parameter.binFolder, "*.bin");
                }
                else
                {
                    return null;
                }
            }

            return binList;
        }

        public void ExcelReportInit(Excel.Worksheet sheet)
        {
            Excel.Range range;
            sheet.Cells[1, 1] = "Item";
            range = sheet.Range["A1"];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Interior.Color = Color.FromArgb(255, 218, 185);

            range = sheet.Range["B1", "M1"];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Merge();

            sheet.Cells[2, 1] = "Conditions";
            range = sheet.Range["A2", "A16"];
            range.Merge();
            range.Interior.Color = Color.FromArgb(255, 218, 185);
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            range = sheet.Range["B2", "M16"];
            range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Merge();

            range = sheet.Rows["20:20"];
            range.Interior.Color = Color.AliceBlue;
            range.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle = Excel.XlLineStyle.xlContinuous;
            range.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
        }

        public void SaveExcelReport(string path, string fileName, Excel.Workbook _book)
        {
            string buf = path.Substring(path.Length - 1, 1) == @"\" ? path.Substring(0, path.Length - 1) : path;
            buf = buf + @"\" + fileName + ".xlsx";
            System.Threading.Thread.Sleep(2000);
            if (Directory.Exists(path))
            {
                _book.SaveAs(buf);
            }
            else
            {
                Directory.CreateDirectory(path);
                _book.SaveAs(buf);
            }
        }

        static public Excel.Chart CreateChart(Excel.Worksheet sheet, Excel.Range range, string title, string x_title, string y_title)
        {
            double top = range.Top;
            double left = range.Left;
            double width = range.Width;
            double height = range.Height;

            Excel.Chart page;
            Excel.ChartObjects objects = sheet.ChartObjects();
            Excel.ChartObject obj = objects.Add(left, top, width, height);
            page = obj.Chart;

            page.ChartWizard(
                System.Type.Missing,
                Excel.XlChartType.xlXYScatterSmooth,
                System.Type.Missing,
                Excel.XlRowCol.xlColumns,
                0, 0, true,
                title, x_title, y_title,
                System.Type.Missing
                );
            return page;
        }


        public void testCondition(Excel.Worksheet sheet, string item, int bin_cnt, double temperature)
        {
            string conditions = "";
            sheet.Cells[1, 2] = item;

            string str_vin = "";
            foreach (double temp in test_parameter.Vin_table)
            {
                str_vin += string.Format("{0:0.##}V, ", temp);
            }

            string str_iout = "";
            //foreach (double temp in test_parameter.Iout_table)
            //{
            //    str_iout += string.Format("{0:0.##}A, ", temp);
            //}
            str_iout = string.Format("{0}A ~ {1}A, ", test_parameter.Iout_table[0], test_parameter.Iout_table[test_parameter.Iout_table.Count - 1]);

            conditions += "Temperature = " + temperature.ToString() + "\r\n";
            conditions += "Vin= " + str_vin + "\r\n";
            conditions += "Iout= " + str_iout + "\r\n";
            conditions += "bin file number = " + bin_cnt.ToString();
            sheet.Cells[2, 2] = conditions;
        }

        public static bool Vincompensation(double targetV, ref double reV)
        {
            /* current vol : vin, target offset : offset */
            double vin = Convert.ToDouble(string.Format("{0:##.0000}", InsControl._34970A.Get_100Vol(1)));
            double offset = targetV - vin;

            if (Math.Abs(vin) < 0.5)
            {
                //MessageBox.Show("Please check 34970A connecttion");
                //Console.WriteLine("Please check 34970A connecttion");
                return false;
            }

            if (vin < (targetV * 1.0002) || vin > (targetV * 0.9998))
                if (vin > (targetV * 1.0002) || vin < (targetV * 0.9998))
                {
                    reV += (float)offset;
                    InsControl._power.AutoSelPowerOn(reV);
                    System.Threading.Thread.Sleep(800);
                    vin = Convert.ToDouble(string.Format("{0:##.0000}", InsControl._34970A.Get_100Vol(1)));
                    Console.WriteLine(vin);

                    if (Math.Abs(targetV - vin) < 0.003) goto DONE;
                    if (vin < (targetV * 1.0002) || vin > (targetV * 0.9998))
                        if (vin > (targetV * 1.0002) || vin < (targetV * 0.9998))
                        {
                            Vincompensation(targetV, ref reV);
                        }
                }
            DONE:;
            return true;
        }

        public static void Relay_Reset(bool is400mA)
        {
            InsControl._eload.AllChannel_LoadOff();
            InsControl._eload.CH1_Loading(0);

            if(is400mA)
            {
                InsControl._dmm1.ChangeCurrentLevel(is400mA);
                InsControl._dmm2.ChangeCurrentLevel(is400mA);
                RTBBControl.Meter400mA(RTBBControl.GPIO2_0);
                RTBBControl.Meter400mA(RTBBControl.GPIO2_1);
            }
            else
            {
                InsControl._dmm1.ChangeCurrentLevel(!is400mA);
                InsControl._dmm2.ChangeCurrentLevel(!is400mA);
                RTBBControl.Meter10A(RTBBControl.GPIO2_0);
                RTBBControl.Meter10A(RTBBControl.GPIO2_1);
            }

        }

        public static void Switch_ELoadLevel(double level)
        {
            if (level < 0.1)
                InsControl._eload.CCL_Mode();
            else if (level > 0.1 && level < 1)
                InsControl._eload.CCM_Mode();
            else
                InsControl._eload.CCH_Mode();
        }

        public static void WaveformCheck()
        {
            InsControl._scope.DoCommand("*CLS");
            while (!(InsControl._scope.doQeury(":ADER?") == "+1\n")) ;
        }

        public static void ProcessCheck()
        {
            InsControl._scope.DoCommand("*CLS");
            while (!(InsControl._scope.doQeury(":PDER?") == "+1\n")) ;
        }

        public static void Relay_Process(int port, double curr_cmp, int iout_idx, int vin_idx, bool isIin, bool sw400mA, ref bool en)
        {
            double meter_limit = 0.4 * 0.75;
            if(curr_cmp > meter_limit && !en)
            {
                double vin = InsControl._power.GetVoltage();
                InsControl._power.AutoPowerOff();
                InsControl._eload.AllChannel_LoadOff();
                if (isIin) InsControl._dmm1.ChangeCurrentLevel(sw400mA);
                else InsControl._dmm2.ChangeCurrentLevel(sw400mA);
                RTBBControl.Meter10A(port);

                
                InsControl._power.AutoSelPowerOn(test_parameter.Vin_table[vin_idx]);
                InsControl._eload.CH1_Loading(test_parameter.Iout_table[iout_idx]);

                en = true;
            }
        }

        public static void Delay1ms(int cnt)
        {
            if (cnt < 1) return;
            System.Threading.Thread.Sleep(cnt);
        }

        public static void Delay1s(int cnt)
        {
            if (cnt < 1) return;
            Delay1ms(cnt * 1000);
        }

        public static void FuncGen_Fixedparameter(double freq, double duty, double tr, double tf)
        {
            InsControl._funcgen.CH1_ContinuousMode();
            InsControl._funcgen.CH1_PulseMode();
            InsControl._funcgen.CH1_Frequency(freq);
            InsControl._funcgen.CH1_DutyCycle(duty);
            InsControl._funcgen.SetCH1_TrTfFunc(tr, tf);
            InsControl._funcgen.CHl1_HiLevel(0.1);
            InsControl._funcgen.CH1_LoLevel(0);
            InsControl._funcgen.CH1_LoadImpedanceHiz();
        }

        public static void FuncGen_loopparameter(double hi, double lo)
        {
            InsControl._funcgen.CHl1_HiLevel(hi);
            InsControl._funcgen.CH1_LoLevel(lo);
            InsControl._funcgen.CH1_On();
        }
    }
}
