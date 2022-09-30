using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using InsLibDotNet;
using System.IO;
using System.Timers;

namespace IN528ATE_tool
{
    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
    };

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

    public class MyLib
    {
        static Timer timer = new Timer();
        static private int timer_cnt = 0;
        
        public MyLib()
        {
            timer.Enabled = false;
            timer.Interval = 100;
            timer.Elapsed += OnTickEvent;
            timer_cnt = 0;
            timer.Stop();
        }

        private void OnTickEvent(object sender, ElapsedEventArgs e)
        {
            timer_cnt += 1;
        }


        public static void WaveformCheck()
        {
            timer.Start();
            InsControl._scope.DoCommand("*CLS");
            while (!(InsControl._scope.doQeury(":ADER?") == "+1\n"))
            {
                Delay1ms(50);
                if (timer_cnt >= 100)
                {
                    timer.Stop();
                    timer_cnt = 0;
                    Console.WriteLine("WaveformCheck time out !!!");
                    break;
                }
            }
            timer.Stop();
            timer_cnt = 0;
        }

        public static void ProcessCheck()
        {
            timer.Start();
            InsControl._scope.DoCommand("*CLS");
            while (!(InsControl._scope.doQeury(":PDER?") == "+1\n"))
            {
                Delay1ms(50);
                if (timer_cnt >= 100)
                {
                    timer.Stop();
                    timer_cnt = 0;
                    Console.WriteLine("ProcessCheck time out !!!");
                    break;
                }
            }

            timer.Stop();
            timer_cnt = 0;
        }


        public int calculate_cnt(double start, double stop, double step)
        {
            double offset;
            if (start > stop)
                offset = start - stop;
            else if (stop > start)
                offset = stop - start;
            else
                offset = 1;

            int cnt = (offset == 1) ? 1 : (int)(offset / step);
            return cnt;
        }

        public List<double> conditions(double start, double stop, double step)
        {
            List<double> temp = new List<double>();
            int cnt = calculate_cnt(start, stop, step);
            double max = start > stop ? start : stop;
            double min = start < stop ? start : stop;

            for(int i = 0; i < cnt; i++)
            {
                double res = 0;
                if(stop > start)
                {
                    res = start + i * step;
                    temp.Add(res);
                }
                else if(start > stop)
                {
                    res = start - i * step;
                    temp.Add(res);
                }
                else
                {
                    temp.Add(start);
                }
            }

            double last_val = temp[temp.Count - 1];
            if(stop > start)
            {
                if (last_val < max) temp.Add(max);
            }
            else if(start > stop)
            {
                if (last_val > min) temp.Add(min);
            }
            
            return temp;
        }

        public static void Delay1s(int cnt)
        {
            System.Threading.Thread.Sleep(1000 * cnt);
        }

        public static void Delay1ms(int cnt)
        {
            System.Threading.Thread.Sleep(cnt);
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

        public void eLoadLevelSwich(EloadModule eload, double iout)
        {
            if (iout < 0.15) eload.CCL_Mode();
            else if (iout >= 0.15 && iout < 1.5) eload.CCM_Mode();
            else eload.CCH_Mode();
        }

        public bool Vincompensation(PowerModule powerModule, MultiChannelModule _34970A, double targetV, ref double reV)
        {
            /* current vol : vin, target offset : offset */
            double vin = Convert.ToDouble(string.Format("{0:##.0000}", _34970A.Get_100Vol(1)));
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
                    powerModule.AutoSelPowerOn(reV);
                    System.Threading.Thread.Sleep(800);
                    vin = Convert.ToDouble(string.Format("{0:##.0000}", _34970A.Get_100Vol(1)));
                    Console.WriteLine(vin);

                    if (Math.Abs(targetV - vin) < 0.003) goto DONE;
                    if (vin < (targetV * 1.0002) || vin > (targetV * 0.9998))
                        if (vin > (targetV * 1.0002) || vin < (targetV * 0.9998))
                        {
                            Vincompensation(powerModule, _34970A, targetV, ref reV);
                        }
                }
            DONE:;
            return true;
        }

        public void Channel_LevelSetting(AgilentOSC scope, int channel)
        {
            int idx = 0;
            double temp = scope.Meas_CH1VPP();
            while ((temp > (9.999 * Math.Pow(10, 10))) && (idx <= 100))
            {
                scope.CH1_Level(0.05 + (idx * 0.05));
                temp = scope.Meas_CH1VPP();
                idx++;
                System.Threading.Thread.Sleep(10);
            }

            for(int i = 0; i <= 15; i++)
            {
                double avg = 0;
                double vmax = scope.Measure_Ch_Max(channel);
                double vmin = scope.Measure_Ch_min(channel);
                avg = vmax - vmin;
                scope.CHx_Level(channel, avg / 2);
                System.Threading.Thread.Sleep(200);
            }
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

        public Excel.Chart CreateChart(Excel.Worksheet sheet, Excel.Range range, string title, string x_axis, string y_axis)
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
                Excel.XlChartType.xlXYScatterLinesNoMarkers,
                System.Type.Missing,
                Excel.XlRowCol.xlColumns,
                0, 0, true,
                title, x_axis, y_axis,
                System.Type.Missing
                );
            return page;
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
        public void testCondition(Excel.Worksheet sheet, string item, int bin_cnt, double temperature)
        {
            string conditions = "";
            sheet.Cells[1, 2] =item;

            string str_vin = "";
            foreach (double temp in test_parameter.VinList)
            {
                str_vin += string.Format("{0:0.##}V, ", temp);
            }

            string str_iout = "";
            foreach (double temp in test_parameter.IoutList)
            {
                str_iout += string.Format("{0:0.##}A, ", temp);
            }

            conditions += "Temperature = " + temperature.ToString() + "\r\n";
            conditions += "Vin= " + str_vin + "\r\n";
            conditions += "Iout= " + str_iout + "\r\n";
            conditions += "bin file number = " + bin_cnt.ToString();
            sheet.Cells[2, 2] = conditions;
        }
    }


}
