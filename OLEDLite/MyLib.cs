using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace OLEDLite
{
    public class MyLib
    {
        static System.Timers.Timer timer = new System.Timers.Timer();
        static private int timer_cnt = 0;

        public MyLib()
        {
            timer.Enabled = false;
            timer.Interval = 100;
            timer.Elapsed += OnTickEvent;
            timer_cnt = 0;
            timer.Stop();
        }

        private void OnTickEvent(object sender, System.Timers.ElapsedEventArgs e)
        {
            timer_cnt += 1;
        }

        public static List<double> DGData(DataGridView dataGrid)
        {
            List<double> data = new List<double>();

            for (int row_idx = 0; row_idx < dataGrid.RowCount; row_idx++)
            {
                double start = Convert.ToDouble(dataGrid[0, row_idx].Value);
                double step = Convert.ToDouble(dataGrid[1, row_idx].Value);
                double stop = Convert.ToDouble(dataGrid[2, row_idx].Value);
                double res = 0;
                for (int idx = 0; res < stop; idx++)
                {
                    res = start + step * idx;

                    if(res > stop)
                    {
                        data.Add(stop);
                        break;
                    }

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
            double step = Convert.ToDouble(str_data[2]);
            //double step = 0.1;
            double res = 0;

            for (int idx = 0; res < stop; idx++)
            {
                res = start + step * idx;
                data.Add(res);
            }
            return data;
        }

        public static string[] ListBinFile(string path)
        {
            string[] binList = new string[1];
            try
            {
                if (Directory.Exists(test_parameter.bin_path))
                {
                    binList = Directory.GetFiles(test_parameter.bin_path, "*.bin");
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
                if (Directory.Exists(test_parameter.bin_path))
                {
                    binList = Directory.GetFiles(test_parameter.bin_path, "*.bin");
                }
                else
                {
                    return null;
                }
            }

            return binList;
        }

        public static void SaveExcelReport(string path, string fileName, Excel.Workbook _book)
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

        static public Excel.Chart CreateChart(Excel.Worksheet sheet, Excel.Range range, string title, string x_title, string y_title, bool isXY = false)
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
                isXY ? Excel.XlChartType.xlXYScatterSmooth : Excel.XlChartType.xlColumnClustered,
                //Excel.XlChartType.xlXYScatterSmooth,
                System.Type.Missing,
                Excel.XlRowCol.xlColumns,
                0, 0, true,
                title, x_title, y_title,
                System.Type.Missing
                );

            return page;
        }

        public static bool Vincompensation(double targetV, ref double reV)
        {
            /* current vol : vin, target offset : offset */
            double vin = Convert.ToDouble(string.Format("{0:##.0000}", InsControl._34970A.Get_100Vol(1)));
            double offset = targetV - vin;

            if (Math.Abs(vin) < 0.5 || vin >= test_parameter.vin_threshold)
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

        public static void WaveformCheck()
        {
            timer.Start();
            InsControl._scope.DoCommand("*CLS");
            while (!(InsControl._scope.doQeury(":ADER?") == "+1\n"))
            {
                Delay1ms(150);
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

        public static void Channel_LevelSetting(int channel)
        {
            int idx = 0;
            double issue_num = 9.99999 * Math.Pow(10, 10);
            for (int i = 0; i <= 8; i++)
            {
                string info = "";
                double avg = 0;
                double vmax = InsControl._scope.Measure_Ch_Max(channel);
                System.Threading.Thread.Sleep(100);
                double vmin = InsControl._scope.Measure_Ch_Min(channel);
                System.Threading.Thread.Sleep(100);
                double temp = InsControl._scope.Measure_Ch_Vpp(channel);
                System.Threading.Thread.Sleep(100);

                if (vmax < 0)
                {
                    InsControl._scope.CHx_Level(channel, 3);
                    continue;
                }

                if (vmax > issue_num || temp > issue_num)
                {
                    InsControl._scope.CHx_Level(channel, 3);
                    continue;
                }

                avg = vmax - vmin;
                InsControl._scope.CHx_Level(channel, avg / 3);
                System.Threading.Thread.Sleep(300);
            }
        }

        public static void EloadFixChannel()
        {
            for(int i = 0; i < test_parameter.eload_iout.Length; i++)
            {
                if(test_parameter.eload_en[i])
                {
                    switch(i)
                    {
                        case 0:
                            InsControl._eload.CH1_Loading(test_parameter.eload_iout[i]);
                            break;
                        case 1:
                            InsControl._eload.CH2_Loading(test_parameter.eload_iout[i]);
                            break;
                        case 2:
                            InsControl._eload.CH3_Loading(test_parameter.eload_iout[i]);
                            break;
                        case 3:
                            InsControl._eload.CH4_Loading(test_parameter.eload_iout[i]);
                            break;
                    }
                }
            }
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

        public static void Switch_ELoadLevel(double level)
        {
            if (level < 0.1)
                InsControl._eload.CCL_Mode();
            else if (level > 0.1 && level < 1)
                InsControl._eload.CCM_Mode();
            else
                InsControl._eload.CCH_Mode();
        }

    }
}
