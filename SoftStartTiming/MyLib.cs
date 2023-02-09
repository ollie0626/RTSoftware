using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace SoftStartTiming
{

    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
        BA, BB, BC, BD, BE ,BF, BG, BH, BI, BJ, BK, BL, BM, BN, BO, BP, BQ, BR, BS, BT, BU, BV, BW, BX, BY, BZ
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
                if (Directory.Exists(path))
                {
                    binList = Directory.GetFiles(path, "*.bin");
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
                if (Directory.Exists(path))
                {
                    binList = Directory.GetFiles(path, "*.bin");
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
            if (InsControl._tek_scope_en) return;
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
            if (InsControl._tek_scope_en) return;
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


        public static void Switch_ELoadLevel(double level)
        {
            if (level < 0.1)
                InsControl._eload.CCL_Mode();
            else if (level > 0.1 && level < 1)
                InsControl._eload.CCM_Mode();
            else
                InsControl._eload.CCH_Mode();
        }


        public static void PastWaveform(Excel.Worksheet _sheet, Excel.Range _range, string wavePath, string fileName)
        {
            string buf = wavePath.Substring(wavePath.Length - 1, 1) == @"/" ? wavePath.Substring(0, wavePath.Length - 1) : wavePath;
            buf = buf + @"/" + fileName + ".png";


            double left = _range.Left;
            double top = _range.Top;
            double width = _range.Width;
            double height = _range.Height;

            _sheet.Shapes.AddPicture(wavePath + "\\" + fileName + ".png",
                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                    Microsoft.Office.Core.MsoTriState.msoTrue,
                                    (float)left, (float)top, (float)width, (float)height);
        }


        public static double GetCriteria_time(string info)
        {
            double res = 0;
            int idx = info.IndexOf("check");
            if (idx == -1) return 0;
            string tmp = info.Substring(idx);
            string Criteria = tmp.Split('_')[1];
            string data;
            Console.WriteLine(tmp);
            Console.WriteLine(Criteria);
            double unit;
            if (Criteria.IndexOf("ms") != -1)
            {
                unit = Math.Pow(10, -3);
                idx = Criteria.IndexOf("ms");
                data = Criteria.Substring(0, idx);
                Console.WriteLine(data);
            }
            else if (Criteria.IndexOf("us") != -1)
            {
                unit = Math.Pow(10, -6);
                idx = Criteria.IndexOf("us");
                data = Criteria.Substring(0, idx);
                Console.WriteLine(data);
            }
            else // s
            {
                unit = 1;
                idx = Criteria.IndexOf("s");
                data = Criteria.Substring(0, idx);
            }
            res = Convert.ToDouble(data) * unit;
            return res;
        }


        public static double GetCriteria_vol(string info)
        {
            double res = 0;
            int idx = info.IndexOf("check");
            if (idx == -1) return 0;
            string tmp = info.Substring(idx);
            string Criteria = tmp.Split('_')[1];
            string data;
            Console.WriteLine(tmp);
            Console.WriteLine(Criteria);
            double unit;
            if (Criteria.IndexOf("mV") != -1)
            {
                unit = Math.Pow(10, -3);
                idx = Criteria.IndexOf("mV");
                data = Criteria.Substring(0, idx);
                Console.WriteLine(data);
            }
            else // V
            {
                unit = 1;
                idx = Criteria.IndexOf("V");
                data = Criteria.Substring(0, idx);
            }
            res = Convert.ToDouble(data) * unit;
            return res;
        }

    }
}
