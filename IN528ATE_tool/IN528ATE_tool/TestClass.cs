using InsLibDotNet;
using System;
using System.Collections.Generic;
using System.Linq;

namespace IN528ATE_tool
{
    internal class TestClass
    {
        public double CalcSSTime(AgilentOSC Scope, double sampalerate = 0.000005)
        {
            double time = 0.0;
            if (Scope != null && Scope.InsState())
            {
                double[] Arrdata;
                Scope.SaveWaveformData(1, ref sampalerate, out Arrdata, true);
                //int N = 15;
                //if (N < 5) N = 5;
                if (Arrdata.Length < 100) return 0.0;
                List<double> WmaData = CalcWMA(Arrdata, 10);
                CalcNormalize(ref WmaData);
                WmaData.Sort();
                bool Is2Seg = Check2Segment(WmaData,50);
                double[] hist = Hist(WmaData);
                int[] MaxIdx = { 0, 0, 0};
                double[] MaxVal = { 0, 0, 0 };
                double SumMax = 0.0;
                for (int i = 0; i < 3; ++i)
                {
                    for(int j = 0; j < hist.Length; ++j)
                    {
                        if(hist[j] > hist[MaxIdx[i]])
                        {
                            MaxIdx[i] = j;
                        }
                    }
                    MaxVal[i] = hist[MaxIdx[i]];
                    hist[MaxIdx[i]] = 0.0;
                    SumMax += MaxVal[i];
                }

                double AveVal = (1 - SumMax) / (hist.Length - 3);

                Array.Sort(hist);
                bool flag = hist[hist.Length - 1] - hist[3] > 0.025;
                flag |= MaxVal[2] - hist[3] > 0.04;
                for (int i = 3; i < hist.Length; ++i)
                {
                    if (Math.Abs(hist[i] - AveVal) > 0.025)
                        flag = true;
                    else if (i > 3 && Math.Abs(hist[i] - hist[i - 1]) > 0.02)
                        flag = true;
                }

                int idx1 = FindSegIdx2(WmaData, 0, 0.1, 0.8, 1.01);
                int idx3 = FindSegIdx2(WmaData, 0.89, 1, 0.0, 0.4);
                int idx2 = FindSegIdx2(WmaData, MaxIdx[2]*0.1, 0.1*(MaxIdx[2]+1), 0.01, 0.4);

                double time1 = (idx2 - idx1) * sampalerate;
                double time2 = (idx3 - idx2 + 1) * sampalerate;

                if (flag || Is2Seg) time = time2;
                else time = time2 + time1;
            }
            return time;
        }

        public List<double> CalcWMA(double[] Arrdata, int N)
        {
            List<double> wma = new List<double>();

            if (Arrdata.Length < 100) return wma;
          
            for (int i = 10; i < (Arrdata.Length - N - 20); i += 1)
            {
                double tmpdata = 0.0;
                for (int j = 1; j <= N; ++j)
                {
                    tmpdata += j * Arrdata[j + i - 1];
                }
                wma.Add(Math.Abs(2.0 * tmpdata/N/(N+1)));
            }

            return wma;
        }

        public void CalcNormalize(ref List<double> Arrdata)
        {
            if (Arrdata.Count < 10) return;
            double MinVal = Arrdata.Min();
            double MaxVal = Arrdata.Max();
            double Width = MaxVal - MinVal;
            for (int i = 0; i < Arrdata.Count; ++i)
            {
                Arrdata[i] = (Arrdata[i] - MinVal) / Width;
            }
        }

        public bool Check2Segment(List<double> Arrdata, int N = 100)
        {
            if (Arrdata.Count < 10) return false;

            int[] Pos = { 0, 0, 0 };
            
            for(int i = 10; i < Arrdata.Count - 10; i += 2)
            {
                if (Pos[0] == 0 && Arrdata[i] >= 0.2) Pos[0] = i;
                else if (Pos[1] == 0 && Arrdata[i] >= 0.5) Pos[1] = i;
                else if (Pos[2] == 0 && Arrdata[i] >= 0.8)
                {
                    Pos[2] = i;
                    break;
                }
            }

            int Sg1 = Pos[1] - Pos[0];
            int Sg2 = Pos[2] - Pos[1];

            return Math.Abs(Sg1 - Sg2) > N;
        }

        public double[] Hist(List<double> Arrdata, int N = 10)
        {
            double[] hist = new double[N];

            for(int i = 0; i < N; ++i)
            {
                hist[i] = 0.0;
            }
            double step = 1.0 / N;
            for (int i = 0; i < Arrdata.Count; i ++)
            {
                for(int j = (N - 1); j >= 0; j--)
                {
                    if(Arrdata[i] < 0.00001)
                    {
                        hist[0]++;
                        break;
                    }
                    else if (Arrdata[i] > step * j)
                    {
                        hist[j]++;
                        break;
                    }
                }
            }

            for (int i = 0; i < N; ++i)
            {
                hist[i] /= Arrdata.Count;
            }

            return hist;
        }

        public int FindSegIdx(List<double> Arrdata, double low, double high,bool IsLow)
        {
            double[] hist = new double[3];

            for (int i = 0; i < 3; ++i)
            {
                hist[i] = 0.0;
            }
            double step = (high - low) / 3;

            for (int i = 0; i < Arrdata.Count; i++)
            {
                if (Arrdata[i] > high || Arrdata[i] < low) continue;
                for (int j = 2; j >= 0; j--)
                {
                    if(Arrdata[i] < low + 0.001)
                    {
                        hist[0]++;
                    }
                    if (Arrdata[i] > step * j + low)
                    {
                        hist[j]++;
                        break;
                    }
                }
            }

            int MaxIdx = hist[0] > hist[1] ? 0 : 1;
            MaxIdx = hist[2] > hist[MaxIdx] ? 2 : MaxIdx;

            if (MaxIdx > 0 && (hist[MaxIdx - 1] + hist[MaxIdx] * 0.25) > hist[MaxIdx]) MaxIdx--;
            if (IsLow) MaxIdx += 1;

            for (int i = 0; i < Arrdata.Count; i++)
            {
                if (Arrdata[i] > low + (MaxIdx) * step)
                    return i;
            }
            return Arrdata.Count;
        }

        public int FindSegIdx2(List<double> Arrdata, double low, double high, double StartCDF, double EndCDF)
        {
            int Cnt = 0;
            int Start = -1;
            
            for (int i = 0; i < Arrdata.Count; i++)
            {
                if (Arrdata[i] > high) break;
                else if(Arrdata[i] < low) continue;
                else
                {
                    if (Start < 0) Start = i;
                    Cnt++;
                }
            }
            double DeltaMaxV = 0.0;
            int Idx1 = 0;
            int Idx2 = -1;
            int MaxIdx = 0;
            double tmpMaxV = 0.0;
            for (double pos = StartCDF; pos < EndCDF; pos += 0.02)
            {
                Idx1 = (int)(Cnt * pos) + Start;
                if (Idx2 >= 0)
                {
                    tmpMaxV = Arrdata[Idx1] - Arrdata[Idx2];
                    if (tmpMaxV > DeltaMaxV)
                    {
                        DeltaMaxV = tmpMaxV;
                        MaxIdx = Idx1;
                    }
                }
                Idx2 = Idx1;
            }
            return MaxIdx;
        }
    }
}