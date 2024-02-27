using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Threading;

namespace Scope_Simple_tool.MyPage
{
    public class MyControl
    {
        static public BackgroundWorker BW = null;
        static public ManualResetEvent ManualReset = new ManualResetEvent(false);
        static public bool IsStop = false;
        static public bool IsPause = false;

        static public void SendPercentage(int n, string Title)
        {
            if (MyControl.BW != null)
                MyControl.BW.ReportProgress(n, Title);
        }

        static public void Sleep100Ms(int n, bool IsDelay = true)
        {
            for (int i = 0; i < n; ++i)
            {
                if (IsStop) break;
                //ManualReset.WaitOne();
                if (!VisaCommand._IsDebug && IsDelay) Thread.Sleep(100);
            }
            //ManualReset.WaitOne();
        }

        static public void SleepMs(int n, bool IsDelay = true)
        {
            int t1 = n / 100;
            if (t1 > 0)
            {
                Sleep100Ms(t1);
            }
            if (!VisaCommand._IsDebug && IsDelay) Thread.Sleep(n % 100);
            //ManualReset.WaitOne();
        }



    }
}
