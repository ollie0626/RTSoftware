using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using InsLibDotNet;
using System.Windows.Forms;
using System.Threading;

namespace OLEDLite
{
    public class Chamber_Thread
    {
        //public static Label label1 = new Label();
        //public static ProgressBar progressBar1 = new ProgressBar();
        public static string res_chamber;
        public static int steadyTime;
        

        private static bool RecountTime()
        {
            steadyTime--; System.Threading.Thread.Sleep(1000);
            return true;
        }

        private static Task<bool> TaskRecount()
        {
            return Task.Factory.StartNew(() => RecountTime());
        }

        public static async void Chamber_Task(object obj)
        {
            for (int i = 0; i < test_parameter.tempList.Count; i++)
            {
                if (!Directory.Exists(test_parameter.wave_path + @"\" + test_parameter.tempList[i] + "C"))
                {
                    Directory.CreateDirectory(test_parameter.wave_path + @"\" + test_parameter.tempList[i] + "C");
                }
                test_parameter.wave_path = test_parameter.wave_path + @"\" + test_parameter.tempList[i] + "C";

                
                // chamber control
                InsControl._chamber = new ChamberModule(res_chamber);
                InsControl._chamber.ConnectChamber(res_chamber);
                InsControl._chamber.ChamberOn(test_parameter.tempList[i]);
                InsControl._chamber.ChamberOn(test_parameter.tempList[i]);
                await InsControl._chamber.ChamberStable(test_parameter.tempList[i]);
                steadyTime = test_parameter.steadyTime;

                for (; steadyTime > 0;)
                {
                    await TaskRecount();
                    //((ProgressBar)progressBar1).Value = test_parameter.steadyTime;
                    //((Label)label1).Invoke((MethodInvoker)(() => ((Label)label1).Text = "count down: " + (test_parameter.steadyTime / 60).ToString() + ":" + (test_parameter.steadyTime % 60).ToString()));
                    //label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString();
                }

                //ate_table[(int)idx].temp = test_parameter.tempList[i];
            }
            if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
        }



    }
}
