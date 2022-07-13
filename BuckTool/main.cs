using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Sunny.UI;
using InsLibDotNet;
using System.Threading;
using System.Diagnostics;


//using RTBBLibDotNet;


namespace BuckTool
{
    public interface ITask
    {
        void ATE_Task();
    }


    public partial class main : Sunny.UI.UIForm
    {
        public main()
        {
            InitializeComponent();
            //RTBBControl.BoardInit();
            //RTBBControl.GpioInit();
        }

        private void uibt_osc_connect_Click(object sender, EventArgs e)
        {
            UIButton bt = (UIButton)sender;
            int idx = bt.TabIndex;

            switch (idx)
            {
                case 0:
                    InsControl._scope = new AgilentOSC(tb_osc.Text);
                    if (InsControl._scope.InsState())
                        led_osc.Color = Color.LightGreen;
                    else
                        led_osc.Color = Color.Red;
                    break;
                case 1:
                    InsControl._power = new PowerModule((int)nu_power.Value);
                    if (InsControl._power.InsState())
                        led_power.Color = Color.LightGreen;
                    else
                        led_power.Color = Color.Red;
                    break;
                case 2:
                    InsControl._eload = new EloadModule((int)nu_eload.Value);
                    if (InsControl._eload.InsState())
                        led_eload.Color = Color.LightGreen;
                    else
                        led_eload.Color = Color.Red;
                    break;
                case 3:
                    InsControl._34970A = new MultiChannelModule((int)nu_34970A.Value);
                    if (InsControl._34970A.InsState())
                        led_37940.Color = Color.LightGreen;
                    else
                        led_37940.Color = Color.Red;
                    break;
                case 4:
                    InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                    if (InsControl._chamber.InsState())
                        led_chamber.Color = Color.LightGreen;
                    else
                        led_chamber.Color = Color.Red;
                    break;
            }
        }

        private void uibt_kill_Click(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
        }

        private void bt_load_add_Click(object sender, EventArgs e)
        {
            
        }
    }
}
