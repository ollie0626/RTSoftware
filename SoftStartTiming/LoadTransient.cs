using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using RTBBLibDotNet;
using System.Threading;
using InsLibDotNet;

namespace SoftStartTiming
{
    public partial class LoadTransient : Form
    {
        string win_name = "Load Transient v1.0";
        RTBBControl RTDev = new RTBBControl();
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        TaskRun[] ate_table;
        string[] tempList;
        int SteadyTime;
        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();


        public LoadTransient()
        {
            InitializeComponent();
            this.Text = win_name;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void BTPause_Click(object sender, EventArgs e)
        {

        }

        private void BTStop_Click(object sender, EventArgs e)
        {

        }

        private void BTRun_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
        }

        private int ConnectFunc(string res, int ins_sel)
        {
            switch (ins_sel)
            {
                case 0:
                    InsControl._oscilloscope = new OscilloscopesModule(res);
                    break;
                case 1: InsControl._power = new PowerModule(res); break;
                case 2: InsControl._eload = new EloadModule(res); break;
                case 3: InsControl._funcgen = new FuncGenModule(res); break;
                case 4: InsControl._chamber = new ChamberModule(res); break;
            }
            return 0;
        }

        private Task<int> ConnectTask(string res, int ins_sel)
        {
            return Task.Factory.StartNew(() => ConnectFunc(res, ins_sel));
        }

        private async void uibt_osc_connect_Click(object sender, EventArgs e)
        {
            BTScan_Click(null, null);

            Button bt = (Button)sender;
            int idx = bt.TabIndex;
            string[] scope_name = new string[] { "DSOS054A", "DSO9064A", "DPO7054C", "DPO7104C" };
            // scope idn name keysight DSOS054A DSO9064A  Tek DPO7054C DSO9064A

            for (int i = 0; i < scope_name.Length; i++)
            {
                if (Device_map.ContainsKey(scope_name[i]))
                {
                    await ConnectTask(Device_map[scope_name[i]], 0);
                    tb_osc.Text = "Scope:" + scope_name[i];
                }
            }

            if (Device_map.ContainsKey("CBPower.Text"))
            {
                await ConnectTask(Device_map[CBPower.Text], 1);
            }

            if (Device_map.ContainsKey("63600-2"))
            {
                await ConnectTask(Device_map["63600-2"], 2);
                tb_eload.Text = "ELoad:63600-2";
            }

            if (Device_map.ContainsKey("AFG3"))
            {
                await ConnectTask(Device_map["AFG3"], 3);
                tb_funcgen.Text = "Funcgen:AFG3xxx";
            }

            await ConnectTask("GPIB0::3::INSTR", 4);
            MyLib.Delay1s(1);
            check_ins_state();
        }

        private void check_ins_state()
        {
            if (InsControl._oscilloscope != null)
            {
                if (InsControl._oscilloscope.InsState())
                    led_osc.BackColor = Color.LightGreen;
                else
                    led_osc.BackColor = Color.Red;
            }

            if (InsControl._power != null)
            {
                if (InsControl._power.InsState())
                    led_power.BackColor = Color.LightGreen;
                else
                    led_power.BackColor = Color.Red;
            }

            if (InsControl._eload != null)
            {
                if (InsControl._eload.InsState())
                    led_eload.BackColor = Color.LightGreen;
                else
                    led_eload.BackColor = Color.Red;
            }

            if (InsControl._funcgen != null)
            {
                if (InsControl._funcgen.InsState())
                    led_funcgen.BackColor = Color.LightGreen;
                else
                    led_funcgen.BackColor = Color.Red;
            }

            if (InsControl._chamber != null)
            {
                if (InsControl._chamber.InsState())
                    led_chamber.BackColor = Color.LightGreen;
                else
                    led_chamber.BackColor = Color.Red;
            }

        }

        private void BTScan_Click(object sender, EventArgs e)
        {
            list_ins.Items.Clear();
            string[] scope_name = new string[] { "DSOS054A", "DSO9064A", "DPO7054C", "DPO7104C" };
            string[] ins_list = ViCMD.ScanIns();
            if (ins_list == null) return;
            foreach (string ins in ins_list)
            {
                list_ins.Items.Add(ins);

                VisaCommand visaCommand = new VisaCommand();
                visaCommand.LinkingIns(ins);
                string idn = visaCommand.doQueryIDN();
                string name = "";

                if (idn.Split(',').Length != 1)
                    name = idn.Split(',')[1] != null ? idn.Split(',')[1] : "";

                if (idn.Split(',').Length != 1)
                {
                    if (idn.Split(',')[0] == "TEKTRONIX")
                    {
                        for (int i = 0; i < scope_name.Length; i++)
                        {
                            if (name == scope_name[i])
                                InsControl._tek_scope_en = true;
                        }
                    }
                }

                if (Device_map.ContainsKey(name) == false)
                {
                    Device_map.Add(name, ins);
                    if (name.IndexOf("E363") != -1)
                    {
                        CBPower.Enabled = true;
                        CBPower.Items.Add(name);
                    }
                }
            }
        }

        private void LoadTransient_Load(object sender, EventArgs e)
        {
            RTDev.BoadInit();
            List<byte> list = RTDev.ScanSlaveID();

            if (list != null)
            {
                if (list.Count > 0)
                    nuslave.Value = list[0];
            }
        }
    }

    public class LoadTransient_parameter
    {
        // function gen parameter
        public double freq;
        public double duty;

        // eload parameter
        public double T1;
        public double T2;

        public double Tr;
        public double Tf;
        public double zoom_in_ratio;

        // dynamic loading setting
        public List<double> hi_current = new List<double>();
        public List<double> lo_current = new List<double>();

        // for evb current offset
        public double offset;

        public bool eload_dev_sel;
        public double gain;
    }
}
