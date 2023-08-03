using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using InsLibDotNet;

using System.Text.RegularExpressions;
using System.Threading;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoftStartTiming
{
    public partial class LTLab : Form
    {

        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();
        string win_name = "LTLab v1.0";
        ParameterizedThreadStart p_thread;
        ATE_LTLab _ate_ltlab;
        Thread ATETask;
        RTBBControl RTDev = new RTBBControl();
        TaskRun[] ate_table;

        public LTLab()
        {
            InitializeComponent();
            this.Text = win_name;
            _ate_ltlab = new ATE_LTLab();
            ate_table = new TaskRun[] { _ate_ltlab };
        }

        private void bt_up_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = dataGridView1.RowCount + 1;
        }

        private void bt_down_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 1)
                dataGridView1.RowCount = dataGridView1.RowCount - 1;
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

        private int ConnectFunc(string res, int ins_sel)
        {
            switch (ins_sel)
            {
                case 0:
                    InsControl._oscilloscope = new OscilloscopesModule(res);
                    break;
                case 1: InsControl._power = new PowerModule(res); break;
                //case 2: InsControl._eload = new EloadModule(res); break;
                //case 3: InsControl._34970A = new MultiChannelModule(res); break;
                //case 4: InsControl._chamber = new ChamberModule(res); break;
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
        }

        private void LTLab_Load(object sender, EventArgs e)
        {
            RTDev.BoadInit();
            List<byte> list = RTDev.ScanSlaveID();

            if (list != null)
            {
                if (list.Count > 0)
                    nuslave.Value = list[0];
            }
        }

        private void test_parameter_copy()
        {
            test_parameter.slave = (byte)nuslave.Value;
            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                test_parameter.lt_lab.addr_list.Add(Convert.ToByte(dataGridView1[0, i].Value.ToString(), 16));
                test_parameter.lt_lab.data_list.Add(Convert.ToByte(dataGridView1[1, i].Value.ToString(), 16));
                test_parameter.lt_lab.vout_list.Add(Convert.ToDouble(dataGridView1[2, i].Value.ToString()));
            }
            
            //test_parameter.lt_lab.time_scale = (double)nuTimeScale.Value;
            //int start = Convert.ToInt32(dataGridView1[1, 0].Value.ToString(), 16);
            //int stop = Convert.ToInt32(dataGridView1[2, 0].Value.ToString(), 16);
            //int step = Convert.ToInt32(dataGridView1[3, 0].Value.ToString(), 16);
            //int res = 0;
            //for(int i = 0; res < stop; i++)
            //{
            //    res = start + i * step;
            //    test_parameter.lt_lab.data_list.Add(Convert.ToByte(res));
            //}
        }

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
        }

        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            try
            {
                RTDev.BoadInit();
                List<byte> list = RTDev.ScanSlaveID();

                if (list != null)
                {
                    if (list.Count > 0)
                        nuslave.Value = list[0];
                }

                test_parameter_copy();

                p_thread = new ParameterizedThreadStart(Run_Single_Task);
                ATETask = new Thread(p_thread);
                ATETask.Start(0);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
                BTRun.Enabled = true;
            }
        }

        private void BTPause_Click(object sender, EventArgs e)
        {
            if (ATETask == null) return;
            System.Threading.ThreadState state = ATETask.ThreadState;
            if (state == System.Threading.ThreadState.Running || state == System.Threading.ThreadState.WaitSleepJoin)
            {
                ATETask.Suspend();
            }
            else if (state == System.Threading.ThreadState.Suspended)
            {
                ATETask.Resume();
            }
        }

        private void BTStop_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = true;
            if (ATETask != null)
            {
                if (ATETask.IsAlive)
                {
                    System.Threading.ThreadState state = ATETask.ThreadState;
                    if (state == System.Threading.ThreadState.Suspended) ATETask.Resume();
                    ATETask.Abort();
                    MessageBox.Show("ATE Task Stop !!", "ATE Tool", MessageBoxButtons.OK);
                    //InsControl._power.AutoPowerOff();
                }
            }
        }
    }

    public class LTLab_parameter
    {
        public List<byte> addr_list = new List<byte>();
        public List<byte> data_list = new List<byte>();
        public List<double> vout_list = new List<double>();
        public double time_scale;
    }

}
