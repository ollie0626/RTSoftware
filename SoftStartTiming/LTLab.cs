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

namespace SoftStartTiming
{
    public partial class LTLab : Form
    {

        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();
        string win_name = "LTLab v1.0";
        RTBBControl RTDev = new RTBBControl();

        public LTLab()
        {
            InitializeComponent();
            this.Text = win_name;
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

            //if (Device_map.ContainsKey("63600-2"))
            //{
            //    await ConnectTask(Device_map["63600-2"], 2);
            //    tb_eload.Text = "ELoad:63600-2";
            //}

            //if (Device_map.ContainsKey("34970A"))
            //{
            //    await ConnectTask(Device_map["34970A"], 3);
            //    tb_daq.Text = "DAQ:34970A";
            //}

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
            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                test_parameter.lt_lab.addr_list.Add(Convert.ToByte(dataGridView1[0, i].Value.ToString(), 16));
                test_parameter.lt_lab.data_list.Add(Convert.ToByte(dataGridView1[1, i].Value.ToString(), 16));
            }
        }

        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            
            try
            {
                test_parameter_copy();
            }
            catch
            {

            }
        }
    }
    public class LTLab_parameter
    {
        public List<byte> addr_list = new List<byte>();
        public List<byte> data_list = new List<byte>();
    }

}
