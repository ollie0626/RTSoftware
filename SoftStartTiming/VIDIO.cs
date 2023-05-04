using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RTBBLibDotNet;
using InsLibDotNet;
using System.Threading;
using System.IO;
using System.Diagnostics;

using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;

namespace SoftStartTiming
{
    public partial class VIDIO : Form
    {
        string win_name = "VIDIO v1.0";
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        TaskRun[] ate_table;
        string[] tempList;
        int SteadyTime;

        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();


        ATE_VIDIO _ate_vid_io = new ATE_VIDIO();


        private void InitDG()
        {
            // lpm mode
            Column1.Items.Add("0");
            Column1.Items.Add("1");

            // g1 logic
            Column2.Items.Add("0");
            Column2.Items.Add("1");

            // g2 logic
            Column3.Items.Add("0");
            Column3.Items.Add("1");

            // after
            Column5.Items.Add("0");
            Column5.Items.Add("1");

            Column6.Items.Add("0");
            Column6.Items.Add("1");

            Column7.Items.Add("0");
            Column7.Items.Add("1");
        }

        public VIDIO()
        {
            InitializeComponent();
            InitDG();
            this.Name = win_name;

            ate_table = new TaskRun[] { _ate_vid_io };
            //RTBBControl dev = new RTBBControl();
            //dev.BoadInit();
            //dev.GPIOnState(0, true);
            //dev.GPIOnState(0, false);
        }

        private void BT_Add_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = dataGridView1.RowCount + 1;
        }

        private void BT_Sub_Click(object sender, EventArgs e)
        {
            if(dataGridView1.RowCount - 1 > 0)
                dataGridView1.RowCount = dataGridView1.RowCount - 1;
            else
                dataGridView1.RowCount = 0;
        }

        private void test_parameter_copy()
        {
            test_parameter.vidio.lpm_sel.Clear();
            test_parameter.vidio.g1_sel.Clear();
            test_parameter.vidio.g2_sel.Clear();
            test_parameter.vidio.vout_list.Clear();

            test_parameter.vidio.lpm_sel_af.Clear();
            test_parameter.vidio.g1_sel_af.Clear();
            test_parameter.vidio.g2_sel_af.Clear();
            test_parameter.vidio.vout_list_af.Clear();

            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.IoutList = tb_iout.Text.Split(',').Select(double.Parse).ToList();

            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                test_parameter.vidio.lpm_sel.Add(Convert.ToInt16(dataGridView1[0, i].Value));
                test_parameter.vidio.g1_sel.Add(Convert.ToInt16(dataGridView1[1, i].Value));
                test_parameter.vidio.g2_sel.Add(Convert.ToInt16(dataGridView1[2, i].Value));
                test_parameter.vidio.vout_list.Add(Convert.ToDouble(dataGridView1[3, i].Value));

                test_parameter.vidio.lpm_sel_af.Add(Convert.ToInt16(dataGridView1[4, i].Value));
                test_parameter.vidio.g1_sel_af.Add(Convert.ToInt16(dataGridView1[5, i].Value));
                test_parameter.vidio.g2_sel_af.Add(Convert.ToInt16(dataGridView1[6, i].Value));
                test_parameter.vidio.vout_list_af.Add(Convert.ToDouble(dataGridView1[7, i].Value));
            }
        }

        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            try
            {
                test_parameter_copy();
                if (ck_chamber_en.Checked)
                {
                    tempList = tb_templist.Text.Split(',');
                    p_thread = new ParameterizedThreadStart(Chamber_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(0);
                }
                else
                {
                    p_thread = new ParameterizedThreadStart(Run_Single_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(0);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
                BTRun.Enabled = true;
            }
        }

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
        }

        private bool RecountTime()
        {
            SteadyTime--; System.Threading.Thread.Sleep(1000);
            return true;
        }

        private Task<bool> TaskRecount()
        {
            return Task.Factory.StartNew(() => RecountTime());
        }

        public async void Chamber_Task(object idx)
        {
            try
            {
                for (int i = 0; i < tempList.Length; i++)
                {
                    if (!Directory.Exists(tbWave.Text + tempList[i] + "C"))
                    {
                        Directory.CreateDirectory(tbWave.Text + tempList[i] + "C");
                    }
                    test_parameter.waveform_path = tbWave.Text + tempList[i] + "C";

                    SteadyTime = (int)nu_steady.Value;
                    //InsControl._chamber = new ChamberModule(tb_chamber.Text);
                    //InsControl._chamber.ConnectChamber(tb_chamber.Text);
                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    //await InsControl._chamber.ChamberStable(Convert.ToDouble(tempList[i]));
                    for (; SteadyTime > 0;)
                    {
                        await TaskRecount();
                        progressBar1.Value = SteadyTime;
                        label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
                    }
                    ate_table[(int)idx].temp = Convert.ToDouble(tempList[i]);
                    ate_table[(int)idx].ATETask();

                }
                if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, win_name, System.Windows.Forms.MessageBoxButtons.OK);
            }
            finally
            {
                if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
            }
            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
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
                case 3: InsControl._34970A = new MultiChannelModule(res); break;
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

            if (Device_map.ContainsKey("34970A"))
            {
                await ConnectTask(Device_map["34970A"], 3);
                tb_daq.Text = "DAQ:34970A";
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

            if (InsControl._34970A != null)
            {
                if (InsControl._34970A.InsState())
                    led_daq.BackColor = Color.LightGreen;
                else
                    led_daq.BackColor = Color.Red;
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

        private void BTStop_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = true;
        }
    }


    public class VIDIO_parameter
    {
        public List<int> lpm_sel = new List<int>();
        public List<int> g1_sel = new List<int>();
        public List<int> g2_sel = new List<int>();
        public List<double> vout_list = new List<double>();

        public List<int> lpm_sel_af = new List<int>();
        public List<int> g1_sel_af = new List<int>();
        public List<int> g2_sel_af = new List<int>();
        public List<double> vout_list_af = new List<double>();
    }


}
