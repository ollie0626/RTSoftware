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
using RTBBLibDotNet;
using System.Text.RegularExpressions;
using System.Threading;
using System.IO;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace SoftStartTiming
{
    public partial class CrossTalk : Form
    {
        private string win_name = "Cross talk v1.0";
        ParameterizedThreadStart p_thread;
        ATE_CrossTalk _ate_crosstalk = new ATE_CrossTalk();
        Thread ATETask;
        int SteadyTime;
        string[] tempList;
        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();

        TaskRun[] ate_table;

        public CrossTalk()
        {
            InitializeComponent();
        }

        private void DataGridviewInit()
        {
            EloadDG_CCM.RowCount = 4;
            FreqDG.RowCount = 4;
            VoutDG.RowCount = 4;

            EloadDG_CCM[0, 0].Value = "Buck1";
            EloadDG_CCM[0, 1].Value = "Buck2";
            EloadDG_CCM[0, 2].Value = "Buck3";
            EloadDG_CCM[0, 3].Value = "Buck4";
            // address
            EloadDG_CCM[1, 0].Value = "01";
            EloadDG_CCM[1, 1].Value = "01";
            EloadDG_CCM[1, 2].Value = "01";
            EloadDG_CCM[1, 3].Value = "01";

            // channel enable
            EloadDG_CCM[2, 0].Value = "01";
            EloadDG_CCM[2, 1].Value = "02";
            EloadDG_CCM[2, 2].Value = "03";
            EloadDG_CCM[2, 3].Value = "04";

            // channel disable
            EloadDG_CCM[3, 0].Value = "00";
            EloadDG_CCM[3, 1].Value = "00";
            EloadDG_CCM[3, 2].Value = "00";
            EloadDG_CCM[3, 3].Value = "00";

            // victim loading
            EloadDG_CCM[4, 0].Value = "0.1,0.6,0.7";
            EloadDG_CCM[4, 1].Value = "0.8,1,2";
            EloadDG_CCM[4, 2].Value = "0.35,0.5,0.7";
            EloadDG_CCM[4, 3].Value = "0.9,1,1.2";

            // aggresor loading
            EloadDG_CCM[4, 0].Value = "1";
            EloadDG_CCM[4, 1].Value = "2";
            EloadDG_CCM[4, 2].Value = "3";
            EloadDG_CCM[4, 3].Value = "4";

            // full load setting
            EloadDG_CCM[5, 0].Value = "0.8";
            EloadDG_CCM[5, 1].Value = "0.9";
            EloadDG_CCM[5, 2].Value = "1";
            EloadDG_CCM[5, 3].Value = "1.2";

            FreqDG[1, 0].Value = "10";
            FreqDG[1, 1].Value = "20";
            FreqDG[1, 2].Value = "30";
            FreqDG[1, 3].Value = "40";

            FreqDG[2, 0].Value = "12";
            FreqDG[2, 1].Value = "22";
            FreqDG[2, 2].Value = "32";
            FreqDG[2, 3].Value = "10";

            FreqDG[3, 0].Value = "600";
            FreqDG[3, 1].Value = "800";
            FreqDG[3, 2].Value = "600";
            FreqDG[3, 3].Value = "800";

            VoutDG[1, 0].Value = "11";
            VoutDG[1, 1].Value = "22";
            VoutDG[1, 2].Value = "33";
            VoutDG[1, 3].Value = "44";

            VoutDG[2, 0].Value = "10";
            VoutDG[2, 1].Value = "20";
            VoutDG[2, 2].Value = "30";
            VoutDG[2, 3].Value = "40";

            VoutDG[3, 0].Value = "3.3";
            VoutDG[3, 1].Value = "2";
            VoutDG[3, 2].Value = "2";
            VoutDG[3, 3].Value = "3.3";
        }

        private void CrossTalk_Load(object sender, EventArgs e)
        {
            this.Text = win_name;
            DataGridviewInit();
            ate_table = new TaskRun[] { _ate_crosstalk };
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

        private void test_parameter_copy()
        {
            test_parameter.freq_data.Clear();
            test_parameter.vout_data.Clear();
            test_parameter.freq_des.Clear();
            test_parameter.vout_des.Clear();
            test_parameter.ccm_eload.Clear();

            test_parameter.vin_conditions = "Vin :" + tb_vinList.Text + " (V)\r\n";
            test_parameter.tool_ver = win_name + "\r\n";

            test_parameter.chamber_en = ck_chamber_en.Checked;
            test_parameter.run_stop = false;
            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.slave = (byte)nuslave.Value;

            // perpare test parameter.
            for (int i = 0; i < test_parameter.cross_en.Length; i++)
            {
                // Eload CCM data grid
                test_parameter.en_addr[i] = Convert.ToByte(Convert.ToString(EloadDG_CCM[1, i].Value), 16);
                test_parameter.en_data[i] = Convert.ToByte(Convert.ToString(EloadDG_CCM[2, i].Value), 16);
                test_parameter.disen_data[i] = Convert.ToByte(Convert.ToString(EloadDG_CCM[3, i].Value), 16);
                test_parameter.ccm_eload.Add(i, ((string)EloadDG_CCM[4, i].Value).Split(',').Select(double.Parse).ToList());
                test_parameter.full_load[i] = Convert.ToDouble(EloadDG_CCM[5, i].Value);

                // freq and vout parameter
                test_parameter.freq_addr[i] = Convert.ToByte(Convert.ToString(FreqDG[1, i].Value), 16);
                test_parameter.vout_addr[i] = Convert.ToByte(Convert.ToString(VoutDG[1, i].Value), 16);
                test_parameter.freq_data.Add(i, ((string)FreqDG[2, i].Value).Split(',').Select(byte.Parse).ToList());
                test_parameter.vout_data.Add(i, ((string)VoutDG[2, i].Value).Split(',').Select(byte.Parse).ToList());
                test_parameter.freq_des.Add(i, ((string)FreqDG[3, i].Value).Split(',').ToList());
                test_parameter.vout_des.Add(i, ((string)VoutDG[3, i].Value).Split(',').ToList());
                test_parameter.cross_en[i] = true;
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
            catch (Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
            }
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
                    //if (!Directory.Exists(tbWave.Text + tempList[i] + "C"))
                    //{
                    //    Directory.CreateDirectory(tbWave.Text + tempList[i] + "C");
                    //}
                    //test_parameter.waveform_path = tbWave.Text + tempList[i] + "C";

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

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
        }

        private void BTSelectWavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbWave.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTStop_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
        }

        private void nuCH_number_ValueChanged(object sender, EventArgs e)
        {
            EloadDG_CCM.RowCount = (int)nuCH_number.Value;
            FreqDG.RowCount = (int)nuCH_number.Value;
            VoutDG.RowCount = (int)nuCH_number.Value;

            test_parameter.freq_addr = new byte[(int)nuCH_number.Value];
            test_parameter.vout_addr = new byte[(int)nuCH_number.Value];
            test_parameter.en_addr = new byte[(int)nuCH_number.Value];
            test_parameter.en_data = new byte[(int)nuCH_number.Value];
            test_parameter.disen_data = new byte[(int)nuCH_number.Value];
            test_parameter.full_load = new double[(int)nuCH_number.Value];

            test_parameter.cross_select = new byte[(int)nuCH_number.Value];
            test_parameter.cross_en = new bool[(int)nuCH_number.Value];

            test_parameter.ch_num = (int)nuCH_number.Value;
            test_parameter.full_load = new double[(int)nuCH_number.Value];
        }

        private void EloadDG_CCM_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            test_parameter.rail_name = new string[(int)nuCH_number.Value];
            for (int i = 0; i < (int)nuCH_number.Value; i++)
            {
                if (EloadDG_CCM.RowCount == 0) break;
                test_parameter.rail_name[i] = (string)EloadDG_CCM[0, i].Value;
                FreqDG[0, i].Value = test_parameter.rail_name[i];
                VoutDG[0, i].Value = test_parameter.rail_name[i];
            }
        }
    }
}
