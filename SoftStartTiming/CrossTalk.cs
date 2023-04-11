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
        private string win_name = "Cross talk v1.2";
        ParameterizedThreadStart p_thread;
        ATE_CrossTalk _ate_crosstalk;
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
            LTDG.RowCount = 4;
            nuCH_number.Value = 4;
            MeasDG.RowCount = 4;


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
            EloadDG_CCM[4, 0].Value = "0.01";
            EloadDG_CCM[4, 1].Value = "0.02";
            EloadDG_CCM[4, 2].Value = "0.03";
            EloadDG_CCM[4, 3].Value = "0.04";

            // full load setting
            EloadDG_CCM[5, 0].Value = "0.03";
            EloadDG_CCM[5, 1].Value = "0.04";
            EloadDG_CCM[5, 2].Value = "0.05";
            EloadDG_CCM[5, 3].Value = "0.06";

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

            VoutDG[3, 0].Value = "1.2";
            VoutDG[3, 1].Value = "4.6";
            VoutDG[3, 2].Value = "7";
            VoutDG[3, 3].Value = "3.3";

            LTDG[1, 0].Value = "0";
            LTDG[1, 1].Value = "0";
            LTDG[1, 2].Value = "0";
            LTDG[1, 3].Value = "0";

            LTDG[2, 0].Value = "0.01";
            LTDG[2, 1].Value = "0.02";
            LTDG[2, 2].Value = "0.03";
            LTDG[2, 3].Value = "0.04";

            LTDG[3, 0].Value = "0.04";
            LTDG[3, 1].Value = "0.05";
            LTDG[3, 2].Value = "0.06";
            LTDG[3, 3].Value = "0.07";


            for (int i = 1; i < 5; i++)
            {
                ScopeCH.Items.Add("CH" + i);
                ELoadCH.Items.Add("CH" + i);
                LxCH.Items.Add("CH" + i);
            }
            LxCH.Items.Add("Non-use");
            
        }


        private void CrossTalk_Load(object sender, EventArgs e)
        {
            this.Text = win_name;
            DataGridviewInit();

            _ate_crosstalk = new ATE_CrossTalk(this);
            ate_table = new TaskRun[] { _ate_crosstalk };

            CBItem.SelectedIndex = 0;

            RTBBControl RTDev = new RTBBControl();
            RTDev.BoadInit();
            List<byte> list = RTDev.ScanSlaveID();

            if(list != null)
            {
                if (list.Count > 0)
                    nuslave.Value = list[0];
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
            nuCH_number_ValueChanged(null, null);
            test_parameter.freq_data.Clear();
            test_parameter.vout_data.Clear();
            test_parameter.freq_des.Clear();
            test_parameter.vout_des.Clear();
            test_parameter.ccm_eload.Clear();
            test_parameter.scope_chx.Clear();
            test_parameter.eload_chx.Clear();
            test_parameter.scope_lx.Clear();

            test_parameter.lt_l1.Clear();
            test_parameter.lt_l2.Clear();

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
                //test_parameter.lt_full[i] = Convert.ToDouble(LTDG[3, i].Value);

                // get scope channel number (data type string -> "CHn")
                DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)MeasDG[1, i];
                string txt = (string)comboBoxCell.Value;
                test_parameter.scope_chx.Add(txt);

                // get eload channel number (data type int)
                comboBoxCell = (DataGridViewComboBoxCell)MeasDG[2, i];
                txt = (string)comboBoxCell.Value;
                test_parameter.eload_chx.Add(Convert.ToInt32(txt.Replace("CH", "")));

                // data type string -> "CHn"
                comboBoxCell = (DataGridViewComboBoxCell)MeasDG[3, i];
                txt = (string)comboBoxCell.Value;
                test_parameter.scope_lx.Add(txt);

                // freq and vout parameter
                test_parameter.freq_addr[i] = Convert.ToByte(Convert.ToString(FreqDG[1, i].Value), 16);
                test_parameter.vout_addr[i] = Convert.ToByte(Convert.ToString(VoutDG[1, i].Value), 16);

                string[] tmp = ((string)FreqDG[2, i].Value).Split(',');
                byte[] byt_tmp = new byte[tmp.Length];
                for(int idx = 0; idx < tmp.Length; idx++)
                {
                    byt_tmp[idx] = Convert.ToByte(tmp[idx], 16);
                }
                test_parameter.freq_data.Add(i, byt_tmp.ToList());

                tmp = ((string)VoutDG[2, i].Value).Split(',');
                byt_tmp = new byte[tmp.Length];
                for (int idx = 0; idx < tmp.Length; idx++)
                {
                    byt_tmp[idx] = Convert.ToByte(tmp[idx], 16);
                }
                test_parameter.vout_data.Add(i, byt_tmp.ToList());

                test_parameter.freq_des.Add(i, ((string)FreqDG[3, i].Value).Split(',').ToList());
                test_parameter.vout_des.Add(i, ((string)VoutDG[3, i].Value).Split(',').ToList());

                switch (CBItem.SelectedIndex)
                {
                    case 2:
                        test_parameter.vid_addr[i] = Convert.ToByte(Convert.ToString(LTDG[1, i].Value), 16);
                        test_parameter.hi_code[i] = Convert.ToByte(Convert.ToString(LTDG[2, i].Value), 16);
                        test_parameter.lo_code[i] = Convert.ToByte(Convert.ToString(LTDG[3, i].Value), 16);
                        break;
                    case 3:
                        test_parameter.lt_l1.Add(i, ((string)LTDG[1, i].Value).Split(',').Select(double.Parse).ToList());
                        test_parameter.lt_l2.Add(i, ((string)LTDG[2, i].Value).Split(',').Select(double.Parse).ToList());
                        break;
                }

                test_parameter.cross_en[i] = true;
            }

            test_parameter.cross_mode = CBItem.SelectedIndex;
            test_parameter.offtime_scale_ms = (double)(nu_ontime_scale.Value / 1000);
            test_parameter.waveform_path = tbWave.Text;
            test_parameter.tolerance = (double)nuToerance.Value / 100;

            int vout_cnt = test_parameter.vout_data[0].Count;
            int freq_cnt = test_parameter.freq_data[0].Count;
            int iout_cnt = test_parameter.ccm_eload[0].Count;

            for(int i = 1; i < nuCH_number.Value; i++)
            {
                if (vout_cnt < test_parameter.vout_data[i].Count)
                    vout_cnt = test_parameter.vout_data[i].Count;

                if (freq_cnt < test_parameter.freq_data[i].Count)
                    freq_cnt = test_parameter.freq_data[i].Count;

                if (iout_cnt < test_parameter.ccm_eload[i].Count)
                    iout_cnt = test_parameter.ccm_eload[i].Count;
            }

            int n = ((int)nuCH_number.Value - 1);
            int progress_max = (int)nuCH_number.Value
                                * test_parameter.VinList.Count
                                * vout_cnt
                                * freq_cnt
                                * iout_cnt
                                * 2 // no load and full load
                                * (int)Math.Pow(2, n)
                                ;

            progressBar2.Maximum = progress_max;
            //Console.WriteLine("progress max = " + progress_max);
        }

        public void UpdateProgressBar(int val)
        {
            progressBar2.Invoke((MethodInvoker)(() => progressBar2.Value = val));
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
                BTRun.Enabled = true;
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
            LTDG.RowCount = (int)nuCH_number.Value;
            MeasDG.RowCount = (int)nuCH_number.Value;

            // others conditions
            test_parameter.freq_addr = new byte[(int)nuCH_number.Value];
            test_parameter.vout_addr = new byte[(int)nuCH_number.Value];
            test_parameter.full_load = new double[(int)nuCH_number.Value];

            test_parameter.cross_select = new byte[(int)nuCH_number.Value];
            test_parameter.cross_en = new bool[(int)nuCH_number.Value];
            test_parameter.ch_num = (int)nuCH_number.Value;
            test_parameter.full_load = new double[(int)nuCH_number.Value];

            // vid
            test_parameter.vid_addr = new byte[(int)nuCH_number.Value];
            test_parameter.lo_code = new byte[(int)nuCH_number.Value];
            test_parameter.hi_code = new byte[(int)nuCH_number.Value];

            // en on / off
            test_parameter.en_addr = new byte[(int)nuCH_number.Value];
            test_parameter.en_data = new byte[(int)nuCH_number.Value];
            test_parameter.disen_data = new byte[(int)nuCH_number.Value];

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
                LTDG[0, i].Value = test_parameter.rail_name[i];
                MeasDG[0, i].Value = test_parameter.rail_name[i];

                MeasDG[1, i].Value = "CH" + (i + 1);
                MeasDG[2, i].Value = "CH" + (i + 1);
                MeasDG[3, i].Value = "Non-use";

            }
        }

        private void CBPower_SelectedIndexChanged(object sender, EventArgs e)
        {
            Console.WriteLine(Device_map[CBPower.Text]);

            InsControl._power = new PowerModule(Device_map[CBPower.Text]);

            tb_power.Text = "Power: " + CBPower.Text;
            if (InsControl._power.InsState())
                led_power.BackColor = Color.LightGreen;
            else
                led_power.BackColor = Color.Red;

            CBChannel.Items.Clear();
            CBChannel.Enabled = true;
            switch (CBPower.Text)
            {
                case "E3631A":
                    CBChannel.Items.Add("+6V");
                    CBChannel.Items.Add("+25V");
                    CBChannel.Items.Add("-25V");
                    CBChannel.SelectedIndex = 0;
                    break;
                case "E3632A":
                    CBChannel.Items.Add("15V");
                    CBChannel.Items.Add("30V");
                    CBChannel.SelectedIndex = 0;
                    break;
                case "E3633A":
                    CBChannel.Items.Add("8V");
                    CBChannel.Items.Add("20V");
                    CBChannel.SelectedIndex = 0;
                    break;
                case "E3634A":
                    CBChannel.Items.Add("25V");
                    CBChannel.Items.Add("50V");
                    CBChannel.SelectedIndex = 0;
                    break;
                case "62006P":
                    CBChannel.Items.Add("600V");
                    CBChannel.SelectedIndex = 0;
                    break;
            }
        }

        private void CBItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*
                Item 1 CCM
                Item 2 EN on/off
                Item 3 VID
                Item 4 LT
             */

            if (CBItem.SelectedIndex >= 2) groupBox2.Visible = true;
            else groupBox2.Visible = false;


            switch (CBItem.SelectedIndex)
            {
                case 2:
                    groupBox2.Text = "VID Group Setting";
                    LTDG.Columns[1].HeaderText = "Addr (Hex)";
                    LTDG.Columns[2].HeaderText = "Hi (Hex)";
                    LTDG.Columns[3].HeaderText = "Low (Hex)";
                    break;
                case 3:
                    groupBox2.Text = "Eload LT Setting";
                    LTDG.Columns[1].HeaderText = "L1(A)";
                    LTDG.Columns[2].HeaderText = "L2(A)";
                    LTDG.Columns[3].HeaderText = "None Use";
                    break;

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

        private void BT_SaveSetting_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            savedlg.Filter = "settings|*.tb_info";

            if (savedlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = savedlg.FileName;
                SaveSettings(file_name);
            }
        }

        private void SaveSettings(string file)
        {
            string settings = "";

            settings = "0.Slave=$" + nuslave.Value + "$\r\n";
            settings += "1.WavePath=$" + tbWave.Text + "$\r\n";
            settings += "2.Vin=$" + tb_vinList.Text + "$\r\n";
            settings += "3.Item=$" + CBItem.SelectedIndex + "$\r\n";
            settings += "6.Tolerance=$" + nuToerance.Value + "$\r\n";
            settings += "7.ChamberEn=$" + (ck_chamber_en.Checked ? "1" : "0") + "$\r\n";
            settings += "8.Temp=$" + tb_templist.Text + "$\r\n";
            settings += "9.SteadyTime=$" + nu_steady.Value + "$\r\n";
            settings += "10.ScaleTime=$" + nu_ontime_scale.Value + "$\r\n";
            settings += "CH_num=$" + nuCH_number.Value + "$\r\n";
            settings += "11.CCM_Rows=$" + EloadDG_CCM.RowCount + "$\r\n";
            settings += "12.Freq_Rows=$" + FreqDG.RowCount + "$\r\n";
            settings += "13.Vout_Rows=$" + VoutDG.RowCount + "$\r\n";
            settings += "14.Item_Rows=$" + LTDG.RowCount + "$\r\n";
            settings += "15.Meas_Rows=$" + MeasDG.RowCount + "$\r\n";

            for (int i = 0; i < EloadDG_CCM.RowCount; i++)
            {
                settings += (i + 16).ToString() + ".Rail=$" + EloadDG_CCM[0, i].Value.ToString() + "$\r\n";
                settings += (i + 17).ToString() + ".En_Addr=$" + EloadDG_CCM[1, i].Value.ToString() + "$\r\n";
                settings += (i + 18).ToString() + ".EN_On=$" + EloadDG_CCM[2, i].Value.ToString() + "$\r\n";
                settings += (i + 19).ToString() + ".EN_Off=$" + EloadDG_CCM[3, i].Value.ToString() + "$\r\n";
                settings += (i + 20).ToString() + ".CCM_Loading=$" + EloadDG_CCM[4, i].Value.ToString() + "$\r\n";
                settings += (i + 21).ToString() + ".Full_Loading=$" + EloadDG_CCM[5, i].Value.ToString() + "$\r\n";
            }

            for (int i = 0; i < FreqDG.RowCount; i++)
            {
                settings += (i + 22).ToString() + ".Rail=$" + FreqDG[0, i].Value.ToString() + "$\r\n";
                settings += (i + 23).ToString() + ".Freq_Addr=$" + FreqDG[1, i].Value.ToString() + "$\r\n";
                settings += (i + 24).ToString() + ".Freq_Data=$" + FreqDG[2, i].Value.ToString() + "$\r\n";
                settings += (i + 25).ToString() + ".Freq_Des=$" + FreqDG[3, i].Value.ToString() + "$\r\n";
            }

            for (int i = 0; i < VoutDG.RowCount; i++)
            {
                settings += (i + 26).ToString() + ".Rail=$" + VoutDG[0, i].Value.ToString() + "$\r\n";
                settings += (i + 27).ToString() + ".Vout_Addr=$" + VoutDG[1, i].Value.ToString() + "$\r\n";
                settings += (i + 28).ToString() + ".Vout_Data=$" + VoutDG[2, i].Value.ToString() + "$\r\n";
                settings += (i + 29).ToString() + ".Vout_Des=$" + VoutDG[3, i].Value.ToString() + "$\r\n";
            }

            for (int i = 0; i < LTDG.RowCount; i++)
            {
                settings += (i + 30).ToString() + ".Rail=$" + LTDG[0, i].Value.ToString() + "$\r\n";
                settings += (i + 31).ToString() + ".Item_setting1=$" + LTDG[1, i].Value.ToString() + "$\r\n";
                settings += (i + 32).ToString() + ".Item_setting2=$" + LTDG[2, i].Value.ToString() + "$\r\n";
                settings += (i + 33).ToString() + ".Item_setting2=$" + LTDG[3, i].Value.ToString() + "$\r\n";
            }

            for (int i = 0; i < MeasDG.RowCount; i++)
            {
                settings += (i + 34).ToString() + ".Rail=$" + MeasDG[0, i].Value.ToString() + "$\r\n";
                settings += (i + 35).ToString() + ".Meas1=$" + MeasDG[1, i].Value.ToString() + "$\r\n";
                settings += (i + 36).ToString() + ".Meas2=$" + MeasDG[2, i].Value.ToString() + "$\r\n";
                settings += (i + 37).ToString() + ".Meas3=$" + MeasDG[3, i].Value.ToString() + "$\r\n";
            }



            using (StreamWriter sw = new StreamWriter(file))
            {
                sw.Write(settings);
            }
        }

        private void BT_LoadSetting_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "settings|*.tb_info";
            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                LoadSettings(opendlg.FileName);
            }
        }

        private void LoadSettings(string file)
        {
            object[] obj_arr = new object[]
            {
                nuslave, tbWave, tb_vinList, CBItem, nuToerance, ck_chamber_en, tb_templist, nu_steady, nu_ontime_scale, nuCH_number,
                EloadDG_CCM, FreqDG, VoutDG, LTDG, MeasDG
            };
            List<string> info = new List<string>();
            using (StreamReader sr = new StreamReader(file))
            {
                string pattern = @"(?<=\$)(.*)(?=\$)";
                MatchCollection matches;
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    Regex rg = new Regex(pattern);
                    matches = rg.Matches(line);
                    Match match = matches[0];
                    info.Add(match.Value);
                }

                int idx = 0;
                for (int i = 0; i < obj_arr.Length; i++)
                {
                    switch (obj_arr[i].GetType().Name)
                    {
                        case "TextBox":
                            ((TextBox)obj_arr[i]).Text = info[i];
                            break;
                        case "CheckBox":
                            ((CheckBox)obj_arr[i]).Checked = info[i] == "1" ? true : false;
                            break;
                        case "NumericUpDown":
                            ((NumericUpDown)obj_arr[i]).Value = Convert.ToDecimal(info[i]);
                            break;
                        case "ComboBox":
                            ((ComboBox)obj_arr[i]).SelectedIndex = Convert.ToInt32(info[i]);
                            break;
                        case "DataGridView":
                            ((DataGridView)obj_arr[i]).RowCount = Convert.ToInt32(info[i]);
                            ((DataGridView)obj_arr[i + 1]).RowCount = Convert.ToInt32(info[i + 1]);
                            ((DataGridView)obj_arr[i + 2]).RowCount = Convert.ToInt32(info[i + 2]);
                            ((DataGridView)obj_arr[i + 3]).RowCount = Convert.ToInt32(info[i + 3]);
                            ((DataGridView)obj_arr[i + 4]).RowCount = Convert.ToInt32(info[i + 4]);

                            idx = i + 4;
                            goto fullDG;

                            break;
                    }
                }

            fullDG:
                for (int i = 0; i < EloadDG_CCM.RowCount; i++)
                {
                    EloadDG_CCM[0, i].Value = Convert.ToString(info[idx + 1]); // start
                    EloadDG_CCM[1, i].Value = Convert.ToString(info[idx + 2]); // step
                    EloadDG_CCM[2, i].Value = Convert.ToString(info[idx + 3]); // stop
                    EloadDG_CCM[3, i].Value = Convert.ToString(info[idx + 4]); // start
                    EloadDG_CCM[4, i].Value = Convert.ToString(info[idx + 5]); // step
                    EloadDG_CCM[5, i].Value = Convert.ToString(info[idx + 6]); // stop
                    idx += 6;
                }

                for (int i = 0; i < FreqDG.RowCount; i++)
                {
                    FreqDG[0, i].Value = Convert.ToString(info[idx + 1]); // start
                    FreqDG[1, i].Value = Convert.ToString(info[idx + 2]); // step
                    FreqDG[2, i].Value = Convert.ToString(info[idx + 3]); // stop
                    FreqDG[3, i].Value = Convert.ToString(info[idx + 4]); // start
                    idx += 4;
                }

                for (int i = 0; i < VoutDG.RowCount; i++)
                {
                    VoutDG[0, i].Value = Convert.ToString(info[idx + 1]); // start
                    VoutDG[1, i].Value = Convert.ToString(info[idx + 2]); // step
                    VoutDG[2, i].Value = Convert.ToString(info[idx + 3]); // stop
                    VoutDG[3, i].Value = Convert.ToString(info[idx + 4]); // start
                    idx += 4;
                }

                for (int i = 0; i < LTDG.RowCount; i++)
                {
                    LTDG[0, i].Value = Convert.ToString(info[idx + 1]); // start
                    LTDG[1, i].Value = Convert.ToString(info[idx + 2]); // step
                    LTDG[2, i].Value = Convert.ToString(info[idx + 3]); // stop
                    LTDG[3, i].Value = Convert.ToString(info[idx + 4]); // start
                    idx += 4;
                }

                for (int i = 0; i < MeasDG.RowCount; i++)
                {
                    MeasDG[0, i].Value = Convert.ToString(info[idx + 1]);
                    MeasDG[1, i].Value = Convert.ToString(info[idx + 2]);
                    MeasDG[2, i].Value = Convert.ToString(info[idx + 3]);
                    MeasDG[3, i].Value = Convert.ToString(info[idx + 4]);
                    idx += 4;
                }

            }
        }

        private void EloadDG_CCM_Enter(object sender, EventArgs e)
        {
            //ScopeCH.Items.Add("CH" + i);
            //DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)EloadDG_CCM[6, 0];
            //Console.WriteLine(comboBoxCell.Value);
        }
    }
}


