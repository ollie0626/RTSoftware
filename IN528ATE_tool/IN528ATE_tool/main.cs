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
using System.Diagnostics;
using System.Threading;
using System.IO;

namespace IN528ATE_tool
{
    public partial class main : Sunny.UI.UIForm
    {
        delegate void MyDelegate();
        MyDelegate Message;
        FolderBrowserDialog FolderBrow;

        RTBBControl RTDev;
        MyLib myLib;
        int SteadyTime;
        int TCPServerTime;

        ParameterizedThreadStart p_thread;
        Thread ATETask;

        ATE_OutputRipple _ate_ripple;
        ATE_CodeInrush _ate_code_inrush;
        ATE_PowerOn _ate_poweron;
        TaskRun[] ate_table;

        string[] tempList;
        string templist;
        int item_sel;

        private void GUIInit()
        {
            /* class init */
            this.Text = "ATE Tool v2.12";
            RTDev = new RTBBControl();
            myLib = new MyLib();

            _ate_ripple = new ATE_OutputRipple();
            _ate_code_inrush = new ATE_CodeInrush();
            _ate_poweron = new ATE_PowerOn();

            led_osc.Color = Color.Red;
            led_power.Color = Color.Red;
            led_eload.Color = Color.Red;
            led_37940.Color = Color.Red;
            led_chamber.Color = Color.Red;
            cb_item.SelectedIndex = 0;
            ate_table = new TaskRun[] { _ate_ripple, _ate_code_inrush, _ate_poweron };
            Message = new MyDelegate(MessageCallback);
            TCPServerTime = 28800;

            test_parameter.run_stop = false;
            test_parameter.chamber_en = false;
        }

        public main()
        {
            InitializeComponent();
            GUIInit();

            // 2 ^ 10
            Console.WriteLine(Math.Pow(10, 10));
        }

        private void MessageCallback()
        {
            MessageBox.Show("Callback message test!!");
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

        private bool CheckTCPConnect_MS(int Time_1S)
        {
            ChamberCtr.CreatTCPServer();
            for (int i = 0; i < Time_1S; ++i)
            {
                if (ChamberCtr.myTcpListener.Pending())
                {
                    ChamberCtr.mySocket = ChamberCtr.myTcpListener.AcceptSocket();
                    ChamberCtr.mySocket.Close();
                    ChamberCtr.myTcpListener.Stop();
                    ChamberCtr.IsTCPConnected = true;
                    return true;
                }
                System.Threading.Thread.Sleep(1000);
                Console.WriteLine("wait for slave ~~");
                // test
                if (InsControl._chamber != null) InsControl._chamber.GetChamberTemperature();
            }
            ChamberCtr.IsTCPConnected = false;
            return false;
        }

        private bool CheckTCPConnect_SV(int Time_1S)
        {
            for (int i = 0; i < Time_1S; ++i)
            {
                if (ChamberCtr.CreatSlaveConnect())
                {
                    ChamberCtr.IsTCPConnected = true;
                    return true;
                }
                System.Threading.Thread.Sleep(1000);
                Console.WriteLine("wait for master");
            }
            ChamberCtr.IsTCPConnected = false;
            return false;
        }

        private bool CheckTCPConnect(int Time_1S)
        {
            if (ck_multi_chamber.Checked)
            {
                if (!ck_slave.Checked)
                {
                    if (!CheckTCPConnect_MS(Time_1S))
                    {
                        return false;
                    }
                }
                else
                {
                    if (!CheckTCPConnect_SV(Time_1S))
                    {
                        return false;
                    }
                    System.Threading.Thread.Sleep(1000);
                }
            }
            return true;
        }

        private Task<bool> TaskConnect(int Time_1S)
        {
            return Task.Factory.StartNew(() => CheckTCPConnect(Time_1S));
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

        private void uibt_pause_Click(object sender, EventArgs e)
        {
            if (ATETask == null) return;
            System.Threading.ThreadState state = ATETask.ThreadState;
            if (state == System.Threading.ThreadState.Running || state == System.Threading.ThreadState.WaitSleepJoin)
            {
                ATETask.Suspend();
                uibt_pause.SymbolColor = Color.Red;
            }
            else if (state == System.Threading.ThreadState.Suspended)
            {
                ATETask.Resume();
                uibt_pause.SymbolColor = Color.White;
            }
        }

        private void uibt_run_Click(object sender, EventArgs e)
        {
            try
            {
                templist = tb_templist.Text;
                tempList = tb_templist.Text.Split(',');
                uiProcessBar1.Maximum = (int)nu_steady.Value;
                RTDev.BoadInit();
                /* test conditons assign */
                test_parameter.chamber_en = ck_chaber_en.Checked;
                test_parameter.run_stop = false;
                test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
                test_parameter.IoutList = tb_ioutList.Text.Split(',').Select(double.Parse).ToList();
                test_parameter.specify_id = (byte)nu_specify.Value;
                test_parameter.slave = (byte)nu_slave.Value;
                test_parameter.binFolder = textBox1.Text;
                test_parameter.specify_bin = textBox2.Text;
                test_parameter.waveform_path = tbWave.Text;
                test_parameter.time_scale_ms = (double)nu_time_scale.Value;
                test_parameter.addr = (byte)nu_addr.Value;
                test_parameter.max = (byte)nu_code_max.Value;
                test_parameter.min = (byte)nu_code_min.Value;
                test_parameter.vol_max = (double)nu_vol_max.Value;
                test_parameter.vol_min = (double)nu_vol_min.Value;
                test_parameter.all_en = ck_all_test.Checked;
                test_parameter.trigger_vin_en = ck_vin_trigger.Checked;
                item_sel = cb_item.SelectedIndex;

                ChamberCtr.ChamberName = tb_chamber.Text;
                SteadyTime = (int)nu_steady.Value;

                if (ck_multi_chamber.Checked && ck_chaber_en.Checked)
                {
                    ATETask = new Thread(MultiChamber_Task);
                    ATETask.Start();
                }
                else if (ck_chaber_en.Checked)
                {
                    ATETask = new Thread(Chamber_Task);
                    ATETask.Start();
                }
                else
                {
                    // single no chamber conditions
                    _ate_ripple.temp = 25;
                    _ate_poweron.temp = 25;
                    _ate_code_inrush.temp = 25;
                    if (ck_all_test.Checked)
                    {

                        ATETask = new Thread(Run_Task_Flow);
                        ATETask.Start();
                    }
                    else
                    {
                        p_thread = new ParameterizedThreadStart(Run_Single_Task);
                        ATETask = new Thread(p_thread);
                        ATETask.Start(cb_item.SelectedIndex);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
            }
        }

        // ATE Process: MultiChamber_Task, Chamber_Task, Run_Task_Flow, Run_Single_Task
        public async void MultiChamber_Task()
        {
            ChamberCtr.ChamberName = tb_chamber.Text;
            ChamberCtr.CreatShareChamberFolder();
            if (!ck_slave.Checked)
            {
                // master
                ChamberCtr.DeleteShareChamberFile();
                ChamberCtr.CreatTempList(templist);
            }
            else
            {
                // slave
                templist = ChamberCtr.ReadTempList();
            }

            if (!await TaskConnect(300)) return;// connect

            for (int i = 0; i < tempList.Length; i++)
            {
                if (!await TaskConnect(3000 + SteadyTime * 60)) return;
                else
                {
                    SteadyTime = (int)nu_steady.Value;
                    // new construct and connect chamber
                    InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                    InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                    bool res = InsControl._chamber.InsState();

                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    await InsControl._chamber.ChamberStable(templist[i]);

                    for (; SteadyTime > 0;)
                    {
                        await TaskRecount();
                        uiProcessBar1.Value = SteadyTime;
                        label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));

                    }
                    if (!await TaskConnect(TCPServerTime)) TCPServerTime = 0;
                }

                // ripple test
                _ate_ripple.temp = Convert.ToDouble(tempList[i]);
                _ate_code_inrush.temp = Convert.ToDouble(tempList[i]);
                _ate_poweron.temp = Convert.ToDouble(tempList[i]);

                if (!test_parameter.all_en)
                {
                    switch (item_sel)
                    {
                        case 0:
                            _ate_ripple.ATETask();
                            break;
                        case 1:
                            _ate_code_inrush.ATETask();
                            break;
                        case 2:
                            _ate_poweron.ATETask();
                            break;
                    }
                }
                else
                {
                    _ate_ripple.ATETask();
                    _ate_code_inrush.ATETask();
                }

                // test finished
                if (ck_multi_chamber.Checked && ck_slave.Checked)
                {
                    // slave
                    if (!await TaskConnect(TCPServerTime)) return;
                }
                else
                {
                    // server
                    await TaskConnect(TCPServerTime);
                }
                if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
            }
        }

        public async void Chamber_Task()
        {
            for (int i = 0; i < tempList.Length; i++)
            {
                if (!Directory.Exists(tbWave.Text + @"\" + tempList[i] + "C"))
                {
                    Directory.CreateDirectory(tbWave.Text + @"\" + tempList[i] + "C");
                }
                test_parameter.waveform_path = tbWave.Text + @"\" + tempList[i] + "C";

                SteadyTime = (int)nu_steady.Value;
                InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                await InsControl._chamber.ChamberStable(Convert.ToDouble(tempList[i]));
                for (; SteadyTime > 0;)
                {
                    await TaskRecount();
                    uiProcessBar1.Value = SteadyTime;
                    label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
                    //label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString();
                }
                _ate_ripple.temp = Convert.ToDouble(tempList[i]);
                _ate_code_inrush.temp = Convert.ToDouble(tempList[i]);
                _ate_poweron.temp = Convert.ToDouble(tempList[i]);

                if (!test_parameter.all_en)
                {
                    //await Ripple_Task(item_sel);

                    switch (item_sel)
                    {
                        case 0:
                            _ate_ripple.ATETask();
                            break;
                        case 1:
                            _ate_code_inrush.ATETask();
                            break;
                        case 2:
                            _ate_poweron.ATETask();
                            break;
                    }
                }
                else
                {
                    _ate_ripple.ATETask();
                    _ate_code_inrush.ATETask();
                }

            }
            if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
        }

        private void Run_Task_Flow()
        {
            for (int i = 0; i < 2; i++)
            {
                switch (i)
                {
                    case 0:
                        _ate_ripple.ATETask();
                        break;
                    case 1:
                        _ate_code_inrush.ATETask();
                        break;
                }
            }
            Message.Invoke();
        }

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].ATETask();
        }

        private void uiSymbolButton1_Click(object sender, EventArgs e)
        {
            test_parameter.run_stop = true;
            if (ATETask != null)
            {
                if (ATETask.IsAlive)
                {
                    System.Threading.ThreadState state = ATETask.ThreadState;
                    if (state == System.Threading.ThreadState.Suspended) ATETask.Resume();
                    ATETask.Abort();
                    MessageBox.Show("ATE Task Stop !!", "ATE Tool", MessageBoxButtons.OK);
                    InsControl._power.AutoPowerOff();
                }
            }
        }

        private void ck_multi_chamber_CheckedChanged(object sender, EventArgs e)
        {
            ck_chaber_en.Checked = true;
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

        private void uibut_binfile_Click(object sender, EventArgs e)
        {
            FolderBrow = new FolderBrowserDialog();
            if (FolderBrow.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = FolderBrow.SelectedPath;
            }
        }

        private void uibt_Wavepath_Click(object sender, EventArgs e)
        {
            FolderBrow = new FolderBrowserDialog();
            if (FolderBrow.ShowDialog() == DialogResult.OK)
            {
                tbWave.Text = FolderBrow.SelectedPath;
            }
        }

        private void uibt_specify_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "bin File (*.bin)|*.bin";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox2.Text = openFileDialog1.FileName;
            }
        }

        private void ck_all_test_CheckedChanged(object sender, EventArgs e)
        {
            if (ck_all_test.Checked)
            {
                cb_item.Enabled = false;
            }
            else
            {
                cb_item.Enabled = true;
            }
        }

        private void cb_item_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cb_item.SelectedIndex == 2)
            { ck_vin_trigger.Checked = false; }
            else
            { ck_vin_trigger.Checked = false; }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            IN528ATE_tool.Properties.Settings.Default.binpath = this.textBox1.Text;
            IN528ATE_tool.Properties.Settings.Default.specifypath = this.textBox2.Text;
            IN528ATE_tool.Properties.Settings.Default.wavepath = this.tbWave.Text;
            Properties.Settings.Default.vinList = tb_vinList.Text;
            Properties.Settings.Default.IoutList = tb_ioutList.Text;
            Properties.Settings.Default.itemSel = cb_item.SelectedIndex;
            IN528ATE_tool.Properties.Settings.Default.Save();
        }

        private void main_Load(object sender, EventArgs e)
        {
            this.textBox1.Text = IN528ATE_tool.Properties.Settings.Default.binpath;
            this.textBox2.Text = IN528ATE_tool.Properties.Settings.Default.specifypath;
            this.tbWave.Text = IN528ATE_tool.Properties.Settings.Default.wavepath;
            tb_vinList.Text = Properties.Settings.Default.vinList;
            tb_ioutList.Text = Properties.Settings.Default.IoutList;
            cb_item.SelectedIndex = Properties.Settings.Default.itemSel;
        }

        private void uiSymbolButton2_Click(object sender, EventArgs e)
        {
            RTDev.BoadInit();
            RTDev.I2C_WriteBin(0x9E >> 1, 0x00, textBox2.Text);
        }


    }
}
