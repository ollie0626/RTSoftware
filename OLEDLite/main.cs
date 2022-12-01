using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MaterialSkin;
using MaterialSkin.Controls;
using System.IO;
using InsLibDotNet;
using System.Threading;
using System.Net.Sockets;
using System.Net;
using System.Diagnostics;


namespace OLEDLite
{
    public partial class main : Form
    {
        private static string ver = "v1.4";
        private string win_name = "OLED sATE tool " + ver;
        //private readonly MaterialSkinManager materialSkinManager;

        private ParameterizedThreadStart p_thread;
        private Thread ATETask;
        // test item instance
        private ATE_TDMA _ate_tdma = new ATE_TDMA();
        private ATE_OutputRipple _ate_outputripple = new ATE_OutputRipple();
        private ATE_CodeInrush _ate_codeinrush = new ATE_CodeInrush();
        private ATE_Eff _ate_eff = new ATE_Eff();
        private ATE_Line _ate_line = new ATE_Line();
        private ATE_UVPDly _ate_uvp_dly = new ATE_UVPDly();
        private ATE_UVPLevel _ate_uvp_level = new ATE_UVPLevel();
        private ATE_CurrentLimit _ate_ocp = new ATE_CurrentLimit();

        private TaskRun[] ate_table;
        ChamberCtr chamberCtr = new ChamberCtr();
        //private static SynchronizationContext _syncContext = null;

        public void UpdateRunButton()
        {
            this.Invoke((Action)(() => bt_run.Enabled = true));
        }

        public main()
        {
            InitializeComponent();
            materialTabSelector1.Width = this.Width;
            materialTabSelector1.Height = 25;
            this.Text = win_name;
            CK_I2c.Checked = true;
            CK_I2c.Checked = false;
            nu_load1.Enabled = false;
            ck_ch1_en.Enabled = false;
            cb_mode_sel.SelectedIndex = 0;
            materialTabControl1.SelectedIndex = 1;
            cb_item.SelectedIndex = 0;
            nu_swire_num.Value = 1;
            CBEload.SelectedIndex = 0;
            CBIinSelect.SelectedIndex = 0;

            swireTable.ColumnCount = 2;
            swireTable.Columns[0].Name = "ESwire";
            swireTable.Columns[0].Width = 150;
            swireTable.Columns[1].Name = "ASwire";
            swireTable.Columns[1].Width = 150;
            //RB_ASwire.Checked = true;
            ATEItemInit();
        }

        private void ATEItemInit()
        {
            ate_table = new TaskRun[] { 
                _ate_tdma,
                _ate_outputripple,
                _ate_codeinrush,
                _ate_eff,
                _ate_line,
                _ate_uvp_dly,
                _ate_uvp_level,
                _ate_ocp
            };
        }

        private void main_Resize(object sender, EventArgs e)
        {
            materialTabSelector1.Width = this.Width;
        }

        private int ConnectFunc(string res, int ins_sel)
        {
            switch (ins_sel)
            {
                case 0:
                    InsControl._scope = new AgilentOSC(res);
                    break;
                case 1:
                    InsControl._power = new PowerModule(res);

                    break;
                case 2:
                    InsControl._eload = new EloadModule(res);

                    break;
                case 3:
                    InsControl._34970A = new MultiChannelModule(res);

                    break;
                case 4:
                    InsControl._funcgen = new FuncGenModule(res);

                    break;
                case 5:
                    InsControl._dmm1 = new DMMModule(res);

                    break;
                case 6:
                    InsControl._dmm2 = new DMMModule(res);

                    break;
                case 7:
                    InsControl._chamber = new ChamberModule(res);
                    break;
            }

            return 0;
        }

        private Task<int> ConnectTask(string res, int ins_sel)
        {
            return Task.Factory.StartNew(() => ConnectFunc(res, ins_sel));
        }

        private async void bt_connect_Click(object sender, EventArgs e)
        {
            await ConnectTask(tb_res_scope.Text, 0);
            await ConnectTask(tb_res_power.Text, 1);
            await ConnectTask(tb_res_eload.Text, 2);
            await ConnectTask(tb_res_daq.Text, 3);
            await ConnectTask(tb_res_func.Text, 4);
            await ConnectTask(tb_res_meter_in.Text, 5);
            await ConnectTask(tb_res_meter_out.Text, 6);
            await ConnectTask(tb_res_chamber.Text, 7);

            MyLib.Delay1s(1);
            if (InsControl._scope.InsState()) ck_scope.Checked = true;
            else ck_scope.Checked = false;
            if (InsControl._power.InsState()) ck_power.Checked = true;
            else ck_power.Checked = false;
            if (InsControl._eload.InsState()) ck_eload.Checked = true;
            else ck_eload.Checked = true;
            if (InsControl._34970A.InsState()) ck_daq.Checked = true;
            else ck_daq.Checked = false;
            if (InsControl._funcgen.InsState()) ck_func.Checked = true;
            else ck_func.Checked = false;
            if (InsControl._dmm1.InsState()) ck_meter_in.Checked = true;
            else ck_meter_in.Checked = false;
            if (InsControl._dmm2.InsState()) ck_meter_out.Checked = true;
            else ck_meter_out.Checked = false;
            if (InsControl._chamber.InsState()) ck_chamber.Checked = true;
            else ck_chamber.Checked = false;
        }

        private void bt_scanIns_Click(object sender, EventArgs e)
        {
            string[] ins_list = ViCMD.ScanIns();
            foreach (string ins in ins_list)
            {
                list_ins.Items.Add(ins);
            }
        }

        private void bt_eload_add_Click(object sender, EventArgs e)
        {
            Eload_DG.RowCount = Eload_DG.RowCount + 1;
        }

        private void bt_eload_sub_Click(object sender, EventArgs e)
        {
            if (Eload_DG.RowCount == 0) return;
            Eload_DG.RowCount = Eload_DG.RowCount - 1;
        }

        private void bt_func_set_Click(object sender, EventArgs e)
        {
            if (InsControl._funcgen == null ||
                tb_High_level.Text  == "" ||
                tb_Low_level.Text == "") return;

            string[] hi_str = tb_High_level.Text.Split(',');
            string[] lo_str = tb_Low_level.Text.Split(',');

            double hi = Convert.ToDouble(hi_str[0]);
            double lo = Convert.ToDouble(lo_str[0]);

            InsControl._funcgen.CH1_ContinuousMode();
            InsControl._funcgen.CH1_PulseMode();
            InsControl._funcgen.CH1_Frequency((double)(nu_Freq.Value * 1000));
            InsControl._funcgen.CH1_DutyCycle((double)nu_duty.Value);
            InsControl._funcgen.CH1_LoadImpedanceHiz();
            InsControl._funcgen.SetCH1_TrTfFunc((double)nu_Tr.Value, (double)nu_Tf.Value);
            InsControl._funcgen.CHl1_HiLevel(hi);
            InsControl._funcgen.CH1_LoLevel(lo);
            InsControl._funcgen.CH1_On();
        }

        private bool test_parameter_copy()
        {
            // interface
            test_parameter.i2c_enable = CK_I2c.Checked;
            test_parameter.slave = (byte)nu_slave.Value;
            test_parameter.bin_path = tb_bin.Text;
            test_parameter.wave_path = tb_wave_path.Text;
            test_parameter.special_file = tb_initial_bin.Text;

            // function gen
            test_parameter.Freq = (double)nu_Freq.Value;
            test_parameter.duty = (double)nu_duty.Value;
            test_parameter.tr = (double)nu_Tr.Value;
            test_parameter.tf = (double)nu_Tf.Value;

            test_parameter.HiLo_table.Clear();
            test_parameter.HiLevel = tb_High_level.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.LoLevel = tb_Low_level.Text.Split(',').Select(double.Parse).ToList();
            
            Hi_Lo level = new Hi_Lo();
            for (int hi_index = 0; hi_index < test_parameter.HiLevel.Count; hi_index++)
            {
                for (int lo_index = 0; lo_index < test_parameter.LoLevel.Count; lo_index++)
                {
                    level.Highlevel = test_parameter.HiLevel[hi_index];
                    level.LowLevel = test_parameter.LoLevel[lo_index];
                    test_parameter.HiLo_table.Add(level);
                }
            }

            // fix iout different channel
            test_parameter.eload_ch_select = CBEload.SelectedIndex;
            test_parameter.eload_iin_select = CBIinSelect.SelectedIndex;
            test_parameter.eload_en = new bool[4] { ck_ch1_en.Checked, 
                                                    ck_ch2_en.Checked, 
                                                    ck_ch3_en.Checked, 
                                                    ck_ch4_en.Checked };
            test_parameter.eload_iout = new double[4] { (double)nu_load1.Value,
                                                        (double)nu_load2.Value,
                                                        (double)nu_load3.Value,
                                                        (double)nu_load4.Value };

            //0.TDMA
            //1.Ripple
            //2.Code Inrush
            //3.Eff / Load regulation
            //4.Line regulation
            //5.UVP Delay
            //6.UVP Level

            if (cb_item.SelectedIndex == 4)
            {
                //test_parameter.vinList = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                test_parameter.vinList = MyLib.TBData(tb_Vin);
            }
            else
            {
                test_parameter.vinList = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
            }


            switch (cb_item.SelectedIndex)
            {
                // TDMA, Code Inrush, Line regulation
                case 0:
                case 2:
                case 4:
                case 5:
                case 6:
                    test_parameter.ioutList = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    break;
                // Ripple, Efficicency
                case 1:
                case 3:
                    test_parameter.ioutList = MyLib.DGData(Eload_DG);
                    break;
            }

            if (!CK_I2c.Checked)
            {
                test_parameter.ESwire_state = CK_ESwire.Checked;
                test_parameter.ASwire_state = CK_ASwire.Checked;
                test_parameter.ENVO4_state = CK_ENVO4.Checked;

                test_parameter.swire_cnt = (int)nu_swire_num.Value;
                test_parameter.ESwireList.Clear();
                test_parameter.ASwireList.Clear();

                for (int i = 0; i < swireTable.RowCount; i++)
                {
                    test_parameter.ESwireList.Add((string)swireTable[0, i].Value);
                    test_parameter.ASwireList.Add((string)swireTable[1, i].Value);
                    if (swireTable[0, i].Value == null || swireTable[1, i].Value == null)
                    {
                        MessageBox.Show("Please input swire conditions", win_name);
                        return false;
                    }
                }
            }

            test_parameter.tempList = tb_templist.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.steadyTime = (int)nu_steady.Value;
            test_parameter.run_stop = false;
            test_parameter.burst_period = (1 / ((double)nu_Sys_clk.Value * Math.Pow(10, 6)));

            // code Inrush
            test_parameter.addr = (byte)nu_addr.Value;
            test_parameter.code_min = (int)nu_code_min.Value;
            test_parameter.code_max = (int)nu_code_max.Value;
            test_parameter.vol_max = (double)nu_vol_max.Value;
            test_parameter.vol_min = (double)nu_vol_min.Value;
            test_parameter.ontime_scale_ms = (double)nu_timeScale.Value;
            test_parameter.CodeInrush_ESwire = RB_ESwire.Checked;

            // test condition
            test_parameter.vin_info = tb_Vin.Text + "V";
            test_parameter.eload_info = test_parameter.ioutList[0] * 1000 + "mA~" + test_parameter.ioutList[test_parameter.ioutList.Count - 1] * 1000 + "mA";
            test_parameter.ver_info = win_name;
            test_parameter.date_info = DateTime.Now.ToString("yyyyMMdd_hhmm");

            // Eload CV setting
            test_parameter.cv_setting = (double)nu_CVSetting.Value;
            test_parameter.cv_step = (double)nu_CVStep.Value;
            test_parameter.cv_wait = (double)nu_CVwait.Value;

            return true;
        }

        private void bt_run_Click(object sender, EventArgs e)
        {
            try
            {
                bt_run.Enabled = false;
                if (!test_parameter_copy())
                {
                    bt_run.Enabled = true; 
                    return;
                }

                if(ck_multi_chamber.Checked && ck_chamber_en.Checked)
                {
                    p_thread = new ParameterizedThreadStart(multi_ate_process);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(cb_item.SelectedIndex);
                }
                else if(ck_chamber_en.Checked)
                {
                    p_thread = new ParameterizedThreadStart(Chamber_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(cb_item.SelectedIndex);
                }
                else
                {
                    // Lab run ate item
                    p_thread = new ParameterizedThreadStart(Run_Single_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(cb_item.SelectedIndex);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
            }
        }
        // TODO: FFT funciton
        //:MEASure:FFT:MAGNitude?\sFUNCtion1,1

        private void nu_swire_num_ValueChanged(object sender, EventArgs e)
        {
            swireTable.RowCount = (int)nu_swire_num.Value;
        }

        private void cb_item_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch(cb_item.SelectedIndex)
            {
                case 0:
                    // TDMA
                    group_power.Enabled = false;
                    Eload_DG.Enabled = false;
                    ck_Iout_mode.Checked = false;
                    ck_Iout_mode.Enabled = false;
                    tb_Iout.Enabled = true;
                    bt_eload_add.Enabled = false;
                    bt_eload_sub.Enabled = false;
                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = false;

                    CBIinSelect.SelectedIndex = 0;
                    CBIinSelect.Enabled = false;
                    break;
                case 1:
                    // output ripple
                    group_power.Enabled = true;
                    Eload_DG.Enabled = true;
                    ck_Iout_mode.Checked = true;
                    ck_Iout_mode.Enabled = true;
                    tb_Iout.Enabled = false;
                    bt_eload_add.Enabled = true;
                    bt_eload_sub.Enabled = true;
                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = false;
                    CBIinSelect.SelectedIndex = 0;
                    CBIinSelect.Enabled = false;
                    break;
                case 2:
                    // code Inrush
                    group_power.Enabled = true;
                    Eload_DG.Enabled = true;
                    ck_Iout_mode.Checked = true;
                    ck_Iout_mode.Enabled = true;
                    tb_Iout.Enabled = false;
                    bt_eload_add.Enabled = true;
                    bt_eload_sub.Enabled = true;
                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = false;

                    CBIinSelect.Enabled = false;
                    break;
                case 3:
                    // Eff and Load regulation
                    group_power.Enabled = true;
                    Eload_DG.Enabled = true;
                    ck_Iout_mode.Checked = true;
                    ck_Iout_mode.Enabled = true;
                    tb_Iout.Enabled = false;
                    bt_eload_add.Enabled = true;
                    bt_eload_sub.Enabled = true;
                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = true;

                    CBIinSelect.Enabled = true;
                    CBIinSelect.SelectedIndex = 0;
                    break;
                case 4:
                    // line regulation
                    group_power.Enabled = true;
                    Eload_DG.Enabled = false;
                    ck_Iout_mode.Checked = false;
                    ck_Iout_mode.Enabled = false;
                    tb_Iout.Enabled = true;
                    bt_eload_add.Enabled = false;
                    bt_eload_sub.Enabled = false;

                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = true;

                    CBIinSelect.SelectedIndex = 1;
                    CBIinSelect.Enabled = false;
                    break;
                case 5:
                    // UVP Delay
                    group_power.Enabled = true;
                    Eload_DG.Enabled = false;

                    ck_Iout_mode.Checked = false;
                    ck_Iout_mode.Enabled = false;

                    tb_Iout.Enabled = true;
                    bt_eload_add.Enabled = false;
                    bt_eload_sub.Enabled = false;

                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = true;

                    CBIinSelect.SelectedIndex = 1;
                    CBIinSelect.Enabled = false;
                    break;
                case 6:
                    // UVP level
                    break;
                case 7:
                    // Current limit
                    group_power.Enabled = true;
                    Eload_DG.Enabled = false;

                    ck_Iout_mode.Checked = false;
                    ck_Iout_mode.Enabled = false;

                    tb_Iout.Enabled = true;
                    bt_eload_add.Enabled = false;
                    bt_eload_sub.Enabled = false;

                    CBEload.SelectedIndex = 0;
                    CBEload.Enabled = true;

                    CBIinSelect.SelectedIndex = 1;
                    CBIinSelect.Enabled = false;
                    break;

            }
        }

        private void ck_multi_chamber_CheckedChanged(object sender, EventArgs e)
        {
            ck_chamber_en.Checked = true;
        }

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
            UpdateRunButton();
        }

        // ------------------------------------------------------------------------------------------
        // Chamber Thread Event
        // ------------------------------------------------------------------------------------------

        private bool RecountTime()
        {
            test_parameter.steadyTime--; System.Threading.Thread.Sleep(1000);
            return true;
        }

        private Task<bool> TaskRecount()
        {
            return Task.Factory.StartNew(() => RecountTime());
        }

        private void GetTemperature(string input)
        {
            test_parameter.tempList.Clear();
            string[] temp = input.Split(',');
            foreach (string str in temp)
            {
                test_parameter.tempList.Add(Convert.ToInt32(str));
            }
        }

        private async void Chamber_Task(object idx)
        {
            for (int i = 0; i < test_parameter.tempList.Count; i++)
            {
                if (!Directory.Exists(tb_wave_path.Text + @"\" + test_parameter.tempList[i] + "C"))
                {
                    Directory.CreateDirectory(tb_wave_path.Text + @"\" + test_parameter.tempList[i] + "C");
                }
                test_parameter.wave_path = tb_wave_path.Text + @"\" + test_parameter.tempList[i] + "C";

                test_parameter.steadyTime = (int)nu_steady.Value;
                // chamber control
                InsControl._chamber = new ChamberModule(tb_res_chamber.Text);
                InsControl._chamber.ConnectChamber(tb_res_chamber.Text);
                InsControl._chamber.ChamberOn(test_parameter.tempList[i]);
                InsControl._chamber.ChamberOn(test_parameter.tempList[i]);

                await InsControl._chamber.ChamberStable(test_parameter.tempList[i]);

                for (; test_parameter.steadyTime > 0;)
                {
                    await TaskRecount();
                    progressBar1.Value = test_parameter.steadyTime;
                    label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " 
                    + (test_parameter.steadyTime / 60).ToString() + ":" 
                    + (test_parameter.steadyTime % 60).ToString()));
                }

                ate_table[(int)idx].temp = test_parameter.tempList[i];
                ate_table[(int)idx].ATETask();
            }
            if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
            UpdateRunButton();
        }

        private async void multi_ate_process(object idx)
        {
            int timer;
            //MyLib myLib = new MyLib();
            //myLib.time = (int)nu_steady.Value;
            chamberCtr.Init(tb_templist.Text);

            if (chamberCtr.Role == "Master")
            {
                GetTemperature(tb_templist.Text);
                chamberCtr.Dispose();

                foreach (int Temp in test_parameter.tempList)
                {
                    ate_table[(int)idx].temp = Temp;

                    Console.WriteLine("StartTime：{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    test_parameter.steadyTime = (int)nu_steady.Value;
                    InsControl._chamber = new ChamberModule(tb_res_chamber.Text);
                    InsControl._chamber.ConnectChamber(tb_res_chamber.Text);
                    bool res = InsControl._chamber.InsState();
                    InsControl._chamber.ChamberOn(Convert.ToDouble(Temp));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(Temp));
                    await InsControl._chamber.ChamberStable(Convert.ToDouble(Temp));

                    for (; test_parameter.steadyTime > 0;)
                    {
                        await TaskRecount();
                        progressBar1.Value = test_parameter.steadyTime;
                        label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " 
                        + (test_parameter.steadyTime / 60).ToString() + ":" 
                        + (test_parameter.steadyTime % 60).ToString()));
                    }

                    //STATUS: WAIT
                    //Get N ready for RUN
                    timer = 0;
                    while (true)
                    {
                        if (chamberCtr.CheckAllClientStatus("ready"))
                        {
                            chamberCtr.MasterStatus = "RUN";
                            break;
                        }
                        if (timer > 9600) // wait for 8 Hour at most
                        {
                            chamberCtr.MasterStatus = "RUN";
                            Console.WriteLine("[Master]: Time's UP! Change to RUN status.");
                            break;
                        }
                        timer += 1;
                        System.Threading.Thread.Sleep(3000);
                    }
                    //ate_table[(int)idx].temp = test_parameter.tempList[i];
                    ate_table[(int)idx].ATETask();

                    //STATUS: RUN
                    //Get N over for STOP
                    timer = 0;
                    while (true)
                    {
                        if (chamberCtr.CheckAllClientStatus("idle"))
                        {
                            chamberCtr.MasterStatus = "STOP";
                            break;
                        }
                        if (timer > 9600) // wait for 8 Hour at most
                        {
                            chamberCtr.MasterStatus = "STOP";
                            break;
                        }
                        timer += 1;
                        System.Threading.Thread.Sleep(3000);
                    }
                }
            }
            else if (chamberCtr.Role == "Slave")
            {
                GetTemperature(chamberCtr.SendRequest("temperature"));

                foreach (int Temp in test_parameter.tempList)
                {
                    ate_table[(int)idx].temp = Temp;
                    //status: ready.
                    //sent status and request master status 
                    while (true)
                    {
                        if (chamberCtr.SendRequest("ready") == "RUN")
                            break;
                        System.Threading.Thread.Sleep(3000);
                    }
                    ate_table[(int)idx].ATETask();

                    //status: over.
                    //
                    while (true)
                    {
                        if (chamberCtr.SendRequest("idle") == "STOP")
                            break;
                        System.Threading.Thread.Sleep(3000);
                    }
                }
                chamberCtr.Exit();
            }
            UpdateRunButton();
        }

        private void bt_stop_Click(object sender, EventArgs e)
        {
            test_parameter.run_stop = true;
            bt_run.Enabled = true;
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

        private void bt_pause_Click(object sender, EventArgs e)
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

        private void cb_mode_sel_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cb_mode_sel.SelectedIndex == 0)
            {
                bt_start.Text = "START";
                chamberCtr.Role = "Master";
                button1.Enabled = true;
            }
            else
            {
                bt_start.Text = "Connect";
                chamberCtr.Role = "Slave";
                button1.Enabled = false;
            }
        }

        private void bt_start_Click(object sender, EventArgs e)
        {
            if (cb_mode_sel.SelectedIndex == 0)
            {
                chamberCtr.MasterLisening();
            }
            else
            {
                chamberCtr.ClientConnect(tb_IPAddress.Text);
            }
        }

        private void bt_ipaddress_Click(object sender, EventArgs e)
        {
            IPAddress[] ipa = Dns.GetHostAddresses(Dns.GetHostName());
            tb_IPAddress.Text = ipa[1].ToString();
        }

        private void materialButton1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
        }

        private void CK_I2c_CheckedChanged(object sender, EventArgs e)
        {
            if(CK_I2c.Checked)
            {
                nu_addr.Visible = true;
                nu_code_max.Hexadecimal = true;
                nu_code_min.Hexadecimal = true;
                label29.Visible = true;
            }
            else
            {
                nu_addr.Visible = false;
                nu_code_max.Hexadecimal = false;
                nu_code_min.Hexadecimal = false;
                label29.Visible = false;
            }
        }
    }
}
