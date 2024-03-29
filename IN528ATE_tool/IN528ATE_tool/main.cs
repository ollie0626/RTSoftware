﻿using System;
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
using System.Net.Sockets;
using System.Net;
// this's new basic branch

using System.Text.RegularExpressions;

namespace IN528ATE_tool
{
    public partial class main : Sunny.UI.UIForm
    {
        delegate void MyDelegate();
        MyDelegate Message;
        FolderBrowserDialog FolderBrow;
        ChamberCtr chamberCtr = new ChamberCtr();

        RTBBControl RTDev;
        MyLib myLib;
        int SteadyTime;

        public static bool isChamberEn = false;

        ParameterizedThreadStart p_thread;
        Thread ATETask;

        ATE_OutputRipple _ate_ripple;
        ATE_CodeInrush _ate_code_inrush;
        ATE_PowerOn _ate_poweron;
        ATE_CurrentLimit _ate_current_limit;
        ATE_UVPLevel _ate_uvp;
        ATE_UVPDly _ate_dly;
        TaskRun[] ate_table;

        string[] tempList;
        string templist;
        int item_sel;

        //private ChamberLink chamberlink = new ChamberLink();

        private void GUIInit()
        {
            /* class init */
            this.Text = "ATE Tool v3.45";
            RTDev = new RTBBControl();
            myLib = new MyLib();

            _ate_ripple = new ATE_OutputRipple();
            _ate_code_inrush = new ATE_CodeInrush();
            _ate_poweron = new ATE_PowerOn();
            _ate_current_limit = new ATE_CurrentLimit();
            _ate_uvp = new ATE_UVPLevel();
            _ate_dly = new ATE_UVPDly();

            led_osc.Color = Color.Red;
            led_power.Color = Color.Red;
            led_eload.Color = Color.Red;
            led_37940.Color = Color.Red;
            led_chamber.Color = Color.Red;
            cb_item.SelectedIndex = 0;
            ate_table = new TaskRun[] { _ate_ripple, _ate_code_inrush, _ate_poweron, _ate_current_limit, _ate_uvp, _ate_dly };
            Message = new MyDelegate(MessageCallback);


            for(int i = 1; i < 21; i++)
            {
                tb_chamber.Items.Add("ATE_" + i.ToString());
            }

            tb_chamber.SelectedIndex = 0;
            Console.WriteLine(tb_chamber.Text);
            test_parameter.run_stop = false;
            test_parameter.chamber_en = false;
            cb_mode_sel.SelectedIndex = 0;

            tabControl1.SelectedIndex = 1;
        }

        public main()
        {
            InitializeComponent();
            GUIInit();
        }

        private void MessageCallback()
        {
            MessageBox.Show("Callback message test!!");
        }


        private void connect_Ins(int idx)
        {
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

        private void test_parameter_copy()
        {
            //nu_ch2_level
            test_parameter.ch2_level = (double)nu_ch2_level.Value;
            test_parameter.ch2_user_define = ck_ch2_user_define.Checked;


            test_parameter.chamber_en = ck_chaber_en.Checked;
            test_parameter.run_stop = false;
            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.IoutList = tb_ioutList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.specify_id = (byte)nu_specify.Value;
            test_parameter.slave = (byte)nu_slave.Value;
            test_parameter.binFolder = textBox1.Text;
            test_parameter.specify_bin = textBox2.Text;
            test_parameter.waveform_path = tbWave.Text;
            test_parameter.ontime_scale_ms = (double)nu_ontime_scale.Value;
            test_parameter.offtime_scale_ms = (double)nu_offtime_scale.Value;
            test_parameter.addr = (byte)nu_addr.Value;
            test_parameter.max = (byte)nu_code_max.Value;
            test_parameter.min = (byte)nu_code_min.Value;
            test_parameter.vol_max = (double)nu_vol_max.Value;
            test_parameter.vol_min = (double)nu_vol_min.Value;
            test_parameter.all_en = ck_all_test.Checked;
            test_parameter.trigger_vin_en = ck_vin_trigger.Checked;
            test_parameter.trigger_en = ck_en_trigger.Checked;
            test_parameter.trigger_level = (double)nu_ch1_trigger_level.Value;
            test_parameter.mtp_slave = (byte)nu_mtp_slave.Value;
            test_parameter.mtp_addr = (byte)nu_mtp_addr.Value;
            test_parameter.mtp_data = (byte)nu_mtp_data.Value;
            test_parameter.measure_level = (double)nu_measure_level.Value;
            test_parameter.mtp_enable = CK_Program.Checked;
            test_parameter.cv_setting = (double)nu_CVSetting.Value;
            test_parameter.cv_step = (double)nu_CVStep.Value;
            test_parameter.cv_wait = (double)nu_CVwait.Value;

            test_parameter.ripple_time_manual = check_ripple_time.Checked;


            test_parameter.lovol = (double)nu_LoVol.Value;
            test_parameter.midvol = (double)nu_MidVol.Value;
            test_parameter.hivol = (double)nu_HiVol.Value;

            // swire
            test_parameter.swireList.Clear();
            for (int i = 0; i < swireTable.RowCount; i++)
            {
                test_parameter.swireList.Add((string)swireTable[0, i].Value);
                //test_parameter.voutList.Add(Convert.ToDouble(swireTable[1, i].Value));
            }
            test_parameter.swire_en = ck_swire.Checked;
            test_parameter.swire_20 = RB20.Checked;
            test_parameter.bw_en = CK_Bwen.Checked;

            test_parameter.dt_rising_en = !ck_negvol.Checked;
            test_parameter.sst_sel = ck_sst_define.Checked ? 1 : 0;

            test_parameter.lovout = (double)nu_loVout.Value;
            test_parameter.midvout = (double)nu_midVout.Value;
            test_parameter.hivout = (double)nu_hiVout.Value;
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
                test_parameter_copy();
                item_sel = cb_item.SelectedIndex;

                //ChamberCtr.ChamberName = tb_chamber.Text;
                SteadyTime = (int)nu_steady.Value;

                if (ck_multi_chamber.Checked && ck_chaber_en.Checked)
                {
                    //ATETask = new Thread(MultiChamber_Task);
                    p_thread = new ParameterizedThreadStart(multi_ate_process);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(cb_item.SelectedIndex);
                }
                else if (ck_chaber_en.Checked)
                {
                    ATETask = new Thread(Chamber_Task);
                    ATETask.Start();
                }
                else
                {
                    // single no chamber conditions

                    for(int i = 0; i < ate_table.Length; i++)
                    {
                        ate_table[i].temp = 25;
                    }

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

        private List<int> TemperatureList = new List<int>();

        private void GetTemperature(string input)
        {
            TemperatureList.Clear();
            string[] temp = input.Split(',');
            foreach (string str in temp)
            {
                TemperatureList.Add(Convert.ToInt32(str));
            }
        }

        private async void multi_ate_process(object idx)
        {
            
            int timer;
            MyLib myLib = new MyLib();
            //myLib.time = (int)nu_steady.Value;
            chamberCtr.Init(tb_templist.Text);

            if (chamberCtr.Role == "Master")
            {
                GetTemperature(tb_templist.Text);
                chamberCtr.Dispose();

                foreach (int Temp in TemperatureList)
                {
                    ate_table[(int)idx].temp = Temp;

                    Console.WriteLine("StartTime：{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                    SteadyTime = (int)nu_steady.Value;
                    InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                    InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                    bool res = InsControl._chamber.InsState();
                    InsControl._chamber.ChamberOn(Convert.ToDouble(Temp));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(Temp));
                    await InsControl._chamber.ChamberStable(Convert.ToDouble(Temp));

                    for (; SteadyTime > 0;)
                    {
                        await TaskRecount();
                        uiProcessBar1.Value = SteadyTime;
                        label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
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

                foreach (int Temp in TemperatureList)
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
        }

        // ATE Process: MultiChamber_Task, Chamber_Task, Run_Task_Flow, Run_Single_Task

        //public async void MultiChamber_Task()
        //{
        //    ChamberCtr.IsTCPConnected = false;
        //    ChamberCtr.ChamberName = tb_chamber.Text;
        //    ChamberCtr.CreatShareChamberFolder();

        //    if (!ck_slave.Checked)
        //    {
        //        // master
        //        ChamberCtr.DeleteShareChamberFile();
        //        ChamberCtr.CreatTempList(templist);
        //    }
        //    else
        //    {
        //        // slave
        //        System.Threading.Thread.Sleep(1000);
        //        templist = ChamberCtr.ReadTempList();
        //        isChamberEn = !string.IsNullOrEmpty(templist);
        //        tempList = templist.Split(',');
        //    }

        //    ChamberCtr.InitTCPTimer(!ck_slave.Checked);
        //    ChamberCtr.CurrentStateMaster = "Busy,-999";
        //    ChamberCtr.CurrenStateSlave = "Busy,-999";
        //    ChamberCtr.IsTCPNoConnected = !ck_multi_chamber.Checked;
        //    ChamberCtr.SetTCPTimerState(true);
        //    Console.WriteLine("StartTime：{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

        //    //if (!await TaskConnect(300)) return;// connect

        //    for (int i = 0; i < tempList.Length; i++)
        //    {
        //        if(ck_slave.Checked)
        //        {
        //            ChamberCtr.CurrenStateSlave = "Idle," + tempList[i].ToString();
        //        }
        //        else
        //        {
        //            ChamberCtr.CurrentStateMaster = "Busy," + tempList[i].ToString();
        //            SteadyTime = (int)nu_steady.Value;
        //            InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
        //            InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
        //            bool res = InsControl._chamber.InsState();
        //            InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
        //            InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
        //            await InsControl._chamber.ChamberStable(Convert.ToDouble(tempList[i]));

        //            for (; SteadyTime > 0;)
        //            {
        //                await TaskRecount();
        //                uiProcessBar1.Value = SteadyTime;
        //                label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
        //            }
        //            ChamberCtr.CurrentStateMaster = "Idle," + tempList[i].ToString();
        //        }

        //        ChamberCtr.CheckTCP_ChamberIdle();
        //        if (ck_slave.Checked) Console.WriteLine("Slave----------Start Run------------------");
        //        else Console.WriteLine("Master----------Start Run------------------");
        //        if (ck_slave.Checked) ChamberCtr.CurrenStateSlave = "Busy," + tempList[i].ToString();
        //        else ChamberCtr.CurrentStateMaster = "Busy," + tempList[i].ToString();

        //        // ripple test
        //        _ate_ripple.temp = Convert.ToDouble(tempList[i]);
        //        _ate_code_inrush.temp = Convert.ToDouble(tempList[i]);
        //        _ate_poweron.temp = Convert.ToDouble(tempList[i]);
        //        _ate_current_limit.temp = Convert.ToDouble(tempList[i]);

        //        //_ate_current_limit.temp = 25;
        //        _ate_uvp.temp = Convert.ToDouble(tempList[i]);
        //        _ate_dly.temp = Convert.ToDouble(tempList[i]);

        //        if (!test_parameter.all_en)
        //        {
        //            switch (item_sel)
        //            {
        //                case 0:
        //                    _ate_ripple.ATETask();
        //                    break;
        //                case 1:
        //                    _ate_code_inrush.ATETask();
        //                    break;
        //                case 2:
        //                    _ate_poweron.ATETask();
        //                    break;
        //                case 3:
        //                    _ate_current_limit.ATETask();
        //                    break;
        //            }
        //        }
        //        else
        //        {
        //            _ate_ripple.ATETask();
        //            _ate_code_inrush.ATETask();
        //        }

        //        if (ck_slave.Checked) ChamberCtr.CurrenStateSlave = "Idle,9999";
        //        else ChamberCtr.CurrentStateMaster = "Idle,9999";
        //        if (ck_slave.Checked) Console.WriteLine("Slave----------WaitFIN------------------");
        //        else Console.WriteLine("Master----------WaitFIN------------------");
        //        ChamberCtr.CheckTCP_ChamberIdle();
        //        if (ck_slave.Checked) Console.WriteLine("Slave----------FIN------------------");
        //        else Console.WriteLine("Master----------FIN------------------");
        //        if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
        //    }
        //}

        public async void Chamber_Task()
        {
            try
            {
                for (int i = 0; i < tempList.Length; i++)
                {
                    if (!Directory.Exists(tbWave.Text +tempList[i] + "C"))
                    {
                        Directory.CreateDirectory(tbWave.Text  + tempList[i] + "C");
                    }
                    test_parameter.waveform_path = tbWave.Text + tempList[i] + "C";

                    SteadyTime = (int)nu_steady.Value;
                    InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                    InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(tempList[i]));
                    //await InsControl._chamber.ChamberStable(Convert.ToDouble(tempList[i]));
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
                    _ate_current_limit.temp = Convert.ToDouble(tempList[i]);

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
                            case 3:
                                _ate_current_limit.ATETask();
                                break;
                        }
                    }
                    else
                    {
                        _ate_ripple.ATETask();
                        _ate_code_inrush.ATETask();
                    }

                }
                if (InsControl._chamber != null)
                {
                    InsControl._chamber.ChamberOn(25);
                    InsControl._chamber.ChamberStableCheck(25);
                    InsControl._chamber.ChamberOff();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.StackTrace, "ATE Tool", System.Windows.Forms.MessageBoxButtons.OK);
            }
            finally
            {
                if (InsControl._chamber != null)
                {
                    InsControl._chamber.ChamberOn(25);
                    InsControl._chamber.ChamberStableCheck(25);
                    InsControl._chamber.ChamberOff();
                }
            }
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

            /*
                0. Output Ripple
                1. Code Inrush
                2. Delay Time & SST
                3. Current Limit
                4. UVP
             */

            switch (cb_item.SelectedIndex)
            {
                case 0:
                    // Output ripple
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vout \r\n";
                    break;
                case 1:
                    // code inrush
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vout \r\n" + 
                                     "Ch2: Iout";
                    break;
                case 2:
                    // delay time
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vin or Enable \r\n" +
                                     "Ch2: Vout";
                    break;
                case 3:
                    // current
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vout \r\n" +
                                     "Ch2: Lx \r\n" +
                                     "CH3: ILX";
                    break;
                case 4:
                    // UVP
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vout \r\n" +
                                     "Ch2: Lx \r\n" +
                                     "CH3: ILX";
                    break;
                case 5:
                    // UVP Delay time
                    lab_scope.Text = "Scope Info:" + "\r\n" +
                                     "CH1: Vout \r\n" +
                                     "Ch2: Other Lx \r\n" +
                                     "CH3: ILX";
                    break;
            }



        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            IN528ATE_tool.Properties.Settings.Default.binpath = this.textBox1.Text;
            IN528ATE_tool.Properties.Settings.Default.specifypath = this.textBox2.Text;
            IN528ATE_tool.Properties.Settings.Default.wavepath = this.tbWave.Text;
            Properties.Settings.Default.vinList = tb_vinList.Text;
            Properties.Settings.Default.IoutList = tb_ioutList.Text;
            Properties.Settings.Default.itemSel = cb_item.SelectedIndex;
            Properties.Settings.Default.ontime = nu_ontime_scale.Value;
            Properties.Settings.Default.offtime = nu_offtime_scale.Value;
            Properties.Settings.Default.mtp_slave = nu_mtp_slave.Value;
            Properties.Settings.Default.mtp_addr = nu_mtp_addr.Value;
            Properties.Settings.Default.mtp_data = nu_mtp_data.Value;
            Properties.Settings.Default.mtp_en = CK_Program.Checked;
            Properties.Settings.Default.slave = nu_slave.Value;
            Properties.Settings.Default.sp_slave = nu_slave.Value;
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
            nu_ontime_scale.Value = Properties.Settings.Default.ontime;
            nu_offtime_scale.Value = Properties.Settings.Default.offtime;

            nu_mtp_slave.Value = Properties.Settings.Default.mtp_slave;
            nu_mtp_addr.Value = Properties.Settings.Default.mtp_addr;
            nu_mtp_data.Value = Properties.Settings.Default.mtp_data;
            CK_Program.Checked = Properties.Settings.Default.mtp_en;

            nu_slave.Value = Properties.Settings.Default.slave;
            nu_specify.Value = Properties.Settings.Default.sp_slave;
#if false
            connect_Ins(0);
            connect_Ins(1);
            connect_Ins(2);
            connect_Ins(3);
#endif
        }

        private void uiSymbolButton2_Click(object sender, EventArgs e)
        {
            //RTDev.BoadInit();
            //RTDev.I2C_WriteBin(0x46 >> 1, 0x00, textBox2.Text);

            InsControl._eload.SetCV_Vol(5.3);
        }

        private void SwireRow_ValueChanged(object sender, EventArgs e)
        {
            swireTable.ColumnCount = 1;
            swireTable.Columns[0].HeaderText = "swire";
            swireTable.Columns[0].Width = 150;
            //swireTable.Columns[1].HeaderText = "vout";
            swireTable.RowCount = (int)SwireRow.Value;
        }

        private void bt_SwireSave_Click(object sender, EventArgs e)
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
            if(swireTable.RowCount != 0)
            {
                string settings = "";
                for(int cnt = 0; cnt < swireTable.RowCount; cnt++)
                {
                    settings += string.Format("{0}.Row=${1}$\r\n",
                        cnt,
                        swireTable[0, cnt].Value.ToString());
                }
                using (StreamWriter sw = new StreamWriter(file))
                {
                    sw.Write(settings);
                }
            }
        }

        private void bt_SwireLoad_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "settings|*.tb_info";
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                LoadSettings(openFileDialog1.FileName);
            }
        }

        private void LoadSettings(string file)
        {
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
                SwireRow.Value = info.Count;

                for (int i = 0; i < info.Count; i++)
                {
                    string buf = info[i];

                    swireTable[0, i].Value = buf;
                    //swireTable[1, i].Value = buf[1];
                }
            }
        }

        private void cb_mode_sel_SelectedIndexChanged(object sender, EventArgs e)
        {
            if(cb_mode_sel.SelectedIndex == 0)
            {
                bt_start.Text = "START";
                chamberCtr.Role = "Master";
                bt_stop.Enabled = true;
            }
            else
            {
                bt_start.Text = "Connect";
                chamberCtr.Role = "Slave";
                bt_stop.Enabled = false;
            }
            
        }

        private void bt_ipaddress_Click(object sender, EventArgs e)
        {
            IPAddress[] ipa = Dns.GetHostAddresses(Dns.GetHostName());
            tb_IPAddress.Text = ipa[1].ToString();
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

        private void bt_stop_Click(object sender, EventArgs e)
        {

            if(cb_mode_sel.SelectedIndex == 0)
            {
                foreach (var c in ChamberCtr.Socket_List)
                {
                    c.Shutdown(SocketShutdown.Both);
                    c.Close();
                }
                //serverSocket.Close();
                chamberCtr.ClientNowStatus.Clear();
                ChamberCtr.Socket_List.Clear();
            }
        }

        private void saveSettingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            savedlg.Filter = "settings|*.tb_info";

            if (savedlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = savedlg.FileName;
                SaveParameter(file_name);
            }
        }

        private void SaveParameter(string file)
        {
            string settings = "";
            string finish_symbol = "$\r\n";
            settings = "0.bin_path=$" + textBox1.Text + finish_symbol;
            settings += "1.WavePath=$" + tbWave.Text + finish_symbol;
            settings += "2.Specify=$" + textBox2.Text + finish_symbol;

            settings += "3.TempList=$" + tb_templist.Text + finish_symbol;
            settings += "4.VinList=$" + tb_vinList.Text + finish_symbol;
            settings += "5.IoutList=$" + tb_ioutList.Text + finish_symbol;
            settings += "6.IPAddress=$" + tb_IPAddress.Text + finish_symbol;
            settings += "7.Slave=$" + nu_specify.Value + finish_symbol;
            settings += "8.Speicy_Slave=$" + nu_slave.Value + finish_symbol;
            settings += "9.On_time=$" + nu_ontime_scale.Value + finish_symbol;
            settings += "10.Off_time=$" + nu_offtime_scale.Value + finish_symbol;
            settings += "11.MeasLevel=$" + nu_measure_level.Value + finish_symbol;
            settings += "12.TriggerLeve=$" + nu_ch1_trigger_level.Value + finish_symbol;
            settings += "13.Address=$" + nu_addr.Value + finish_symbol;
            settings += "14.Min=$" + nu_code_min.Value + finish_symbol;
            settings += "15.Max=$" + nu_code_max.Value + finish_symbol;
            settings += "16.Max=$" + nu_vol_min.Value + finish_symbol;
            settings += "nu_vol_max=$" + nu_vol_max.Value + finish_symbol;

            settings += "17.MTP_slave=$" + nu_mtp_slave.Value + finish_symbol;
            settings += "18.MTP_addr=$" + nu_mtp_addr.Value + finish_symbol;
            settings += "19.MTP_data=$" + nu_mtp_data.Value + finish_symbol;

            settings += "20.Steady=$" + nu_steady.Value + finish_symbol;
            settings += "21.CVsetting=$" + nu_CVSetting.Value + finish_symbol;
            settings += "22.CVsetp=$" + nu_CVStep.Value + finish_symbol;
            settings += "23.CVWait=$" + nu_CVwait.Value + finish_symbol;
            settings += "24.EloadCh2=$" + nu_eloadch2.Value + finish_symbol;

            settings += "25.item=$" + cb_item.SelectedIndex + finish_symbol;
            settings += "26.MasterSel=$" + cb_mode_sel.SelectedIndex + finish_symbol;
            settings += "27.VinTrigger=$" + (ck_vin_trigger.Checked ? "1" : "0") + finish_symbol;
            settings += "28.EnTrigger=$" + (ck_en_trigger.Checked ? "1" : "0") + finish_symbol;
            settings += "29.MTKEn=$" + (CK_Program.Checked ? "1" : "0") + finish_symbol;
            settings += "30.ChamberEn=$" + (ck_chaber_en.Checked ? "1" : "0") + finish_symbol;
            settings += "31.Multi=$" + (ck_multi_chamber.Checked ? "1" : "0") + finish_symbol;

            using (StreamWriter sw = new StreamWriter(file))
            {
                sw.Write(settings);
            }
        }

        private void loadSettingToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "settings|*.tb_info";
            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                LoadParameter(opendlg.FileName);
            }
        }

        private void LoadParameter(string file)
        {
            object[] obj_arr = new object[]
            {
                textBox1, tbWave, textBox2, tb_templist, tb_vinList, tb_ioutList, tb_IPAddress, nu_specify, nu_slave, nu_ontime_scale, nu_offtime_scale,
                nu_measure_level, nu_ch1_trigger_level, nu_addr, nu_code_min, nu_code_max, nu_vol_min, nu_vol_max,
                nu_mtp_slave, nu_mtp_addr, nu_mtp_data, nu_steady, nu_CVSetting, nu_CVStep, nu_CVwait, nu_eloadch2, cb_item,
                ck_vin_trigger, ck_en_trigger, CK_Program, ck_chaber_en, ck_multi_chamber

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
                            break;
                    }
                }


            }
        }
    }
}
