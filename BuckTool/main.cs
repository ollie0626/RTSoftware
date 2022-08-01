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
using System.Threading;
using System.Diagnostics;
using System.IO;

namespace BuckTool
{
    public interface ITask
    {
        void ATETask();
    }

    public enum XLS_Table
    {
        A = 1, B, C, D, E, F, G, H, I, J, K, L, M, N, O, P, Q, R, S, T, U, V, W, X, Y, Z,
        AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL, AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX, AY, AZ,
    };


    public partial class main : Sunny.UI.UIForm
    {

        // Thread
        FolderBrowserDialog FolderBrow;
        Thread ATETask;
        ParameterizedThreadStart p_thread;
        public static bool isChamberEn = false;
        int SteadyTime;
        string tempList;

        ATE_Eff _ate_eff = new ATE_Eff();
        ATE_Line _ate_line = new ATE_Line();
        ATE_OutputRipple _ate_ripple = new ATE_OutputRipple();
        ATE_Lx _ate_lx = new ATE_Lx();
        ATE_Loadtrans _ate_trans = new ATE_Loadtrans();

        TaskRun[] ate_table;

        public void GUInit()
        {
            cb_item.SelectedIndex = 0;
            Eload_DG.RowCount = 1;
            ate_table = new TaskRun[] { _ate_eff, _ate_line, _ate_ripple, _ate_lx, _ate_trans };

            led_power.Color = Color.Red;
            led_osc.Color = Color.Red;
            led_eload.Color = Color.Red;
            led_dmm2.Color = Color.Red;
            led_dmm1.Color = Color.Red;
            led_chamber.Color = Color.Red;
            led_37940.Color = Color.Red;

            for (int i = 1; i < 21; i++)
            {
                tb_chamber.Items.Add("ATE_" + i.ToString());
            }
            tb_chamber.SelectedIndex = 0;
        }


        public main()
        {
            InitializeComponent();
            RTBBControl.BoardInit();
            RTBBControl.GpioInit();
            GUInit();
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
                case 5:

                    InsControl._dmm1 = new DMMModule((int)nu_dmm1.Value);
                    if (InsControl._dmm1.InsState())
                        led_dmm1.Color = Color.LightGreen;
                    else
                        led_dmm1.Color = Color.Red;
                    break;

                case 6:
                    InsControl._dmm2 = new DMMModule((int)nu_dmm2.Value);
                    if (InsControl._dmm2.InsState())
                        led_dmm2.Color = Color.LightGreen;
                    else
                        led_dmm2.Color = Color.Red;
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
            Eload_DG.RowCount = Eload_DG.RowCount + 1;
        }

        private void bt_load_sub_Click(object sender, EventArgs e)
        {
            if (Eload_DG.RowCount < 1) return;
            Eload_DG.RowCount = Eload_DG.RowCount - 1;
        }


        private void test_parameter_copy()
        {
            //1.Efficiency / Load Regulation
            //2.Line Regulation
            //3.Output Ripple
            //4.Lx
            //5.Bode
            //6.Load Transient


            switch (cb_item.SelectedIndex)
            {
                case 0:
                    test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.Iout_table = MyLib.DGData(Eload_DG);
                    break;
                case 1:
                    test_parameter.Vin_table = MyLib.TBData(tb_Vin);
                    test_parameter.Iout_table = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    break;
                case 2:
                case 3:
                case 4:
                case 5:
                    test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.Iout_table = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    break;
            }
            
            
            test_parameter.Freq_en[0] = ck_freq1.Checked;
            test_parameter.Freq_en[1] = ck_freq2.Checked;
            test_parameter.Freq_des[0] = tb_freqdes1.Text;
            test_parameter.Freq_des[1] = tb_freqdes2.Text;
            test_parameter.waveform_path = tbWave.Text;
        }


        private void uibt_run_Click(object sender, EventArgs e)
        {
            try
            {
                test_parameter_copy();
                test_parameter.run_stop = false;

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
                    p_thread = new ParameterizedThreadStart(Single_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(cb_item.SelectedIndex);
                }
            }
            catch
            {

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


        private async void MultiChamber_Task()
        {
            //test_parameter.temp_table --> List<string> need to converter to double
            ChamberCtr.IsTCPConnected = false;
            ChamberCtr.ChamberName = tb_chamber.Text;
            ChamberCtr.CreatShareChamberFolder();
            if (!ck_slave.Checked)
            {
                // master
                ChamberCtr.DeleteShareChamberFile();
                ChamberCtr.CreatTempList(tb_templist.Text);
                test_parameter.temp_table = tb_templist.Text.Split(',').ToList();
            }
            else
            {
                // slave
                System.Threading.Thread.Sleep(1000);
                tempList = ChamberCtr.ReadTempList();
                isChamberEn = !string.IsNullOrEmpty(tempList);
                test_parameter.temp_table = tempList.Split(',').ToList();
            }

            ChamberCtr.InitTCPTimer(!ck_slave.Checked);
            ChamberCtr.CurrentStateMaster = "Busy,-999";
            ChamberCtr.CurrenStateSlave = "Busy,-999";
            ChamberCtr.IsTCPNoConnected = !ck_multi_chamber.Checked;
            ChamberCtr.SetTCPTimerState(true);
            Console.WriteLine("StartTime：{0}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));

            for (int i = 0; i < test_parameter.temp_table.Count; i++)
            {
                if (ck_slave.Checked)
                {
                    ChamberCtr.CurrenStateSlave = "Idle," + test_parameter.temp_table[i].ToString();
                }
                else
                {
                    ChamberCtr.CurrentStateMaster = "Busy," + test_parameter.temp_table[i].ToString();
                    SteadyTime = (int)nu_steady.Value;
                    InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                    InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                    bool res = InsControl._chamber.InsState();
                    InsControl._chamber.ChamberOn(Convert.ToDouble(test_parameter.temp_table[i]));
                    InsControl._chamber.ChamberOn(Convert.ToDouble(test_parameter.temp_table[i]));
                    await InsControl._chamber.ChamberStable(Convert.ToDouble(test_parameter.temp_table[i]));

                    for (; SteadyTime > 0;)
                    {
                        await TaskRecount();
                        uiProcessBar1.Value = SteadyTime;
                        label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
                    }
                    ChamberCtr.CurrentStateMaster = "Idle," + test_parameter.temp_table[i].ToString();
                }


                ChamberCtr.CheckTCP_ChamberIdle();
                if (ck_slave.Checked) Console.WriteLine("Slave----------Start Run------------------");
                else Console.WriteLine("Master----------Start Run------------------");
                if (ck_slave.Checked) ChamberCtr.CurrenStateSlave = "Busy," + test_parameter.temp_table[i].ToString();
                else ChamberCtr.CurrentStateMaster = "Busy," + test_parameter.temp_table[i].ToString();

                // ATE test task
                ate_table[cb_item.SelectedIndex].temp = Convert.ToDouble(test_parameter.temp_table[i]);
                ate_table[cb_item.SelectedIndex].ATETask();



                if (ck_slave.Checked) ChamberCtr.CurrenStateSlave = "Idle,9999";
                else ChamberCtr.CurrentStateMaster = "Idle,9999";
                if (ck_slave.Checked) Console.WriteLine("Slave----------WaitFIN------------------");
                else Console.WriteLine("Master----------WaitFIN------------------");
                ChamberCtr.CheckTCP_ChamberIdle();
                if (ck_slave.Checked) Console.WriteLine("Slave----------FIN------------------");
                else Console.WriteLine("Master----------FIN------------------");
                if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
            }

        }

        private async void Chamber_Task()
        {
            test_parameter.temp_table = tb_templist.Text.Split(',').ToList();
            for(int i = 0; i < test_parameter.temp_table.Count; i++)
            {
                if (!Directory.Exists(tbWave.Text + @"\" + tempList[i] + "C"))
                {
                    Directory.CreateDirectory(tbWave.Text + @"\" + tempList[i] + "C");
                }
                test_parameter.waveform_path = tbWave.Text + @"\" + tempList[i] + "C";


                SteadyTime = (int)nu_steady.Value;
                InsControl._chamber = new ChamberModule((int)nu_chamber.Value);
                InsControl._chamber.ConnectChamber((int)nu_chamber.Value);
                InsControl._chamber.ChamberOn(Convert.ToDouble(test_parameter.temp_table[i]));
                InsControl._chamber.ChamberOn(Convert.ToDouble(test_parameter.temp_table[i]));
                await InsControl._chamber.ChamberStable(Convert.ToDouble(test_parameter.temp_table[i]));
                for (; SteadyTime > 0;)
                {
                    await TaskRecount();
                    uiProcessBar1.Value = SteadyTime;
                    label1.Invoke((MethodInvoker)(() => label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString()));
                    //label1.Text = "count down: " + (SteadyTime / 60).ToString() + ":" + (SteadyTime % 60).ToString();
                }


                ate_table[cb_item.SelectedIndex].temp = Convert.ToDouble(test_parameter.temp_table[i]);
                ate_table[cb_item.SelectedIndex].ATETask();


            }
            // test finish chamber to 25C
            if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
        }

        private void Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
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

    }
}
