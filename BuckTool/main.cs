using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;

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
        string App_name = "Buck Tool v1.3";

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
            led_funcgen.Color = Color.Red;

            for (int i = 1; i < 21; i++)
            {
                cb_chamber.Items.Add("ATE_" + i.ToString());
            }
            cb_chamber.SelectedIndex = 0;
            this.Text = App_name;
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

                case 7:
                    InsControl._funcgen = new FuncGenModule((int)nu_funcgen.Value);
                    if (InsControl._funcgen.InsState())
                        led_funcgen.Color = Color.LightGreen;
                    else
                        led_funcgen.Color = Color.Red;
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
            //5.Load Transient

            switch (cb_item.SelectedIndex)
            {
                case 0: // eff and load regulation
                case 2: // output ripple
                    test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.Iout_table = MyLib.DGData(Eload_DG);
                    break;
                case 1: // line regulation
                    // start, stop, step
                    test_parameter.Vin_table = MyLib.TBData(tb_lineVin);
                    test_parameter.Iout_table = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    break;
                case 3: // Lx 
                    test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.Iout_table = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    break;
                case 4:
                    test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.HiLevel = tb_Highlevel.Text.Split(',').Select(double.Parse).ToList();
                    test_parameter.LoLevel = tb_Lowlevel.Text.Split(',').Select(double.Parse).ToList();
                    //test_parameter.HiLo_table.Add()
                    Hi_Lo level = new Hi_Lo();

                    for(int hi_index = 0; hi_index < test_parameter.HiLevel.Count; hi_index++)
                    {
                        for(int lo_index = 0; lo_index < test_parameter.LoLevel.Count; lo_index++)
                        {
                            level.Highlevel = test_parameter.HiLevel[hi_index];
                            level.LowLevel = test_parameter.LoLevel[lo_index];
                            test_parameter.HiLo_table.Add(level);
                        }
                    }
                    break;
            }
            
            test_parameter.Freq_en[0] = ck_freq1.Checked;
            test_parameter.Freq_en[1] = ck_freq2.Checked;
            test_parameter.Freq_des[0] = tb_freqdes1.Text;
            test_parameter.Freq_des[1] = tb_freqdes2.Text;
            test_parameter.waveform_path = tbWave.Text;
            test_parameter.freq = (double)nu_Freq.Value;
            test_parameter.duty = (double)nu_duty.Value;
            test_parameter.tr = (double)nu_tr.Value;
            test_parameter.tf = (double)nu_tf.Value;
            test_parameter.vout_ideal = (double)nu_Videa.Value; 

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
            ChamberCtr.ChamberName = cb_chamber.Text;
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
                if(cb_item.SelectedIndex == 5)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        ate_table[j].temp = Convert.ToDouble(test_parameter.temp_table[i]);
                        ate_table[j].ATETask();
                    }
                }
                else
                {
                    ate_table[cb_item.SelectedIndex].temp = Convert.ToDouble(test_parameter.temp_table[i]);
                    ate_table[cb_item.SelectedIndex].ATETask();
                }    

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

                if(cb_item.SelectedIndex == 5)
                {
                    for (int j = 0; j < 2; j++)
                    {
                        ate_table[j].temp = 25;
                        ate_table[j].ATETask();
                    }
                }
                else
                {
                    ate_table[cb_item.SelectedIndex].temp = Convert.ToDouble(test_parameter.temp_table[i]);
                    ate_table[cb_item.SelectedIndex].ATETask();
                }
            }
            // test finish chamber to 25C
            if (InsControl._chamber != null) InsControl._chamber.ChamberOn(25);
        }

        private void Single_Task(object idx)
        {
            if((int)idx == 5)
            {

                for(int i = 0; i < 2; i++)
                {
                    test_parameter.Vin_table.Clear();
                    test_parameter.Iout_table.Clear();
                    if (i == 0)
                    {
                        test_parameter.Vin_table = tb_Vin.Text.Split(',').Select(double.Parse).ToList();
                        test_parameter.Iout_table = MyLib.DGData(Eload_DG);
                    }
                    else if (i == 1)
                    {
                        test_parameter.Vin_table = MyLib.TBData(tb_lineVin);
                        test_parameter.Iout_table = tb_Iout.Text.Split(',').Select(double.Parse).ToList();
                    }


                    ate_table[i].temp = 25;
                    ate_table[i].ATETask();
                }
            }
            else
            {
                ate_table[(int)idx].temp = 25;
                ate_table[(int)idx].ATETask();
            }
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

        private void uibt_save_Click(object sender, EventArgs e)
        {
            SaveFileDialog savedlg = new SaveFileDialog();
            savedlg.Filter = "settings|*.tb_info";
            
            if(savedlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = savedlg.FileName;
                SaveSettings(file_name);
            }

        }


        private void SaveSettings(string file)
        {
            string settings = "";

            settings = "0.BinPath=$" + textBox1.Text + "$\r\n";
            settings += "1.WavePath=$" + tbWave.Text + "$\r\n";
            settings += "2.SpecifyPath=$" + textBox2.Text + "$\r\n";
            settings += "3.Vin=$" + tb_Vin.Text + "$\r\n";
            settings += "4.Freq1_en=$" + (ck_freq1.Checked ? "1" : "0") + "$\r\n";
            settings += "5.Freq1_des=$" + tb_freqdes1.Text + "$\r\n";
            settings += "6.Freq2_en=$" + (ck_freq2.Checked ? "1" : "0") + "$\r\n";
            settings += "7.Freq2_des=$" + tb_freqdes2.Text + "$\r\n";
            settings += "8.Func_freq=$" + nu_Freq.Value.ToString() + "$\r\n";
            settings += "9.Func_duty=$" + nu_duty.Value.ToString() + "$\r\n";
            settings += "10.Func_tr=$" + nu_tr.Value.ToString() + "$\r\n";
            settings += "11.Func_tf=$" + nu_tf.Value.ToString() + "$\r\n";
            settings += "12.Func_hi_level=$" + tb_Highlevel.Text + "$\r\n";
            settings += "13.Func_lo_level=$" + tb_Lowlevel.Text + "$\r\n";
            settings += "14.Iout_non_Seq=$" + tb_Iout.Text + "$\r\n";
            /* connect ins. info */
            settings += "15.Scope_addr=$" + tb_osc.Text + "$\r\n";
            settings += "16.Power_addr=$" + nu_power.Value.ToString() + "$\r\n";
            settings += "17.Eload_addr=$" + nu_eload.Value.ToString() + "$\r\n";
            settings += "18.34970_adr=$" + nu_34970A.Value.ToString() + "$\r\n";
            settings += "19.Chamber_addr=$" + nu_chamber.Value.ToString() + "$\r\n";
            settings += "20.Dmm1_addr=$" + nu_dmm1.Value.ToString() + "$\r\n";
            settings += "21.Dmm2_addr=$" + nu_dmm2.Value.ToString() + "$\r\n";
            settings += "22.Func_addr=$" + nu_funcgen.Value.ToString() + "$\r\n";

            /* chamber info */
            //settings += "23.Chamber_en=$" + (ck_chaber_en.Checked ? "1" : "0") + "$\r\n";
            settings += "23.Chamber_info=$" + tb_templist.Text + "$\r\n";
            settings += "24.Chamber_name=$" + cb_chamber.SelectedIndex.ToString() + "$\r\n";
            settings += "25.Chamber_time=$" + nu_steady.Value.ToString() + "$\r\n";
            settings += "26.Eload_row=$" + Eload_DG.RowCount.ToString() + "$\r\n";
            for(int idx = 0; idx < Eload_DG.RowCount; idx++)
            {
                settings += (idx + 27).ToString() + ".Eload_start=$" + Eload_DG[0, idx].Value.ToString() + "$\r\n";
                settings += (idx + 28).ToString() + ".Eload_step=$" + Eload_DG[1, idx].Value.ToString() + "$\r\n";
                settings += (idx + 29).ToString() + ".Eload_stop=$" + Eload_DG[2, idx].Value.ToString() + "$\r\n";
            }

            using (StreamWriter sw = new StreamWriter(file))
            {
                sw.Write(settings);
            }
        }

        private void uibt_load_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "settings|*.tb_info";
            if(opendlg.ShowDialog() == DialogResult.OK)
            {
                LoadSettings(opendlg.FileName);
            }
        }

        private void LoadSettings(string file)
        {
            object[] obj_arr = new object[]
            {
                textBox1, tbWave, textBox2, tb_Vin, ck_freq1, tb_freqdes1, ck_freq2, tb_freqdes2, nu_Freq, nu_duty, nu_tr,
                nu_tf, tb_Highlevel, tb_Lowlevel, tb_Iout, tb_osc, nu_power, nu_eload, nu_34970A, nu_chamber, nu_dmm1,
                nu_dmm2, nu_funcgen, tb_templist, cb_chamber, nu_steady, Eload_DG
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
                for(int i = 0; i < obj_arr.Length; i++)
                {
                    switch(obj_arr[i].GetType().Name)
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
                            idx = i;
                            goto fullDG;
                            
                            break;
                    }
                }

                fullDG:
                for(int i = 0; i < Eload_DG.RowCount; i++)
                {
                    Eload_DG[0, i].Value = Convert.ToString(info[idx + 1]); // start
                    Eload_DG[1, i].Value = Convert.ToString(info[idx + 2]); // step
                    Eload_DG[2, i].Value = Convert.ToString(info[idx + 3]); // stop
                    idx += 3;
                }
                

            }
        }
    }
}
