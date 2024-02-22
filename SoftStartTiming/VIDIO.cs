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
using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;

namespace SoftStartTiming
{
    public partial class VIDIO : Form
    {
        string win_name = "VIDIO v1.19";
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        TaskRun[] ate_table;
        string[] tempList;
        int SteadyTime;

        Dictionary<string, string> Device_map = new Dictionary<string, string>();


        ATE_VIDIO _ate_vid_io;
        ToolStripMenuItem[] _item = new ToolStripMenuItem[2];
        DataGridView DV = new DataGridView();

        RTBBControl RTDev = new RTBBControl();


        private void InitDG()
        {
            // lpm
            dataGridViewComboBoxColumn1.Items.Add("0");
            dataGridViewComboBoxColumn1.Items.Add("1");

            // g1
            dataGridViewComboBoxColumn2.Items.Add("0");
            dataGridViewComboBoxColumn2.Items.Add("1");

            // g2
            dataGridViewComboBoxColumn3.Items.Add("0");
            dataGridViewComboBoxColumn3.Items.Add("1");

            dataGridView2.RowCount = 8;

            for (int idx = 0; idx < 8; idx++)
            {
                switch (idx)
                {
                    case 0:
                        dataGridView2[0, 0].Value = "0";
                        dataGridView2[1, 0].Value = "0";
                        dataGridView2[2, 0].Value = "0";
                        break;
                    case 1:
                        dataGridView2[0, 1].Value = "0";
                        dataGridView2[1, 1].Value = "0";
                        dataGridView2[2, 1].Value = "1";
                        break;
                    case 2:
                        dataGridView2[0, 2].Value = "0";
                        dataGridView2[1, 2].Value = "1";
                        dataGridView2[2, 2].Value = "0";
                        break;
                    case 3:
                        dataGridView2[0, 3].Value = "0";
                        dataGridView2[1, 3].Value = "1";
                        dataGridView2[2, 3].Value = "1";
                        break;
                    case 4:
                        dataGridView2[0, 4].Value = "1";
                        dataGridView2[1, 4].Value = "0";
                        dataGridView2[2, 4].Value = "0";
                        break;
                    case 5:
                        dataGridView2[0, 5].Value = "1";
                        dataGridView2[1, 5].Value = "0";
                        dataGridView2[2, 5].Value = "1";
                        break;
                    case 6:
                        dataGridView2[0, 6].Value = "1";
                        dataGridView2[1, 6].Value = "1";
                        dataGridView2[2, 6].Value = "0";
                        break;
                    case 7:
                        dataGridView2[0, 7].Value = "1";
                        dataGridView2[1, 7].Value = "1";
                        dataGridView2[2, 7].Value = "1";
                        break;
                }
            }



        }

        public VIDIO()
        {
            InitializeComponent();
            InitDG();

            this.Text = win_name;

            _ate_vid_io = new ATE_VIDIO(this);
            ate_table = new TaskRun[] { _ate_vid_io };

            _item[0] = new ToolStripMenuItem("Delete Row", null, new EventHandler(ToolStripDelete_Click));
            _item[1] = new ToolStripMenuItem("Add Row", null, new EventHandler(ToolStripAdd_Click));

            cbox_scope_ch2.SelectedIndex = 0;
            num_ILXLevel.Value = (decimal)5;
            RTDev.BoadInit();
        }

        private void ToolStripDelete_Click(object sender, EventArgs e)
        {
            DV.Rows.RemoveAt(DV.CurrentRow.Index);
            if (DV.TabIndex == 5)
            {
                for (int i = 0; i <= DV.RowCount - 1; i++)
                    DV[7, i].Value = i + 2;
            }
        }

        private void ToolStripAdd_Click(object sender, EventArgs e)
        {
            DV.RowCount = DV.RowCount + 1;
            if (DV.TabIndex == 5)
            {
                for (int i = 0; i <= DV.RowCount - 1; i++)
                    DV[7, i].Value = i + 2;
            }
        }

        private void BT_Add_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = dataGridView1.RowCount + 1;
        }

        private void BT_Sub_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount - 1 > 0)
                dataGridView1.RowCount = dataGridView1.RowCount - 1;
            else
                dataGridView1.RowCount = 0;
        }

        private void test_parameter_copy()
        {
            test_parameter.vidio.vout_list.Clear();
            test_parameter.vidio.vout_list_af.Clear();
            test_parameter.vidio.vout_map.Clear();
            test_parameter.vidio.io_table.Clear();
            test_parameter.vidio.lpm_vout_map.Clear();
            test_parameter.vidio.lpm_str.Clear();
            test_parameter.vidio.test_cnt = (int)nu_test_cnt.Value;

            test_parameter.tool_ver = win_name + "\r\n";
            test_parameter.vin_conditions = "Vin = " + tb_vinList.Text + "\r\n";
            test_parameter.iout_conditions = "Iout = " + tb_iout.Text + "\r\n" +
                                             "VID Contions number = " + dataGridView2.RowCount + "\r\n";

            test_parameter.waveform_path = tbWave.Text;

            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.IoutList = tb_iout.Text.Split(',').Select(double.Parse).ToList();

            test_parameter.vidio.criteria.Clear();
            test_parameter.vidio.scope_ch2 = cbox_scope_ch2.SelectedIndex;
            test_parameter.vidio.il_level = (double)num_ILXLevel.Value;

            int temp = 0;
            for (int idx = 0; idx < dataGridView2.RowCount; idx++)
            {
                temp = Convert.ToInt16(dataGridView2[0, idx].Value) << 0 | // LPM
                        Convert.ToInt16(dataGridView2[1, idx].Value) << 1 | // G1
                        Convert.ToInt16(dataGridView2[2, idx].Value) << 2;  // G2

                test_parameter.vidio.io_table.Add(temp);
                test_parameter.vidio.vout_map.Add(dataGridView2[3, idx].Value, temp);
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                test_parameter.vidio.vout_list.Add(dataGridView1[0, i].Value);
                test_parameter.vidio.vout_list_af.Add(Convert.ToDouble(dataGridView1[1, i].Value));
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                VIDIO_Criteria_parameter criteria_container = new VIDIO_Criteria_parameter();
                criteria_container.vout_begin = dataGridView1[0, i].Value;          // vout start
                criteria_container.vout_end = dataGridView1[1, i].Value;            // vout end
                criteria_container.rise_time = dataGridView1[2, i].Value;           // rise time
                criteria_container.sr_rise = dataGridView1[3, i].Value;             // slew rate (rise)
                criteria_container.fall_time = dataGridView1[4, i].Value;           // fall time
                criteria_container.sr_fall = dataGridView1[5, i].Value;             // slew rate (fall)
                criteria_container.overshoot = Convert.ToDouble((string)dataGridView1[6, i].Value);          // overshoot (spec)
                criteria_container.undershoot = Convert.ToDouble((string)dataGridView1[7, i].Value);          // undershoot (spec)

                criteria_container.hi = Convert.ToDouble((string)dataGridView1[8, i].Value);
                criteria_container.lo = Convert.ToDouble((string)dataGridView1[9, i].Value);
                if((string)dataGridView1[3, i].Value != "NA") // sr judge enable
                    criteria_container.sr_time_jd = true;
                else 
                    criteria_container.sr_time_jd = false;

                if (dataGridView1[0, i].Value.ToString().IndexOf("LPM") != -1) // lpm judge enable
                    criteria_container.lpm_en = true;
                else
                    criteria_container.lpm_en = false;

                if ((string)dataGridView1[2, i].Value != "NA") // time judge enable
                    criteria_container.time_jd = true;
                else
                    criteria_container.time_jd = false;

                test_parameter.vidio.criteria.Add(criteria_container);
            }

            int program_max = test_parameter.VinList.Count *
                                test_parameter.IoutList.Count *
                                dataGridView1.RowCount;

            progressBar2.Maximum = program_max;
            test_parameter.vidio.discharge_time = (double)(nuDischarge.Value);
            test_parameter.vidio.discharge_load = (double)nuDisLoad.Value;
            test_parameter.vidio.test_cnt = (int)nu_test_cnt.Value;
            test_parameter.vidio.test_fail_cnt = (int)nu_test_fail_cnt.Value;
        }

        public void UpdateProgressBar(int val)
        {
            progressBar2.Invoke((MethodInvoker)(() => progressBar2.Value = val));
            labStatus.Invoke((MethodInvoker)(() => labStatus.Text = string.Format("Status Progress : {0} / {1}", val, progressBar2.Maximum)));
        }

        private void detect_datagridview()
        {
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                {

                    if(dataGridView1[j, i].Value == null)
                    {
                        dataGridView1[j, i].Value = "NA";
                    }
                }
            }
        }


        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            try
            {
                detect_datagridview();
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

            System.Windows.Forms.Button bt = (System.Windows.Forms.Button)sender;
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
            test_parameter.run_stop = true;
            if (ATETask != null)
            {
                if (ATETask.IsAlive)
                {
                    System.Threading.ThreadState state = ATETask.ThreadState;
                    if (state == System.Threading.ThreadState.Suspended) ATETask.Resume();
                    ATETask.Abort();
                    MessageBox.Show("ATE Task Stop !!", win_name, MessageBoxButtons.OK);
                    //InsControl._power.AutoPowerOff();
                }
            }
        }

        private void BTPause_Click(object sender, EventArgs e)
        {
            if (ATETask == null) return;
            System.Threading.ThreadState state = ATETask.ThreadState;
            if (state == System.Threading.ThreadState.Running || state == System.Threading.ThreadState.WaitSleepJoin)
            {
                ATETask.Suspend();
                BTPause.ForeColor = Color.Red;
            }
            else if (state == System.Threading.ThreadState.Suspended)
            {
                ATETask.Resume();
                BTPause.ForeColor = Color.White;
            }
        }

        private void BT_SaveSetting_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "settings|*.tb_info";

            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = saveDlg.FileName;
                SaveSettings(file_name);
            }
        }
        private void SaveSettings(string file)
        {
            string settings = "";

            // chamber info
            settings = "0.Chamber_en=$" + (ck_chamber_en.Checked ? "1" : "0") + "$\r\n";
            settings += "1.Chamber_temp=$" + tb_chamber.Text + "$\r\n";
            settings += "2.Chamber_time=$" + nu_steady.Value.ToString() + "$\r\n";

            // slave id
            settings += "3.Slave=$" + nuslave.Value.ToString() + "$\r\n";
            settings += "4.WavePath=$" + tbWave.Text + "$\r\n";
            settings += "5.Vin=$" + tb_vinList.Text + "$\r\n";
            settings += "6.Iout=$" + tb_iout.Text + "$\r\n";
            settings += "Scope_CH2=$" + cbox_scope_ch2.SelectedIndex + "$\r\n";
            settings += "ILX_Level=$" + num_ILXLevel.Value + "$\r\n";
            settings += "7.DGrow=$" + dataGridView2.RowCount + "$\r\n";
            settings += "8.DGrow_vout=$" + dataGridView1.RowCount + "$\r\n";


            for (int i = 0; i < dataGridView2.RowCount; i++)
            {
                settings += (i + 9).ToString() + ".LPM=$" + dataGridView2[0, i].Value.ToString() + "$\r\n";
                settings += (i + 10).ToString() + ".G1=$" + dataGridView2[1, i].Value.ToString() + "$\r\n";
                settings += (i + 11).ToString() + ".G2=$" + dataGridView2[2, i].Value.ToString() + "$\r\n";
                settings += (i + 12).ToString() + ".Vout=$" + dataGridView2[3, i].Value.ToString() + "$\r\n";
            }

            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                settings += (i + 13).ToString() + ".Vout=$" + dataGridView1[0, i].Value.ToString() + "$\r\n";
                settings += (i + 14).ToString() + ".Vout_af=$" + dataGridView1[1, i].Value.ToString() + "$\r\n";
                settings += (i + 15).ToString() + ".Rise_time=$" + dataGridView1[2, i].Value.ToString() + "$\r\n";
                settings += (i + 16).ToString() + ".Rise_SR=$" + dataGridView1[3, i].Value.ToString() + "$\r\n";
                settings += (i + 17).ToString() + ".Fall_time=$" + dataGridView1[4, i].Value.ToString() + "$\r\n";
                settings += (i + 18).ToString() + ".Fall_SR=$" + dataGridView1[5, i].Value.ToString() + "$\r\n";
                settings += (i + 19).ToString() + ".Vmax=$" + dataGridView1[6, i].Value.ToString() + "$\r\n";
                settings += (i + 20).ToString() + ".Vmin=$" + dataGridView1[7, i].Value.ToString() + "$\r\n";
                settings += (i + 21).ToString() + ".Vrise_point=$" + dataGridView1[8, i].Value.ToString() + "$\r\n";
                settings += (i + 22).ToString() + ".Vfall_point=$" + dataGridView1[9, i].Value.ToString() + "$\r\n";
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
                ck_chamber_en, tb_chamber, nu_steady, nuslave, tbWave, tb_vinList, tb_iout, cbox_scope_ch2, num_ILXLevel, dataGridView2 , dataGridView1
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
                            ((System.Windows.Forms.TextBox)obj_arr[i]).Text = info[i];
                            break;
                        case "CheckBox":
                            ((System.Windows.Forms.CheckBox)obj_arr[i]).Checked = info[i] == "1" ? true : false;
                            break;
                        case "NumericUpDown":
                            ((NumericUpDown)obj_arr[i]).Value = Convert.ToDecimal(info[i]);
                            break;
                        case "ComboBox":
                            ((ComboBox)obj_arr[i]).SelectedIndex = Convert.ToInt32(info[i]);
                            break;
                        case "RadioButton":
                            ((RadioButton)obj_arr[i]).Checked = info[i] == "1" ? true : false;
                            break;
                        case "DataGridView":
                            ((DataGridView)obj_arr[i]).RowCount = Convert.ToInt32(info[i]);
                            ((DataGridView)obj_arr[i + 1]).RowCount = Convert.ToInt32(info[i + 1]);

                            idx = i + 1;
                            goto fullDG;
                    }
                }

            fullDG:
                for (int i = 0; i < dataGridView2.RowCount; i++)
                {
                    dataGridView2[0, i].Value = Convert.ToString(info[idx + 1]); // lpm
                    dataGridView2[1, i].Value = Convert.ToString(info[idx + 2]); // g1
                    dataGridView2[2, i].Value = Convert.ToString(info[idx + 3]); // g2
                    dataGridView2[3, i].Value = Convert.ToString(info[idx + 4]); // vout
                    idx += 4;
                }

                for (int i = 0; i < dataGridView1.RowCount; i++)
                {
                    dataGridView1[0, i].Value = Convert.ToString(info[idx + 1]); // vout
                    dataGridView1[1, i].Value = Convert.ToString(info[idx + 2]); // vout_af

                    dataGridView1[2, i].Value = Convert.ToString(info[idx + 3]); 
                    dataGridView1[3, i].Value = Convert.ToString(info[idx + 4]); 
                    dataGridView1[4, i].Value = Convert.ToString(info[idx + 5]); 
                    dataGridView1[5, i].Value = Convert.ToString(info[idx + 6]); 
                    dataGridView1[6, i].Value = Convert.ToString(info[idx + 7]); 
                    dataGridView1[7, i].Value = Convert.ToString(info[idx + 8]); 
                    dataGridView1[8, i].Value = Convert.ToString(info[idx + 9]); 
                    dataGridView1[9, i].Value = Convert.ToString(info[idx + 10]); 

                    idx += 10;
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

        private void dataGridView1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            DV = (DataGridView)sender;
            if (DV.RowCount == 0)
            {
                ContextMenuStrip menu = new ContextMenuStrip();
                _item[0].Enabled = false;
                menu.Items.AddRange(_item);
                menu.Show(DV, new System.Drawing.Point(e.X, e.Y));
            }
            else
            {
                int row = DV.CurrentRow.Index;
                if (row < 0) return;
                ContextMenuStrip menu = new ContextMenuStrip();
                _item[0].Enabled = true;
                menu.Items.AddRange(_item);
                menu.Show(DV, new System.Drawing.Point(e.X, e.Y));
            }
        }

        private void dataGridView2_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            DV = (DataGridView)sender;
            if (DV.RowCount == 0)
            {
                ContextMenuStrip menu = new ContextMenuStrip();
                _item[0].Enabled = false;
                menu.Items.AddRange(_item);
                menu.Show(DV, new System.Drawing.Point(e.X, e.Y));
            }
            else
            {
                int row = DV.CurrentRow.Index;
                if (row < 0) return;
                ContextMenuStrip menu = new ContextMenuStrip();
                _item[0].Enabled = true;
                menu.Items.AddRange(_item);
                menu.Show(DV, new System.Drawing.Point(e.X, e.Y));
            }
        }

        private void vbase_ValueChanged(object sender, EventArgs e)
        {
            double vbase = (double)nu_vbase.Value;
            double vtop = (double)nu_vtop.Value;
            if (vbase > vtop) return;
            nu_diff.Value = (decimal)(vtop - vbase);
            nu_vmax.Value = (decimal)(vtop * 1.05);
            nu_vmin.Value = (decimal)(vbase * 0.95);
            numericUpDown1.Value = (decimal)(vtop - ((double)nu_diff.Value * 0.2));
            numericUpDown2.Value = (decimal)(vbase + ((double)nu_diff.Value * 0.2));
        }

        private void btn_lpm_on_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;

            switch (btn.TabIndex)
            {
                case 0:
                    RTDev.Gp1En_Enable();
                    break;
                case 1:
                    RTDev.Gp1En_Disable();
                    break;
            }
        }

        private void btn_g0_on_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;

            switch (btn.TabIndex)
            {
                case 0:
                    RTDev.Gp2En_Enable();
                    break;
                case 1:
                    RTDev.Gp2En_Disable();
                    break;
            }
        }

        private void btn_g1_on_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Button btn = (System.Windows.Forms.Button)sender;

            switch (btn.TabIndex)
            {
                case 0:
                    RTDev.Gp3En_Enable();
                    break;
                case 1:
                    RTDev.Gp3En_Disable();
                    break;
            }
        }

        private void cbox_scope_ch2_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbox_scope_ch2.SelectedIndex)
            {
                case 0:
                    num_ILXLevel.Enabled = false;
                    break;
                case 1:
                    num_ILXLevel.Enabled = true;
                    break;
            }
        }
    }


    public class VIDIO_parameter
    {
        public List<int> io_table = new List<int>();
        public Dictionary<object, int> vout_map = new Dictionary<object, int>();
        public Dictionary<double, int> lpm_vout_map = new Dictionary<double, int>();
        public List<string> lpm_str = new List<string>();
        public List<object> vout_list = new List<object>();
        public List<object> vout_list_af = new List<object>();
        public double discharge_time;
        public double discharge_load;
        public int test_cnt;
        public int test_fail_cnt;
        public int scope_ch2;
        public double il_level;

        public List<VIDIO_Criteria_parameter> criteria = new List<VIDIO_Criteria_parameter>();
    }



    public class VIDIO_Criteria_parameter
    {
        public object vout_begin;
        public object vout_end;
        public object rise_time;
        public object fall_time;
        public object sr_rise;
        public object sr_fall;
        public double overshoot;
        public double undershoot;

        public double spec_hi;
        public double spec_lo;

        public double hi;
        public double lo;

        public bool sr_time_jd;
        public bool time_jd;
        public bool lpm_en;
    }


}
