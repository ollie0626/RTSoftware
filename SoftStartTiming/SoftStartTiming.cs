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

namespace SoftStartTiming
{
    public partial class SoftStartTiming : Form
    {
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        int SteadyTime;
        string[] tempList;

        // test item
        ATE_DelayTime _ate_delay_time = new ATE_DelayTime();
        ATE_SoftStartTime _ate_sst = new ATE_SoftStartTime();
        TaskRun[] ate_table;

        // device name
        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();


        public SoftStartTiming()
        {
            InitializeComponent();
            VisaCommand._IsDebug = false;
        }

        private void BTSelectBinPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTSelectBinPath2_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin2.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTSelectBinPath3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin3.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTSelectWavePath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbWave.Text = folderBrowser.SelectedPath;
            }
        }

        private int ConnectFunc(string res, int ins_sel)
        {
            switch (ins_sel)
            {
                case 0: 
                    if(InsControl._tek_scope_en)
                    {
                        InsControl._tek_scope = new TekTronix7Serise(res);
                    }
                    else
                    {
                        InsControl._scope = new AgilentOSC(res);
                    }
                    
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

            // funcgen AFG31022
            MyLib.Delay1s(1);
            check_ins_state();
        }

        private void check_ins_state()
        {
            if (InsControl._scope != null || InsControl._tek_scope != null)
            {
                if(InsControl._tek_scope_en)
                {
                    if (InsControl._tek_scope.InsState())
                        led_osc.BackColor = Color.LightGreen;
                    else
                        led_osc.BackColor = Color.Red;
                }
                else
                {
                    if (InsControl._scope.InsState())
                        led_osc.BackColor = Color.LightGreen;
                    else
                        led_osc.BackColor = Color.Red;
                }

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
            // test condition
            test_parameter.vin_conditions = "Vin :" + tb_vinList.Text + " (V)\r\n";
            test_parameter.bin1_cnt = CkBin1.Checked ? MyLib.ListBinFile(tbBin.Text).Length : 0;
            test_parameter.bin2_cnt = CkBin2.Checked ? MyLib.ListBinFile(tbBin2.Text).Length : 0;
            test_parameter.bin3_cnt = CkBin3.Checked ? MyLib.ListBinFile(tbBin3.Text).Length : 0;

            test_parameter.bin_file_cnt = "Bin1 file cnt : " + test_parameter.bin1_cnt + "\r\n" +
                                          "Bin2 file cnt : " + test_parameter.bin2_cnt + "\r\n" +
                                          "Bin3 file cnt : " + test_parameter.bin3_cnt + "\r\n" +
                                          "Total cnt : " + (test_parameter.bin1_cnt + test_parameter.bin2_cnt + test_parameter.bin3_cnt).ToString() + " \r\n";
            test_parameter.tool_ver = win_name + "\r\n";

            TextBox[] path_table = new TextBox[] { tbBin, tbBin2, tbBin3 };
            test_parameter.chamber_en = ck_chamber_en.Checked;
            test_parameter.run_stop = false;
            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.slave = (byte)nuslave.Value;
            test_parameter.offset_time = (double)nuOffset.Value;
            test_parameter.waveform_path = tbWave.Text;
            test_parameter.ontime_scale_ms = (double)nu_ontime_scale.Value;
            test_parameter.offtime_scale_ms = (double)nu_offtime_scale.Value;
            for (int i = 0; i < test_parameter.bin_path.Length; i++)
            {
                test_parameter.bin_path[i] = path_table[i].Text;
            }

            // need to gui configure
            // scope channel 2 ~ 4
            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                test_parameter.scope_en[i] = ScopeChTable[i].Checked;
                test_parameter.bin_en[i] = binTable[i].Checked;
            }
            test_parameter.trigger_event = CbTrigger.SelectedIndex; // test example gpio trigger
            test_parameter.sleep_mode = false;
            test_parameter.delay_us_en = RBUs.Checked;
            test_parameter.offset_time = RBUs.Checked ? ((double)nuOffset.Value * Math.Pow(10, -6)) : ((double)nuOffset.Value * Math.Pow(10, -3));
            test_parameter.gpio_pin = CBGPIO.SelectedIndex;
            test_parameter.judge_percent = ((double)nuCriteria.Value / 100);
            test_parameter.power_mode = CBChannel.Text;

            test_parameter.LX_Level = (double)nuLX.Value;
            test_parameter.ILX_Level = (double)nuILX.Value;
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
                    ATETask.Start(CBItem.SelectedIndex);
                }
                else
                {
                    // none Chamber

                    // Delay Time / Slot Time
                    // Soft - Start Time
                    p_thread = new ParameterizedThreadStart(Run_Single_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(CBItem.SelectedIndex);
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

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].temp = 25;
            ate_table[(int)idx].ATETask();
            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //ProcessStartInfo psi = new ProcessStartInfo();
            //psi.Arguments = "/im EXCEL.EXE /f";
            //psi.FileName = "taskkill";
            //Process p = new Process();
            //p.StartInfo = psi;
            //p.Start();

            //InsControl._scope.SaveWaveform(@"D:\", "scope");
            OpenFileDialog opendlg = new OpenFileDialog();

            if (opendlg.ShowDialog() == DialogResult.OK)
            {

                Excel.Application _app;
                Excel.Worksheet _sheet;
                Excel.Workbook _book;
                Excel.Range _range;

                _app = new Excel.Application();
                _app.Visible = true;
                _book = (Excel.Workbook)_app.Workbooks.Open(opendlg.FileName);
                _sheet = (Excel.Worksheet)_book.ActiveSheet;
                byte[] data = new byte[255];
                for(int i = 1; i < 256; i++)
                {
                    data[i - 1] = Convert.ToByte(_sheet.Cells[i, 2].Value, 16);
                }
                FileStream myFile = new FileStream(@"D:\123.bin", FileMode.OpenOrCreate);
                BinaryWriter bwr = new BinaryWriter(myFile);
                bwr.Write(data, 0, data.Length);
                myFile.Close();


                //string file_name = opendlg.FileName;
                //StreamReader sr = new StreamReader(file_name);
                //string line;
                //List<byte> temp = new List<byte>();
                //line = sr.ReadLine();
                //while(line != null)
                //{
                //    Console.WriteLine(line);
                //    string[] arr = line.Split('\t');
                //    line = sr.ReadLine();
                //    temp.Add(Convert.ToByte(arr[1], 16));
                //}
                //sr.Close();

                //FileStream myFile = new FileStream(@"D:\123.bin", FileMode.OpenOrCreate);
                //BinaryWriter bwr = new BinaryWriter(myFile);
                //bwr.Write(temp.ToArray(), 0, temp.Count);
                //bwr.Close();
                //myFile.Close();

                //string file_name = Path.GetFileNameWithoutExtension(opendlg.FileName);
                //MyLib.GetCriteria_time(file_name);
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
                    MessageBox.Show("ATE Task Stop !!", "ATE Tool", MessageBoxButtons.OK);
                    InsControl._power.AutoPowerOff();
                }
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
                        for(int i = 0; i < scope_name.Length; i++)
                        {
                            if(name == scope_name[i]) 
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
            switch (CBItem.SelectedIndex)
            {
                case 0:
                    // delay time
                    group_sst.Visible = false;
                    group_channel.Visible = true;

                    CkBin1.Checked = true;
                    CkBin2.Checked = false;
                    CkBin3.Checked = false;
                    CkBin2.Enabled = true;
                    CkBin3.Enabled = true;
                    break;
                case 1:
                    // soft-start time
                    group_sst.Visible = true;
                    group_channel.Visible = false;

                    CkBin1.Checked = true;
                    CkBin2.Checked = false;
                    CkBin3.Checked = false;
                    CkBin2.Enabled = false;
                    CkBin3.Enabled = false;
                    break;
            }
        }
    }
}
