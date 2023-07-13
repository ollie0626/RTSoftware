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

using System.Text.RegularExpressions;
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
        ATE_DelayTime_Off _ate_delay_off = new ATE_DelayTime_Off();
        TaskRun[] ate_table;

        // device name
        System.Collections.Generic.Dictionary<string, string> Device_map = new Dictionary<string, string>();
        RTBBControl RTDev = new RTBBControl();

        public SoftStartTiming()
        {
            InitializeComponent();
            VisaCommand._IsDebug = false;
            RTDev.BoadInit();
            List<byte> list = RTDev.ScanSlaveID();
            if (list != null)
            {
                if (list.Count > 0)
                    nuslave.Value = list[0];
            }
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
                    if (InsControl._tek_scope_en)
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
                if (InsControl._tek_scope_en)
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
            TextBox[] power_off_path_table = new TextBox[] { tbBin4, tbBin5, tbBin6 };
            test_parameter.chamber_en = ck_chamber_en.Checked;
            test_parameter.run_stop = false;
            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.IoutList = tb_iout.Text.Split(',').Select(double.Parse).ToList();

            test_parameter.slave = (byte)nuslave.Value;
            test_parameter.offset_time = (double)nuOffset.Value;
            test_parameter.waveform_path = tbWave.Text;
            test_parameter.ontime_scale_ms = (double)nu_ontime_scale.Value;
            test_parameter.offtime_scale_ms = (double)nu_offtime_scale.Value;

            for (int i = 0; i < test_parameter.bin_path.Length; i++)
            {
                test_parameter.bin_path[i] = path_table[i].Text;
                test_parameter.power_off_bin_path[i] = power_off_path_table[i].Text;
            }

            // need to gui configure
            // scope channel 2 ~ 4
            for (int i = 0; i < test_parameter.scope_en.Length; i++)
            {
                test_parameter.scope_en[i] = ScopeChTable[i].Checked;
                test_parameter.bin_en[i] = binTable[i].Checked;
            }
            test_parameter.trigger_event = CbTrigger.SelectedIndex; // test example gpio trigger
            //test_parameter.sleep_mode = false;
            test_parameter.delay_us_en = RBUs.Checked;
            test_parameter.offset_time = RBUs.Checked ? ((double)nuOffset.Value * Math.Pow(10, -6)) : ((double)nuOffset.Value * Math.Pow(10, -3));
            test_parameter.gpio_pin = CBGPIO.SelectedIndex;
            test_parameter.judge_percent = ((double)nuCriteria.Value / 100);
            test_parameter.power_mode = CBChannel.Text;

            test_parameter.LX_Level = (double)nuLX.Value;
            test_parameter.ILX_Level = (double)nuILX.Value;

            test_parameter.Rail_en = (byte)nuData1.Value;
            test_parameter.Rail_addr = (byte)nuAddr.Value;

            test_parameter.item_idx = CBItem.SelectedIndex;
            test_parameter.eload_cr = ck_crmode.Checked;

            // CBEdge.SelectedIndex = 0 --> rising
            // 
            test_parameter.sleep_mode = (CBEdge.SelectedIndex == 0) ? true : false;
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
            if((int)idx == 3)
            {
                ate_table[0].temp = 25;
                ate_table[0].ATETask();

                ate_table[2].temp = 25;
                ate_table[2].ATETask();

            }
            else
            {
                ate_table[(int)idx].temp = 25;
                ate_table[(int)idx].ATETask();
            }

            BTRun.Invoke((MethodInvoker)(() => BTRun.Enabled = true));
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();

            //InsControl._scope.SaveWaveform(@"D:\", "scope");
            //OpenFileDialog opendlg = new OpenFileDialog();

            //if (opendlg.ShowDialog() == DialogResult.OK)
            //{

            //    Excel.Application _app;
            //    Excel.Worksheet _sheet;
            //    Excel.Workbook _book;
            //    Excel.Range _range;

            //    _app = new Excel.Application();
            //    _app.Visible = true;
            //    _book = (Excel.Workbook)_app.Workbooks.Open(opendlg.FileName);
            //    _sheet = (Excel.Worksheet)_book.ActiveSheet;
            //    byte[] data = new byte[255];
            //    for(int i = 1; i < 256; i++)
            //    {
            //        data[i - 1] = Convert.ToByte(_sheet.Cells[i, 2].Value, 16);
            //    }
            //    FileStream myFile = new FileStream(@"D:\123.bin", FileMode.OpenOrCreate);
            //    BinaryWriter bwr = new BinaryWriter(myFile);
            //    bwr.Write(data, 0, data.Length);
            //    myFile.Close();


            //    //string file_name = opendlg.FileName;
            //    //StreamReader sr = new StreamReader(file_name);
            //    //string line;
            //    //List<byte> temp = new List<byte>();
            //    //line = sr.ReadLine();
            //    //while(line != null)
            //    //{
            //    //    Console.WriteLine(line);
            //    //    string[] arr = line.Split('\t');
            //    //    line = sr.ReadLine();
            //    //    temp.Add(Convert.ToByte(arr[1], 16));
            //    //}
            //    //sr.Close();

            //    //FileStream myFile = new FileStream(@"D:\123.bin", FileMode.OpenOrCreate);
            //    //BinaryWriter bwr = new BinaryWriter(myFile);
            //    //bwr.Write(temp.ToArray(), 0, temp.Count);
            //    //bwr.Close();
            //    //myFile.Close();

            //    //string file_name = Path.GetFileNameWithoutExtension(opendlg.FileName);
            //    //MyLib.GetCriteria_time(file_name);
            //}
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
#if Power_en
                    InsControl._power.AutoPowerOff();
#endif
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

        private void PowerOffBinDisable()
        {
            BTSelectBinPath4.Enabled = false;
            BTSelectBinPath5.Enabled = false;
            BTSelectBinPath6.Enabled = false;
            tbBin4.Enabled = false;
            tbBin5.Enabled = false;
            tbBin6.Enabled = false;
        }

        private void PowerOffBinEnable()
        {
            BTSelectBinPath4.Enabled = true;
            BTSelectBinPath5.Enabled = true;
            BTSelectBinPath6.Enabled = true;
            tbBin4.Enabled = true;
            tbBin5.Enabled = true;
            tbBin6.Enabled = true;
        }

        private void CBItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CBItem.SelectedIndex)
            {
                case 0:
                    group_sst.Visible = false;
                    group_channel.Visible = true;

                    CkBin1.Checked = true;
                    CkBin2.Checked = false;
                    CkBin3.Checked = false;
                    CkBin2.Enabled = true;
                    CkBin3.Enabled = true;

                    switch (CbTrigger.SelectedIndex)
                    {
                        case 0:
                            tb_connect1.Text = "PWRDIS / Sleep";
                            break;
                        case 1:
                            tb_connect1.Text = "I2C (SCL)";
                            break;
                        case 2:
                            tb_connect1.Text = "Vin";
                            break;
                    }

                    tb_connect2.Text = "Rail 1";
                    tb_connect3.Text = "Rail 2";
                    tb_connect4.Text = "Rail 3";

                    PowerOffBinDisable();
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

                    switch (CbTrigger.SelectedIndex)
                    {
                        case 0:
                            tb_connect1.Text = "PWRDIS / Sleep";
                            break;
                        case 1:
                            tb_connect1.Text = "I2C (SCL)";
                            break;
                        case 2:
                            tb_connect1.Text = "Vin";
                            break;
                    }

                    tb_connect2.Text = "Vout";
                    tb_connect3.Text = "LX";
                    tb_connect4.Text = "ILX";

                    PowerOffBinDisable();
                    break;
                case 2:
                    // power off delay
                    group_sst.Visible = false;
                    group_channel.Visible = true;
                    PowerOffBinDisable();
                    break;
                case 3:
                    // continue test
                    group_sst.Visible = false;
                    group_channel.Visible = true;
                    PowerOffBinEnable();
                    break;
            }
        }

        private void CbTrigger_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CbTrigger.SelectedIndex)
            {
                case 0:
                    tb_connect1.Text = "PWRDIS / Sleep";
                    CBGPIO.Enabled = true;

                    // ----------------------
                    // GUI setting
                    labAddr.Visible = false;
                    labRail_en.Visible = false;
                    nuAddr.Visible = false;
                    nuData1.Visible = false;
                    break;
                case 1:
                    tb_connect1.Text = "I2C (SCL)";
                    CBGPIO.Enabled = false;

                    // ----------------------
                    // GUI setting
                    labAddr.Visible = true;
                    labRail_en.Visible = true;
                    nuAddr.Visible = true;
                    nuData1.Visible = true;
                    break;
                case 2:
                    tb_connect1.Text = "Vin";
                    CBGPIO.Enabled = false;

                    // ----------------------
                    // GUI setting
                    labAddr.Visible = false;
                    labRail_en.Visible = true;
                    nuAddr.Visible = false;
                    nuData1.Visible = false;
                    break;
            }
        }

        private void BTSelectBinPath4_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin4.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTSelectBinPath5_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin5.Text = folderBrowser.SelectedPath;
            }
        }

        private void BTSelectBinPath6_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin6.Text = folderBrowser.SelectedPath;
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
            settings += "4.Addr=$" + nuAddr.Value.ToString() + "$\r\n";
            settings += "5.Rail_en=$" + nuData1.Value.ToString() + "$\r\n";

            // bin folder
            settings += "6.Bin1=$" + tbBin.Text + "$\r\n";
            settings += "7.Bin2=$" + tbBin2.Text + "$\r\n";
            settings += "8.Bin3=$" + tbBin3.Text + "$\r\n";
            settings += "9.Bin4=$" + tbBin4.Text + "$\r\n";
            settings += "10.Bin5=$" + tbBin5.Text + "$\r\n";
            settings += "11.Bin6=$" + tbBin6.Text + "$\r\n";

            settings += "12.WavePath=$" + tbWave.Text + "$\r\n";
            settings += "13.Trigger_event=$" + CbTrigger.SelectedIndex + "$\r\n";
            settings += "14.GPIO_sel=$" + CBGPIO.SelectedIndex + "$\r\n";

            settings += "15.Bin1_en=$" + (CkBin1.Checked ? "1" : "0") + "$\r\n";
            settings += "16.Bin2_en=$" + (CkBin2.Checked ? "1" : "0") + "$\r\n";
            settings += "17.Bin3_en=$" + (CkBin3.Checked ? "1" : "0") + "$\r\n";

            settings += "18.Scope_Ch1_en=$" + (CkCH1.Checked ? "1" : "0") + "$\r\n";
            settings += "19.Scope_Ch2_en=$" + (CkCH2.Checked ? "1" : "0") + "$\r\n";
            settings += "20.Scope_Ch3_en=$" + (CkCH3.Checked ? "1" : "0") + "$\r\n";

            settings += "21.Lx_level=$" + nuLX.Value.ToString() + "$\r\n";
            settings += "22.ILx_level=$" + nuILX.Value.ToString() + "$\r\n";

            settings += "23.On_TimeScale=$" + nu_ontime_scale.Value.ToString() + "$\r\n";
            settings += "24.Off_TimeScale=$" + nu_offtime_scale.Value.ToString() + "$\r\n";

            settings += "25.Vin=$" + tb_vinList.Text + "$\r\n";
            settings += "26.Unit=$" + (RBUs.Checked ? "1" : "0") + "$\r\n";
            settings += "27.Time_offset=$" + nuOffset.Value.ToString() + "$\r\n";

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
                ck_chamber_en, tb_chamber, nu_steady, nuslave, nuAddr, nuData1,
                tbBin, tbBin2, tbBin3, tbBin4, tbBin5, tbBin6, tbWave, CbTrigger,
                CBGPIO, CkBin1, CkBin2, CkBin3, CkCH1, CkCH2, CkCH3, nuLX,
                nuILX, nu_ontime_scale, nu_offtime_scale, tb_vinList,
                RBUs, nuOffset
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
                        case "RadioButton":
                            ((RadioButton)obj_arr[i]).Checked = info[i] == "1" ? true : false;
                            break;
                    }
                }
            }
        }

        private void ck_crmode_CheckedChanged(object sender, EventArgs e)
        {
            if(ck_crmode.Checked)
            {
                groupBox2.Text = "Iout Range (ohm)";
            }
            else
            {
                groupBox2.Text = "Iout Range (A)";
            }
        }
    }
}
