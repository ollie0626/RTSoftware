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
using System.Reflection.Emit;
using SoftStartTiming.Properties;

using System.Runtime.InteropServices;
using System.Data.Common;

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


        public struct PowerInfo{
            public string Dev;
            public string ins;
            public int addr;
        }

        List<PowerInfo> PowerInfoList = new List<PowerInfo>();


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
            test_parameter.i2c_init_dg = i2c_datagrid;
            test_parameter.i2c_mtp_dg = i2c_mtp_datagrid;

            string vin_tmp = "";
            for(int i = 0; i < test_dg.RowCount; i++)
            {
                if (i == test_dg.RowCount - 1) vin_tmp += test_dg[0, i].Value.ToString();
                else vin_tmp += test_dg[0, i].Value.ToString() + ", ";
            }

            test_parameter.vin_conditions = "Vin :" + vin_tmp + " (V)\r\n";
            //test_parameter.bin1_cnt = CkBin1.Checked ? MyLib.ListBinFile(tbBin.Text).Length : 0;
            //test_parameter.bin2_cnt = CkBin2.Checked ? MyLib.ListBinFile(tbBin2.Text).Length : 0;
            //test_parameter.bin3_cnt = CkBin3.Checked ? MyLib.ListBinFile(tbBin3.Text).Length : 0;

            //test_parameter.bin_file_cnt = "Bin1 file cnt : " + test_parameter.bin1_cnt + "\r\n" +
            //                              "Bin2 file cnt : " + test_parameter.bin2_cnt + "\r\n" +
            //                              "Bin3 file cnt : " + test_parameter.bin3_cnt + "\r\n" +
            //                              "Total cnt : " + (test_parameter.bin1_cnt + test_parameter.bin2_cnt + test_parameter.bin3_cnt).ToString() + " \r\n";

            test_parameter.conditions = "Measure setting:\r\n" + 
                                        cbox_dly0_from.Text + " → " + cbox_dly0_to.Text + "\r\n" +
                                        cbox_dly1_from.Text + " → " + cbox_dly1_to.Text + "\r\n" +
                                        cbox_dly2_from.Text + " → " + cbox_dly2_to.Text + "\r\n" +
                                       "Test cnt: " + test_dg.RowCount.ToString() + "\r\n";

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
            test_parameter.Rail_dis = (byte)nuData2.Value;
            test_parameter.Rail_addr = (byte)nuAddr.Value;

            test_parameter.item_idx = CBItem.SelectedIndex;
            test_parameter.eload_cr = ck_crmode.Checked;

            // CBEdge.SelectedIndex = 0 --> rising
            // sleep_mode: rising
            // pwr_dis_mode: falling
            test_parameter.sleep_mode = (CBEdge.SelectedIndex == 0) ? true : false;

            // delay time test conditions
            test_parameter.seq_dg = test_dg;
            test_parameter.cursor_disable = ck_cursor_disable.Checked;
            test_parameter.auto_en[0] = CkCH0.Checked;
            test_parameter.auto_en[1] = CkCH1.Checked;
            test_parameter.auto_en[2] = CkCH2.Checked;
            test_parameter.auto_en[3] = CkCH3.Checked;

            test_parameter.seq_en[0] = ck_MeasSeq0.Checked;
            test_parameter.seq_en[1] = ck_MeasSeq1.Checked;
            test_parameter.seq_en[2] = ck_MeasSeq2.Checked;
            test_parameter.seq_en[3] = ck_MeasSeq3.Checked;
        }

        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            try
            {
                RTDev.BoadInit();
                List<byte> list = RTDev.ScanSlaveID();

                if (list != null)
                {
                    if (list.Count > 0)
                        nuslave.Value = list[0];
                }

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
            if ((int)idx == 3)
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

            // ins --> GPIB name
            foreach (string ins in ins_list)
            {
                list_ins.Items.Add(ins);
                VisaCommand visaCommand = new VisaCommand();
                visaCommand.LinkingIns(ins);
                string idn = visaCommand.doQueryIDN();
                string name = "";

                // split scan result (IDN)
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
                }

                if (name.IndexOf("E363") != -1)
                {

                    PowerInfo powerinfo_s = new PowerInfo();
                    int addr = Convert.ToInt32(ins.Replace("::", ":").Split(':')[1]);
                    powerinfo_s.Dev = name;
                    powerinfo_s.ins = ins;
                    powerinfo_s.addr = addr;
                    PowerInfoList.Add(powerinfo_s);
                    CBPower.Enabled = true;
                    CBPower.Items.Add(name);
                }

                if (name.IndexOf("62006P") != -1)
                {
                    PowerInfo powerinfo_s = new PowerInfo();
                    int addr = Convert.ToInt32(ins.Replace("::", ":").Split(':')[1]);
                    powerinfo_s.Dev = name;
                    powerinfo_s.ins = ins;
                    powerinfo_s.addr = addr;
                    PowerInfoList.Add(powerinfo_s);
                    CBPower.Enabled = true;
                    CBPower.Items.Add(name);
                }
            }
        }

        private void CBPower_SelectedIndexChanged(object sender, EventArgs e)
        {
            InsControl._power = new PowerModule(PowerInfoList[CBPower.SelectedIndex].ins);
            nuPower_addr.Value = PowerInfoList[CBPower.SelectedIndex].addr;

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


        private void CbTrigger_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (CbTrigger.SelectedIndex)
            {
                case 0:
                    tb_connect1.Text = "PWRDIS / Sleep";
                    CBGPIO.Enabled = true;

                    // ----------------------
                    // GUI setting
                    //labAddr.Visible = false;
                    //labRail_en.Visible = false;
                    //label17.Visible = false;
                    //nuAddr.Visible = false;
                    //nuData1.Visible = false;
                    //nuData2.Visible = false;
                    break;
                case 1:
                    tb_connect1.Text = "I2C (SCL)";
                    CBGPIO.Enabled = false;

                    // ----------------------
                    // GUI setting
                    labAddr.Visible = true;
                    labRail_en.Visible = true;
                    label17.Visible = true;
                    nuAddr.Visible = true;
                    nuData1.Visible = true;
                    nuData2.Visible = true;
                    break;
                case 2:
                    tb_connect1.Text = "Vin";
                    CBGPIO.Enabled = false;

                    // ----------------------
                    // GUI setting
                    //labAddr.Visible = false;
                    //labRail_en.Visible = false;
                    //label17.Visible = false;
                    //nuAddr.Visible = false;
                    //nuData1.Visible = false;
                    //nuData2.Visible = false;
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
            saveDlg.Filter = "settings|*.xlsx";

            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = saveDlg.FileName;
                SaveSettings(file_name);
            }
        }

        private void title_set(Excel.Worksheet sheet, int row, int col)
        {
            Excel.Range range = sheet.Cells[row, col];
            range.Font.Bold = true;
            range.Interior.Color = 65535;

            Marshal.ReleaseComObject(range);
        }

        private void data_set(Excel.Worksheet sheet, ref int row, DataGridView dg)
        {
            int col = 1;
            for(int i = 0; i < dg.ColumnCount; i++)
            {
                sheet.Cells[row + i, col] = dg.Columns[i].HeaderText;
            }

            for(int i = 0; i < dg.RowCount; i++) // ↓
            {
                for(int j = 0; j < dg.ColumnCount;  j++) // →
                {
                    // excel -> (row, col), dg -> (col, row)
                    sheet.Cells[row + j, col + 1 + i] = dg[j, i].Value;
                }
            }

            for (int i = 0; i < dg.ColumnCount; i++) row++;
        }

        private void SaveSettings(string file)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Add();
            Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[1];
            app.Visible = true;
            //string settings = "";

            int row = 1;
            int col = 1;

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "Chamber Config";
            sheet.Cells[row, col] = "Chamber En";
            sheet.Cells[row++, col + 1] = ck_chamber_en.Checked;
            sheet.Cells[row, col] = "Chamber Temp";
            sheet.Cells[row++, col + 1] = tb_templist.Text;
            sheet.Cells[row, col] = "Chamber steady time";
            sheet.Cells[row, col + 1].Numberformat = "@";
            sheet.Cells[row++, col + 1] = nu_steady.Value;

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "I2C Rail Enable/Disable";
            sheet.Cells[row, col] = "Slave";
            sheet.Cells[row, col + 1].Numberformat = "@";
            sheet.Cells[row++, col + 1] = string.Format("{0:X}", (int)nuslave.Value);
            sheet.Cells[row, col] = "Addr";
            sheet.Cells[row, col + 1].Numberformat = "@";
            sheet.Cells[row++, col + 1] = string.Format("{0:X}", (int)nuAddr.Value);
            sheet.Cells[row, col] = "Rail Enable/Disable";
            sheet.Cells[row, col + 1].Numberformat = "@";
            sheet.Cells[row, col + 2].Numberformat = "@";
            sheet.Cells[row, col + 1] =  string.Format("{0:X}", (int)nuData1.Value);
            sheet.Cells[row++, col + 2] = string.Format("{0:X}", (int)nuData2.Value);

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "General setting";
            sheet.Cells[row, col] = "Wave path";
            sheet.Cells[row++, col + 1] = tbWave.Text;

            sheet.Cells[row, col] = "Trigger Event";
            sheet.Cells[row++, col + 1] = CbTrigger.SelectedIndex;

            sheet.Cells[row, col] = "Scope Trigger CH";
            sheet.Cells[row++, col + 1] = cbox_trigger.SelectedIndex;


            sheet.Cells[row, col] = "GPIO_sel";
            sheet.Cells[row++, col + 1] = CBGPIO.SelectedIndex;

            // -------------------------------------------------------
            // scope setting
            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "Scope Setting";
            sheet.Cells[row, col] = "Channel Resize";
            sheet.Cells[row, col + 1] = CkCH0.Checked;
            sheet.Cells[row, col + 2] = CkCH1.Checked;
            sheet.Cells[row, col + 3] = CkCH2.Checked;
            sheet.Cells[row++, col + 4] = CkCH3.Checked;

            sheet.Cells[row, col] = "Seq Meas En";
            sheet.Cells[row, col + 1] = ck_MeasSeq0.Checked;
            sheet.Cells[row, col + 2] = ck_MeasSeq1.Checked;
            sheet.Cells[row, col + 3] = ck_MeasSeq2.Checked;
            sheet.Cells[row++, col + 4] = ck_MeasSeq3.Checked;

            title_set(sheet, row, col);
            sheet.Cells[row, col] = "Power on/off TimeScale";
            sheet.Cells[row, col + 1] = nu_ontime_scale.Value.ToString();
            sheet.Cells[row++, col + 2] = nu_offtime_scale.Value.ToString();
            sheet.Cells[row, col] = "Time Offset";
            sheet.Cells[row++, col + 1] = nuOffset.Value;

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "Measure Start and End";
            sheet.Cells[row, col] = "Seq0";
            sheet.Cells[row, col + 1] = cbox_dly0_from.SelectedIndex;
            sheet.Cells[row, col + 2] = cbox_dly0_to.SelectedIndex;
            sheet.Cells[row, col + 3] = nudly0_from.Value;
            sheet.Cells[row, col + 4] = nudly0_end.Value;
            sheet.Cells[row++, col + 5] = nu_ch0_level.Value;

            sheet.Cells[row, col] = "Seq1";
            sheet.Cells[row, col + 1] = cbox_dly1_from.SelectedIndex;
            sheet.Cells[row, col + 2] = cbox_dly1_to.SelectedIndex;
            sheet.Cells[row, col + 3] = nudly1_from.Value;
            sheet.Cells[row, col + 4] = nudly1_end.Value;
            sheet.Cells[row++, col + 5] = nu_ch1_level.Value;

            sheet.Cells[row, col] = "Seq2";
            sheet.Cells[row, col + 1] = cbox_dly2_from.SelectedIndex;
            sheet.Cells[row, col + 2] = cbox_dly2_to.SelectedIndex;
            sheet.Cells[row, col + 3] = nudly2_from.Value;
            sheet.Cells[row, col + 4] = nudly2_end.Value;
            sheet.Cells[row++, col + 5] = nu_ch2_level.Value;

            sheet.Cells[row, col] = "Seq3";
            sheet.Cells[row, col + 1] = cbox_dly3_from.SelectedIndex;
            sheet.Cells[row, col + 2] = cbox_dly3_to.SelectedIndex;
            sheet.Cells[row, col + 3] = nudly3_from.Value;
            sheet.Cells[row, col + 4] = nudly3_end.Value;
            sheet.Cells[row++, col + 5] = nu_ch3_level.Value;

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "Disable cursor function";
            sheet.Cells[row, col] = "Disable State";
            sheet.Cells[row++, col + 1] = ck_cursor_disable.Checked;

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "I2C Init Config";
            data_set(sheet, ref row, i2c_datagrid);

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "I2C MTP Config";
            data_set(sheet, ref row, i2c_mtp_datagrid);

            title_set(sheet, row, col);
            sheet.Cells[row++, col] = "Test Config";
            data_set(sheet, ref row, test_dg);


            sheet.Columns[1].AutoFit();

            sheet.SaveAs(file);
            book.Save();
            book.Close();
            app.Quit();

            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(app);
        }

        private void BT_LoadSetting_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();
            opendlg.Filter = "settings|*.xlsx";
            if (opendlg.ShowDialog() == DialogResult.OK)
            {
                LoadSettings(opendlg.FileName);
            }
        }

        private int GetLastColumn(Excel.Worksheet sheet, int row)
        {
            return sheet.Cells[row, sheet.Columns.Count].End[Excel.XlDirection.xlToLeft].Column;
        }

        private void data_import(Excel.Worksheet sheet, ref int row, DataGridView dg)
        {
            // excel row number
            int row_number = 9;
            string temp = sheet.Cells[row, 1].Value;
            if (temp == "Address") row_number = 2;

            // excel col number --> dg row number
            int last_col = GetLastColumn(sheet, row);
            dg.RowCount = last_col - 1;


            for (int j = 0; j < row_number; j++)
            {
                for (int i = 2; i < last_col + 1; i++)
                {
                    temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(i) + row].Value);
                    dg[j, i - 2].Value = (object)temp;
                }
                row++;
            }
        }

        private void LoadSettings(string file)
        {
            Excel.Application app = new Excel.Application();
            Excel.Workbook book = app.Workbooks.Open(file);
            Excel.Worksheet sheet = (Excel.Worksheet)book.Sheets[1];

            string temp = "";
            int row = 2;
            int col = 1;

            int last_col = GetLastColumn(sheet, row);
            ck_chamber_en.Checked = sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value;
            row++;

            last_col = GetLastColumn(sheet, row);
            tb_templist.Text = sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value;
            row++;

            last_col = GetLastColumn(sheet, row);
            nu_steady.Value = Convert.ToInt32(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            row += 2;

            // i2c Rail Enable/Disable
            last_col = GetLastColumn(sheet, row);
            nuslave.Value = Convert.ToInt32(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value, 16);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            nuAddr.Value = Convert.ToInt32(temp, 16);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nuData1.Value = Convert.ToInt32(temp, 16);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            nuData2.Value = Convert.ToInt32(temp, 16);
            row += 2;

            // General setting
            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            tbWave.Text = temp;
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            CbTrigger.SelectedIndex = Convert.ToInt32(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            cbox_trigger.SelectedIndex = Convert.ToInt32(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            CBGPIO.SelectedIndex = Convert.ToInt32(temp);
            row += 2;


            last_col = GetLastColumn(sheet, row);
            CkCH0.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value;
            CkCH1.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value;
            CkCH2.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value;
            CkCH3.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value;
            row++;

            last_col = GetLastColumn(sheet, row);
            ck_MeasSeq0.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value;
            ck_MeasSeq1.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value;
            ck_MeasSeq2.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value;
            ck_MeasSeq3.Checked = sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value;
            row++;


            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nu_ontime_scale.Value = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            nu_offtime_scale.Value = Convert.ToInt32(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col) + row].Value);
            nuOffset.Value = Convert.ToInt32(temp);
            row += 2;


            // Measure Start and Stop
            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 4) + row].Value);
            cbox_dly0_from.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value);
            cbox_dly0_to.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value);
            nudly0_from.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nudly0_end.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value);
            nu_ch0_level.Value = Convert.ToDecimal(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 4) + row].Value);
            cbox_dly1_from.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value);
            cbox_dly1_to.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value);
            nudly1_from.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nudly1_end.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value);
            nu_ch1_level.Value = Convert.ToDecimal(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 4) + row].Value);
            cbox_dly2_from.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value);
            cbox_dly2_to.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value);
            nudly2_from.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nudly2_end.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value);
            nu_ch2_level.Value = Convert.ToDecimal(temp);
            row++;

            last_col = GetLastColumn(sheet, row);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 4) + row].Value);
            cbox_dly3_from.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 3) + row].Value);
            cbox_dly3_to.SelectedIndex = Convert.ToInt32(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 2) + row].Value);
            nudly3_from.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 1) + row].Value);
            nudly3_end.Value = Convert.ToDecimal(temp);
            temp = Convert.ToString(sheet.Range[MyLib.ConvertToLetter(last_col - 0) + row].Value);
            nu_ch3_level.Value = Convert.ToDecimal(temp);
            row += 2;

            // i2c Init config
            data_import(sheet, ref row, i2c_datagrid); row += 1;
            data_import(sheet, ref row, i2c_mtp_datagrid); row += 1;
            data_import(sheet, ref row, test_dg);

            book.Close();
            app.Quit();

            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(book);
            Marshal.ReleaseComObject(app);

        }

        private void ck_crmode_CheckedChanged(object sender, EventArgs e)
        {
            if (ck_crmode.Checked)
            {
                groupBox2.Text = "Iout Range (ohm)";
            }
            else
            {
                groupBox2.Text = "Iout Range (A)";
            }
        }

        private void btn_i2c_data_Click(object sender, EventArgs e)
        {
            i2c_datagrid.RowCount++;
            int idx = i2c_datagrid.RowCount - 1;
            i2c_datagrid[0, idx].Value = string.Format("{0:X}", (int)nuaddr_to_dg.Value);
            i2c_datagrid[1, idx].Value = string.Format("{0:X}", (int)nudata_to_dg.Value);
        }

        private void btn_i2c_mtp_data_Click(object sender, EventArgs e)
        {
            i2c_mtp_datagrid.RowCount++;
            int idx = i2c_mtp_datagrid.RowCount - 1;
            i2c_mtp_datagrid[0, idx].Value = string.Format("{0:X}", (int)nu_addr_mtp.Value);
            i2c_mtp_datagrid[1, idx].Value = string.Format("{0:X}", (int)nu_data_mtp.Value);
        }

        private void CBItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CBItem.SelectedIndex == 1)
            {
                CkBin2.Enabled = false;
                CkBin3.Enabled = false;

                CkBin2.Checked = false;
                CkBin3.Checked = false;

                CkBin1.Checked = true;
            }
            else
            {
                CkBin2.Enabled = true;
                CkBin3.Enabled = true;
            }

        }

        private void bt_add_to_table_Click(object sender, EventArgs e)
        {
            ComboBox[] fromTable = new ComboBox[] { cbox_dly0_from, cbox_dly1_from, cbox_dly2_from, cbox_dly3_from };
            ComboBox[] toTable = new ComboBox[] { cbox_dly0_to, cbox_dly1_to, cbox_dly2_to, cbox_dly3_to };
            ComboBox[] EloadTable = new ComboBox[] { cbox_eload_ch1, cbox_eload_ch2, cbox_eload_ch3, cbox_eload_ch4 };

            NumericUpDown[] percent_pos1 = new NumericUpDown[] { nudly0_from, nudly1_from, nudly2_from, nudly3_from };
            NumericUpDown[] percent_pos2 = new NumericUpDown[] { nudly0_end, nudly1_end, nudly2_end, nudly3_end };
            NumericUpDown[] initLevel = new NumericUpDown[] { nu_ch0_level, nu_ch1_level, nu_ch2_level, nu_ch3_level };
            NumericUpDown[] seqTable_addr = new NumericUpDown[] { nu_seq0_addr, nu_seq1_addr, nu_seq2_addr, nu_seq3_addr };
            NumericUpDown[] seqTable_data = new NumericUpDown[] { nu_seq0_data, nu_seq1_data, nu_seq2_data, nu_seq3_data };
            NumericUpDown[] idelTable_addr = new NumericUpDown[] { nu_idel0_addr, nu_idel1_addr, nu_idel2_addr, nu_idel3_addr };
            NumericUpDown[] idelTable_data = new NumericUpDown[] { nu_idel0_data, nu_idel1_data, nu_idel2_data, nu_idel3_data };
            NumericUpDown[] idelTable = new NumericUpDown[] { nu_idel_time1, nu_idel_time2, nu_idel_time3, nu_idel_time4 };
            NumericUpDown[] ioutTable = new NumericUpDown[] { nu_eload_ch1, nu_eload_ch2, nu_eload_ch3, nu_eload_ch4 };
            
            test_dg.RowCount = test_dg.RowCount + 1;
            int current_row = test_dg.RowCount - 1;
            
            // add vin
            test_dg[0, current_row].Value = num_vin.Value;

            string seq_info = "";
            string meas_info = "";
            string precent_info = "";
            string idel_info = "";
            string chlevel_info = "";
            string idelTime_info = "";
            string eload_info = "";

            for (int i = 0; i < seqTable_addr.Length; i++)
            {
                int addr = (int)seqTable_addr[i].Value;
                int data = (int)seqTable_data[i].Value;

                // seq reg
                if (i == seqTable_addr.Length - 1)
                    seq_info += string.Format("{0:X2}[{1:X2}]", addr, data);
                else
                    seq_info += string.Format("{0:X2}[{1:X2}],", addr, data);

                // measure ch
                if (i == fromTable.Length - 1)
                    meas_info += fromTable[i].Text + "→" + toTable[i].Text;
                else
                    meas_info += fromTable[i].Text + "→" + toTable[i].Text + ",";

                // precent
                if (i == fromTable.Length - 1)
                    precent_info += percent_pos1[i].Text + "→" + percent_pos2[i].Text;
                else
                    precent_info += percent_pos1[i].Text + "→" + percent_pos2[i].Text + ",";

                // idel time
                addr = (int)idelTable_addr[i].Value;
                data = (int)idelTable_data[i].Value;
                if (i == fromTable.Length - 1)
                    idel_info += string.Format("{0:X2}[{1:X2}]", addr, data);
                else
                    idel_info += string.Format("{0:X2}[{1:X2}],", addr, data);

                if (i == fromTable.Length - 1)
                    chlevel_info += initLevel[i].Value.ToString();
                else
                    chlevel_info += initLevel[i].Value.ToString() + ",";

                if (i == fromTable.Length - 1)
                    idelTime_info += idelTable[i].Value.ToString();
                else
                    idelTime_info += idelTable[i].Value.ToString() + ",";

                if (i == fromTable.Length - 1)
                    eload_info += string.Format("{0}[{1}]", EloadTable[i].Text, ioutTable[i].Value);
                else
                    eload_info += string.Format("{0}[{1}],", EloadTable[i].Text, ioutTable[i].Value);

                // add seq reg setting
                if (i == seqTable_addr.Length - 1) test_dg[1, current_row].Value = seq_info;
                // add measure ch
                if (i == seqTable_addr.Length - 1) test_dg[2, current_row].Value = meas_info;
                // add precentage
                if(i == seqTable_addr.Length - 1) test_dg[3, current_row].Value = precent_info;
                // add idel time
                if (i == seqTable_addr.Length - 1) test_dg[4, current_row].Value = idel_info;
                // add initial level
                if (i == seqTable_addr.Length - 1) test_dg[5, current_row].Value = chlevel_info;
                // add eload 
                if (i == seqTable_addr.Length - 1) test_dg[6, current_row].Value = eload_info;
                // add spec
                if (i == seqTable_addr.Length - 1) test_dg[7, current_row].Value = idelTime_info;

                test_dg[8, current_row].Value = cbox_trigger.Text;
                test_dg[9, current_row].Value = CBEdge.Text;
            }
            
        }
    }
}
