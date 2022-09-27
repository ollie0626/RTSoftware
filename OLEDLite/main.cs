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
using InsLibDotNet;
using System.Threading;


namespace OLEDLite
{
    public partial class main : MaterialSkin.Controls.MaterialForm
    {
        private string win_name = "OLED sATE tool v1.0";
        private readonly MaterialSkinManager materialSkinManager;

        private ParameterizedThreadStart p_thread;
        private Thread ATETask;
        private ATE_TDMA _ate_tdma = new ATE_TDMA();
        private TaskRun[] ate_table;

        public main()
        {
            InitializeComponent();
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            //materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
            materialTabSelector1.Width = this.Width;
            materialTabSelector1.Height = 25;
            this.Text = win_name;

            //this.WindowState = FormWindowState.Maximized;
            //GUI_Design();
            materialTabControl1.SelectedIndex = 1;
            ATEItemInit();
        }


        private void ATEItemInit()
        {
            ate_table = new TaskRun[] { _ate_tdma };
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
            InsControl._funcgen.CH1_Frequency((double)(nu_Freq.Value * 1000));
            InsControl._funcgen.CH1_DutyCycle((double)nu_duty.Value);
            InsControl._funcgen.CH1_LoadImpedanceHiz();
            InsControl._funcgen.SetCH1_TrTfFunc((double)nu_Tr.Value, (double)nu_Tf.Value);
            InsControl._funcgen.CHl1_HiLevel(hi);
            InsControl._funcgen.CH1_LoLevel(lo);
            InsControl._funcgen.CH1_On();
        }

        private void test_parameter_copy()
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
            test_parameter.eload_en = new bool[4] { ck_ch1_en.Checked, 
                                                    ck_ch2_en.Checked, 
                                                    ck_ch3_en.Checked, 
                                                    ck_ch4_en.Checked };
            test_parameter.eload_iout = new double[4] { (double)nu_load1.Value,
                                                        (double)nu_load2.Value,
                                                        (double)nu_load3.Value,
                                                        (double)nu_load4.Value };

            test_parameter.swireList.Clear();
            for (int i = 0; i < swireTable.RowCount; i++)
            {
                test_parameter.swireList.Add((string)swireTable[0, i].Value);
            }
            test_parameter.swire_20 = true;
        }


        private void bt_run_Click(object sender, EventArgs e)
        {
            try
            {
                test_parameter_copy();
                p_thread = new ParameterizedThreadStart(Run_Single_Task);
                ATETask = new Thread(p_thread);
                ATETask.Start(cb_item.SelectedIndex);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
            }
        }

        private void Run_Single_Task(object idx)
        {
            ate_table[(int)idx].ATETask();
        }

        private void nu_swire_num_ValueChanged(object sender, EventArgs e)
        {
            swireTable.RowCount = (int)nu_swire_num.Value;
        }
    }
}
