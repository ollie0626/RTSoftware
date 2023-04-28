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

namespace SoftStartTiming
{
    public partial class VIDIO : Form
    {
        string win_name = "VIDIO v1.0";
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        TaskRun[] ate_table;
        string[] tempList;
        int SteadyTime;

        

        private void InitDG()
        {
            // lpm mode
            Column1.Items.Add("0");
            Column1.Items.Add("1");

            // g1 logic
            Column2.Items.Add("0");
            Column2.Items.Add("1");

            // g2 logic
            Column3.Items.Add("0");
            Column3.Items.Add("1");

            // after
            Column5.Items.Add("0");
            Column5.Items.Add("1");

            Column6.Items.Add("0");
            Column6.Items.Add("1");

            Column7.Items.Add("0");
            Column7.Items.Add("1");
        }

        public VIDIO()
        {
            InitializeComponent();
            InitDG();
            this.Name = win_name;

            //RTBBControl dev = new RTBBControl();
            //dev.BoadInit();
            //dev.GPIOnState(0, true);
            //dev.GPIOnState(0, false);
        }

        private void BT_Add_Click(object sender, EventArgs e)
        {
            dataGridView1.RowCount = dataGridView1.RowCount + 1;
        }

        private void BT_Sub_Click(object sender, EventArgs e)
        {
            if(dataGridView1.RowCount - 1 > 0)
                dataGridView1.RowCount = dataGridView1.RowCount - 1;
            else
                dataGridView1.RowCount = 0;
        }

        private void test_parameter_copy()
        {
            test_parameter.vidio.lpm_sel.Clear();
            test_parameter.vidio.g1_sel.Clear();
            test_parameter.vidio.g2_sel.Clear();
            test_parameter.vidio.vout_list.Clear();

            test_parameter.vidio.lpm_sel_af.Clear();
            test_parameter.vidio.g1_sel_af.Clear();
            test_parameter.vidio.g2_sel_af.Clear();
            test_parameter.vidio.vout_list_af.Clear();

            test_parameter.VinList = tb_vinList.Text.Split(',').Select(double.Parse).ToList();
            test_parameter.IoutList = tb_iout.Text.Split(',').Select(double.Parse).ToList();

            for(int i = 0; i < dataGridView1.RowCount; i++)
            {
                test_parameter.vidio.lpm_sel.Add((int)dataGridView1[0, i].Value);
                test_parameter.vidio.g1_sel.Add((int)dataGridView1[1, i].Value);
                test_parameter.vidio.g2_sel.Add((int)dataGridView1[2, i].Value);
                test_parameter.vidio.vout_list.Add((double)dataGridView1[3, i].Value);

                test_parameter.vidio.lpm_sel_af.Add((int)dataGridView1[4, i].Value);
                test_parameter.vidio.g1_sel_af.Add((int)dataGridView1[5, i].Value);
                test_parameter.vidio.g2_sel_af.Add((int)dataGridView1[6, i].Value);
                test_parameter.vidio.vout_list_af.Add((double)dataGridView1[7, i].Value);
            }
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
                    ATETask.Start(0);
                }
                else
                {
                    p_thread = new ParameterizedThreadStart(Run_Single_Task);
                    ATETask = new Thread(p_thread);
                    ATETask.Start(0);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
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
    }


    public class VIDIO_parameter
    {
        public List<int> lpm_sel = new List<int>();
        public List<int> g1_sel = new List<int>();
        public List<int> g2_sel = new List<int>();
        public List<double> vout_list = new List<double>();

        public List<int> lpm_sel_af = new List<int>();
        public List<int> g1_sel_af = new List<int>();
        public List<int> g2_sel_af = new List<int>();
        public List<double> vout_list_af = new List<double>();
    }


}
