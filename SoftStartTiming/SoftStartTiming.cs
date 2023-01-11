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

namespace SoftStartTiming
{
    public partial class SoftStartTiming : Form
    {
        ParameterizedThreadStart p_thread;
        Thread ATETask;

        // test item
        ATE_SoftStartTiming _ate_sst;


        public SoftStartTiming()
        {
            InitializeComponent();
        }

        private void BTSelectBinPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog folderBrowser = new FolderBrowserDialog();
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                tbBin.Text = folderBrowser.SelectedPath;
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
                case 0: InsControl._scope = new AgilentOSC(res); break;
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
            Button bt = (Button)sender;
            int idx = bt.TabIndex;

            await ConnectTask(tb_osc.Text, 0);
            await ConnectTask(tb_power.Text, 1);
            await ConnectTask(tb_eload.Text, 2);
            await ConnectTask(tb_daq.Text, 3);
            await ConnectTask(tb_chamber.Text, 4);

            MyLib.Delay1s(1);

            if (InsControl._scope.InsState()) 
                led_osc.BackColor = Color.LightGreen;
            else 
                led_osc.BackColor = Color.Red;

            if (InsControl._power.InsState())
                led_power.BackColor = Color.LightGreen;
            else
                led_power.BackColor = Color.Red;

            if (InsControl._eload.InsState())
                led_eload.BackColor = Color.LightGreen;
            else
                led_eload.BackColor = Color.Red;

            if (InsControl._34970A.InsState())
                led_daq.BackColor = Color.LightGreen;
            else
                led_daq.BackColor = Color.Red;

            if (InsControl._chamber.InsState())
                led_chamber.BackColor = Color.LightGreen;
            else
                led_chamber.BackColor = Color.Red;
        }

        private void BTRun_Click(object sender, EventArgs e)
        {
            BTRun.Enabled = false;
            try
            {
                
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error Message:" + ex.Message);
                Console.WriteLine("StackTrace:" + ex.StackTrace);
                MessageBox.Show(ex.StackTrace);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opendlg = new OpenFileDialog();

            if(opendlg.ShowDialog() == DialogResult.OK)
            {
                string file_name = opendlg.FileName;
                StreamReader sr = new StreamReader(file_name);
                string line;
                List<byte> temp = new List<byte>();
                line = sr.ReadLine();
                while(line != null)
                {
                    Console.WriteLine(line);
                    string[] arr = line.Split('\t');
                    line = sr.ReadLine();
                    temp.Add(Convert.ToByte(arr[1], 16));
                }
                sr.Close();

                FileStream myFile = new FileStream(@"D:\123.bin", FileMode.OpenOrCreate);
                BinaryWriter bwr = new BinaryWriter(myFile);
                bwr.Write(temp.ToArray(), 0, temp.Count);
                bwr.Close();
                myFile.Close();
            }
        }
    }
}
