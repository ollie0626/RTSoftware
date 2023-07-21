using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using RTBBLibDotNet;
using System.Threading;

namespace SoftStartTiming
{
    public partial class LoadTransient : Form
    {
        string win_name = "Load Transient v1.0";
        RTBBControl RTDev = new RTBBControl();
        ParameterizedThreadStart p_thread;
        Thread ATETask;
        TaskRun[] ate_table;
        string[] tempList;
        int SteadyTime;


        public LoadTransient()
        {
            InitializeComponent();
            this.Text = win_name;
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

        private void BTPause_Click(object sender, EventArgs e)
        {

        }

        private void BTStop_Click(object sender, EventArgs e)
        {

        }

        private void BTRun_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            ProcessStartInfo psi = new ProcessStartInfo();
            psi.Arguments = "/im EXCEL.EXE /f";
            psi.FileName = "taskkill";
            Process p = new Process();
            p.StartInfo = psi;
            p.Start();
        }
    }

    public class LoadTransient_parameter
    {
        // function gen parameter
        public double freq;
        public double duty;

        // eload parameter
        public double T1;
        public double T2;

        public double Tr;
        public double Tf;
        public double zoom_in_ratio;

        // dynamic loading setting
        public List<double> hi_current = new List<double>();
        public List<double> lo_current = new List<double>();

        // for evb current offset
        public double offset;

        public bool eload_dev_sel;
        public double gain;
    }
}
