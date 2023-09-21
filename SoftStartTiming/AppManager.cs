using System;
using System.Windows.Forms;
using System.Collections.Generic;

namespace SoftStartTiming
{
    public partial class SoftStartTiming
    {
        private string win_name = "Soft start v1.19_beta";


        public CheckBox[] binTable;
        public CheckBox[] ScopeChTable;


        private void SoftStartTiming_Load(object sender, EventArgs e)
        {
            this.Text = win_name;

            CbTrigger.SelectedIndex = 0;
            CBGPIO.SelectedIndex = 0;
            CBPower.Enabled = false;
            CBChannel.Enabled = false;
             
            ate_table = new TaskRun[] { _ate_delay_time, _ate_sst, _ate_delay_off };

            binTable = new CheckBox[] { CkBin1, CkBin2, CkBin3 };
            ScopeChTable = new CheckBox[] { CkCH1, CkCH2, CkCH3 };

            CBItem.SelectedIndex = 1;
            CBEdge.SelectedIndex = 0;
        }
    }


    public class OutputInfo
    {
        public string rail_name;
        public int scope_ch;                // 1 ~ 4
        public int eload_ch;                // 1 ~ 8
        public int lx_scope_ch;             // 1 ~ 4

        public bool aggressor;              // true or false

        // Eload info
        public List<double> ccm_load = new List<double>();
        public List<double> lt_l1 = new List<double>();
        public List<double> lt_l2 = new List<double>();        
        public double full_load;

        // freq reg info
        public byte freq_addr;
        public List<byte> freq_data = new List<byte>();
        public List<string> freq_des = new List<string>();

        // vout reg info
        public byte vout_addr;
        public List<byte> vout_data = new List<byte>();
        public List<string> vout_des = new List<string>();

        // VID reg info
        public byte vid_addr;
        public byte hi_code;
        public byte lo_code;

        // En reg info
        public byte en_addr;
        public byte on_data;
        public byte off_data;
    }

}
