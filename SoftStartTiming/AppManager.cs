using System;
using System.Windows.Forms;

namespace SoftStartTiming
{
    public partial class SoftStartTiming
    {
        private string win_name = "Soft start v1.0";


        public CheckBox[] binTable;
        public CheckBox[] ScopeChTable;


        private void SoftStartTiming_Load(object sender, EventArgs e)
        {
            this.Text = win_name;

            CbTrigger.SelectedIndex = 0;
            ate_table = new TaskRun[] { _ate_sst };

            binTable = new CheckBox[] { CkBin1, CkBin2, CkBin3 };
            ScopeChTable = new CheckBox[] { CkCH1, CkCH2, CkCH3 };
        }
    }
}
