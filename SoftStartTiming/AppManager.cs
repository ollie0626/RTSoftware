using System;
using System.Windows.Forms;

namespace SoftStartTiming
{
    public partial class SoftStartTiming
    {
        private string win_name = "Soft start v1.0";

        private void SoftStartTiming_Load(object sender, EventArgs e)
        {
            this.Text = win_name;

            ate_table = new TaskRun[] { _ate_sst };
        }
    }
}
