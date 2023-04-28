using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SoftStartTiming
{
    public partial class Main : Form
    {
        public Main()
        {
            InitializeComponent();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CBItem_SelectedIndexChanged(object sender, EventArgs e)
        {
            Form form;
            switch(CBItem.SelectedIndex)
            {
                case 0:
                    form = new SoftStartTiming();
                    form.ShowDialog();
                    break;
                case 1:
                    form = new CrossTalk();
                    form.ShowDialog();
                    break;
                case 2:
                    form = new VIDIO();
                    form.ShowDialog();
                    break;
            }
        }
    }
}
