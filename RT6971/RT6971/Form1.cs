using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RT6971
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void AVDDH_ValueChanged(object sender, EventArgs e)
        {
            int code = (int)AVDDH.Value;
            double vol = 13.8;
            if (code > 0)
            {
                vol = (((double)(code - 1) * 1) + 138) / 10;

            }
            AVDDSL.Value = (int)AVDDH.Value;
            AVDDV.Value = (decimal)vol;
            W00.Value = (int)AVDDH.Value | ((int)W00.Value & 0xC0);
        }

        private void AVDDSL_Scroll(object sender, ScrollEventArgs e)
        {
            AVDDH.Value = AVDDSL.Value;
        }
    }
}
