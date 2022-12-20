using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using InsLibDotNet;


namespace SDCTool
{
    public partial class Form : System.Windows.Forms.Form
    {
        private static string ver = "v1.0";
        private string win_name = "SDCTool " + ver;

        public Form()
        {
            InitializeComponent();
            this.Text = win_name;
        }

    }
}
