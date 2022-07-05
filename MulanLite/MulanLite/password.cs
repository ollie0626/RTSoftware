using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Sunny.UI;

namespace MulanLite
{
    public partial class password : Form
    {
        main handle;

        public password(Form from)
        {
            InitializeComponent();
            handle = (main)from;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if(textBox1.Text == "03155")
            {
                handle.uiTabControl1.TabPages.Add(handle.tabPage5);
                handle.uiTabControl1.SelectedIndex = 4;
                this.Close();
            }
            else
            {
                MessageBox.Show("Password error !!!", "Mulan-Lit Tool", MessageBoxButtons.OK);
            }
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyValue == 13)
            {
                button1_Click(null, null);
            }
        }
    }
}
