using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using MaterialSkin;
using MaterialSkin.Controls;
using InsLibDotNet;


namespace OLEDLite
{
    public partial class main : MaterialSkin.Controls.MaterialForm
    {
        private string win_name = "OLED sATE tool v1.0";
        private readonly MaterialSkinManager materialSkinManager;
        public main()
        {
            InitializeComponent();
            materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            //materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;
            materialSkinManager.ColorScheme = new ColorScheme(Primary.BlueGrey800, Primary.BlueGrey900, Primary.BlueGrey500, Accent.LightBlue200, TextShade.WHITE);
            materialTabSelector1.Width = this.Width;
            materialTabSelector1.Height = 25;
            this.Text = win_name;

            //this.WindowState = FormWindowState.Maximized;
            //GUI_Design();
            materialTabControl1.SelectedIndex = 1;
        }

        private void main_Resize(object sender, EventArgs e)
        {
            materialTabSelector1.Width = this.Width;
        }

        private void bt_connect_Click(object sender, EventArgs e)
        {

        }

        private void bt_scanIns_Click(object sender, EventArgs e)
        {
            string[] ins_list = ViCMD.ScanIns();
        }
    }
}
