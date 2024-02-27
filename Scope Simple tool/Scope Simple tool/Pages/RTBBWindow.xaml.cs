using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

using MahApps.Metro.Controls;



namespace Scope_Simple_tool.Pages
{
    /// <summary>
    /// RTBBWindow.xaml 的互動邏輯
    /// </summary>
    public partial class RTBBWindow : MetroWindow
    {

        static public bool IsOn = true;
        static public bool IsShow = false;



        public RTBBWindow()
        {
            InitializeComponent();
            base.Closed += FormCloased;
        }


        public void FormCloased(object sender, System.EventArgs e)
        {
            IsOn = false;
        }
    }
}
