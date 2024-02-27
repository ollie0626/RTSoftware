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
    /// SubWindow.xaml 的互動邏輯
    /// </summary>
    public partial class SubWindow : MetroWindow
    {

        static public SubWindowViewModel view;
        static public bool IsOn = true;
        static public bool IsShow = false;

        public SubWindow()
        {
            InitializeComponent();
            base.Closed += FormCloased;
            view = (SubWindowViewModel)base.DataContext;
        }

        public static void PrintDebugMessage(string mess)
        {
            view.DebugMessage += mess + Environment.NewLine;
        }

        public void FormCloased(object sender, System.EventArgs e)
        {
            IsOn = false;
        }


    }
}
