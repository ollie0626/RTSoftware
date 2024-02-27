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
using System.Windows.Navigation;
using System.Windows.Shapes;

using InsLibDotNet;
using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;
using System.Windows.Interop;
using System.Runtime.InteropServices;

namespace Scope_Simple_tool.Pages
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class ShellView : MetroWindow
    {

        public delegate void MessageFunc(string message, string title);
        static public MessageFunc BKMessage = null;
        static public bool IsRun = false;

        public ShellView()
        {
            InitializeComponent();

            base.WindowStartupLocation = WindowStartupLocation.CenterScreen;

            MyPage.MessageExt.Instance.ShowDialog = ShowDialog;
            MyPage.MessageExt.Instance.ShowYesNo = ShowYesNo;
            BKMessage = ShowYesNo;
        }

        public void ShowDialog(string message = "", string title = "")
        {
            this.Dispatcher.Invoke((Action)(async () =>
            {
                var mySettings = new MetroDialogSettings()
                {
                    AffirmativeButtonText = "OK",
                    ColorScheme = MetroDialogColorScheme.Inverted,
                };
                var result = await this.ShowMessageAsync(title, message, MessageDialogStyle.Affirmative, mySettings);

            }));
        }

        public async void ShowYesNo(string message, string title, Action action = null)
        {
            IsRun = false;
            var mySettings = new MetroDialogSettings()
            {
                AffirmativeButtonText = "OK",
                NegativeButtonText = "Exit",
                ColorScheme = MetroDialogColorScheme.Theme
            };

            MessageDialogResult result = await this.ShowMessageAsync(title, message, MessageDialogStyle.AffirmativeAndNegative, mySettings);
            if (result == MessageDialogResult.Affirmative)
                await Task.Factory.StartNew(action);
            if (result == MessageDialogResult.Affirmative) IsRun = true;
        }

        public void ShowYesNo(string message, string title)
        {
            IsRun = false;
            this.Dispatcher.Invoke((Action)(async () =>
            {
                var MySettings = new MetroDialogSettings()
                {
                    AffirmativeButtonText = "Continue",
                    NegativeButtonText = "Cancel",
                };
                MessageDialogResult result = await this.ShowMessageAsync(title, message, MessageDialogStyle.AffirmativeAndNegative, MySettings);
                if (result == MessageDialogResult.Affirmative) IsRun = true;
                if (!MyPage.MyControl.IsPause)
                {
                    MyPage.MyControl.ManualReset.Set();
                }
            }));
        }

    }






}
