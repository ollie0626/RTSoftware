using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Windows.Input;
using NLua;
using Scope_Simple_tool.MyPage;

namespace Scope_Simple_tool.Pages
{
    public class LuaWindowViewModel : INotifyPropertyChanged
    {
        private static Lua lua = new Lua();
        private string _Luapath;
        private string _LuaScript_str;
        private ICommand _openCommand;
        private ICommand _runCommand;
        private ICommand _clearCommand;
        private ICommand _runstringCommand;
        private ICommand _clearLuaWindownCommand;

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public string LuaPath
        {
            get { return _Luapath; }
            set { _Luapath = value; RaisePropertyChanged("LuaPath"); }
        }

        public string LuaScript_str
        {
            get { return _LuaScript_str; }
            set { _LuaScript_str = value; RaisePropertyChanged("LuaScript_str"); }
        }

        public bool CanExecute
        {
            get
            {
                // check if executing is allowed, i.e., validate, check if a process is running, etc. 
                return true;
            }
        }

        public ICommand OpenCommand
        {
            get { return _openCommand ?? (_openCommand = new MyPage.CommandHandler(() => OpenFile(), () => CanExecute)); }
        }

        public ICommand RunCommand
        {
            get { return _runCommand ?? (_runCommand = new MyPage.CommandHandler(() => RunLuaScript(), () => CanExecute)); }
        }

        public ICommand ClearCommand
        {
            get { return _clearCommand ?? (_clearCommand = new MyPage.CommandHandler(() => ClearConsole(), () => CanExecute)); }
        }

        public ICommand RunstringCommand
        {
            get { return _runstringCommand ?? (_runstringCommand = new MyPage.CommandHandler(() => RunLuaScript_string(), () => CanExecute)); }
        }

        public ICommand ClearLuaWindowCommand
        {
            get { return _clearLuaWindownCommand ?? (_clearLuaWindownCommand = new MyPage.CommandHandler(() => LuaWindowClear(), () => CanExecute)); }
        }


        public LuaWindowViewModel()
        {
            lua.LoadCLRPackage();
            LuaRegister();
        }

        public static void LuaWindowMessage(string str)
        {
            Console.WriteLine(str);
        }

        //public static void LuaScriptPrint(string str)
        //{
        //    Console.WriteLine(str);
        //}

        public void OpenFile()
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            openFileDlg.Filter = "lua script(*.lua)|*.lua";
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true)
            {
                LuaPath = openFileDlg.FileName;
            }
        }

        public void LuaRegister()
        {
            lua.RegisterFunction("I2cWrite", null, typeof(RTControl).GetMethod("I2c_SingleWrite"));
            lua.RegisterFunction("I2cWriteBin", null, typeof(RTControl).GetMethod("I2cWriteBinfile"));
            lua.RegisterFunction("InitBridgeBoard", null, typeof(RTControl).GetMethod("InitBridgeBoard"));
            //lua.RegisterFunction("print", null, typeof(LuaWindowViewModel).GetMethod("LuaScriptPrint"));
        }

        public void RunLuaScript()
        {
            try
            {
                lua.DoFile(LuaPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Trace : " + ex.StackTrace);
                Console.WriteLine("Message : " + ex.Message);
            }
        }

        public void ClearConsole()
        {
            MyPage.ConsoleControl.Clear_AppConsole();
        }


        public void RunLuaScript_string()
        {
            try
            {
                lua.DoString(LuaScript_str);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Trace : " + ex.StackTrace);
                Console.WriteLine("Message : " + ex.Message);
            }
        }

        public void LuaWindowClear()
        {
            LuaScript_str = string.Empty;
            MyPage.ConsoleControl.Clear_AppConsole();
        }


    }

}
