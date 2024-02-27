using System;
using Stylet;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Media;
using InsLibDotNet;
using System.Threading.Tasks;
using System.Windows;
using System.IO;
using NLua;

namespace Scope_Simple_tool.Pages
{
    public class ShellViewModel : Screen
    {
        /* sub window */
        SubWindow sub = new SubWindow();
        LuaWindow luaWin = new LuaWindow();
        RTBBWindow RTBBWin = new RTBBWindow();
        MyLib myLib = new MyLib();
        MyPage.ConsoleControl cons = new MyPage.ConsoleControl();

        /* initial variable */
        static public string _InsName;
        static public string _WaveFormPath;
        static public string _FileName;
        static public string _BinFile;
        static public int _slave;

        string _Ver;
        string _IDN;
        string _ProgressStatus;
        string _PauseOrResume;
        int _index;
        int _ProMax;
        double _delay;
        bool _BTStopEn;
        bool _BTRunEn;
        bool _BTPauseEn;
        Brush _PauseColor;
        Visibility _IsConnect;
        Visibility _IsSave;
        AgilentOSC _scope;

        #region "GUI Property"
        public string InsName
        {
            get { return _InsName; }
            set { SetAndNotify(ref _InsName, value); }
        }

        public string WaveFormPath
        {
            get { return _WaveFormPath; }
            set { SetAndNotify(ref _WaveFormPath, value); }
        }

        public string FileName
        {
            get { return _FileName; }
            set { SetAndNotify(ref _FileName, value); }
        }

        public string BinFile
        {
            get { return _BinFile; }
            set { SetAndNotify(ref _BinFile, value); }
        }

        public int Slave
        {
            get { return _slave; }
            set { SetAndNotify(ref _slave, value); }
        }

        public string IDN
        {
            get { return _IDN; }
            set { SetAndNotify(ref _IDN, value); }
        }

        public string Ver
        {
            get { return _Ver; }
            set { SetAndNotify(ref _Ver, value); }
        }

        public string ProgressStatus
        {
            get { return _ProgressStatus; }
            set { SetAndNotify(ref _ProgressStatus, value); }
        }

        public string PauseOrResume
        {
            get { return _PauseOrResume; }
            set { SetAndNotify(ref _PauseOrResume, value); }
        }

        public int ProMax
        {
            get { return _ProMax; }
            set { SetAndNotify(ref _ProMax, value); }
        }

        public int Index
        {
            get { return _index; }
            set { SetAndNotify(ref _index, value); }
        }

        public Brush PauseColor
        {
            get { return _PauseColor; }
            set { SetAndNotify(ref _PauseColor, value); }
        }

        public bool BTStopEn
        {
            get { return _BTStopEn; }
            set { SetAndNotify(ref _BTStopEn, value); }
        }

        public bool BTRunEn
        {
            get { return _BTRunEn; }
            set { SetAndNotify(ref _BTRunEn, value); }
        }

        public bool BTPauseEn
        {
            get { return _BTPauseEn; }
            set { SetAndNotify(ref _BTPauseEn, value); }
        }

        public Visibility IsConnect
        {
            get { return _IsConnect; }
            set { SetAndNotify(ref _IsConnect, value); }
        }

        public Visibility IsSave
        {
            get { return _IsSave; }
            set { SetAndNotify(ref _IsSave, value); }
        }

        public double Delay
        {
            get { return _delay; }
            set { SetAndNotify(ref _delay, value); }
        }
        #endregion

        public ShellViewModel()
        {
            InitBackgroundWorker();
            InitGUI();
            MyPage.LoadIni.LoadFile();
        }

        public void FormClosed(object sender, System.EventArgs e)
        {
            sub.Close();
        }

        private void InitGUI()
        {
            base.Closed += FormClosed;

            _InsName = "TCPIP0::168.254.95.0::hislip0::INSTR";
            _ProgressStatus = "Progress Status";
            _IsConnect = Visibility.Collapsed;
            _IsSave = Visibility.Collapsed;
            _IDN = "NA";
            _WaveFormPath = @"D:\";
            _FileName = "TempFile";
            _BinFile = @"D:\Desktop\MiniLED src\ModeBin";
            _Ver = "Scope Simple Tool V3";
            _PauseColor = Brushes.Black;
            _BTStopEn = false;
            _BTPauseEn = false;
            _BTRunEn = true;
            _PauseOrResume = "Pause";
            _slave = 0x42;
            _delay = 0.5;
            Bt_Connect();
        }

        public List<string> GetBinFileList(string BinFolder)
        {
            List<string> binList = new List<string>();
            DirectoryInfo di = new DirectoryInfo(BinFolder);
            foreach (var fi in di.GetFiles("*.bin"))
            {
                Console.WriteLine(fi.Name);
                binList.Add(fi.Name);
            }
            return binList;
        }

        public bool CheckBinInFile()
        {
            if(Directory.Exists(BinFile))
            {
                List<string> list = GetBinFileList(BinFile);
                if (list.Count == 0) return false;
            }
            return true;
        }

        private void InitBackgroundWorker()
        {
            MyPage.MyControl.BW = new BackgroundWorker();
            MyPage.MyControl.BW.WorkerReportsProgress = true;
            MyPage.MyControl.BW.WorkerSupportsCancellation = true;
            MyPage.MyControl.BW.DoWork += new DoWorkEventHandler(bw_DoWork);
            MyPage.MyControl.BW.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            MyPage.MyControl.BW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            if (!MyPage.RTControl.InitBridgeBoard())
            {
                ShellView.BKMessage("\nRTBB Connected Fail !!\nDo you want to continue?", _Ver);
                SubWindow.PrintDebugMessage("RTBB Connected Fail !!");
                MyPage.MyControl.ManualReset.WaitOne();
                if (!ShellView.IsRun) return;
            }


            SubWindow.PrintDebugMessage("\r\n-------------------------------------------");
            SubWindow.PrintDebugMessage("           Run Save Waveform   ");

            if (File.Exists(BinFile))
            {
                /* do single file */
                MyPage.RTControl.I2cWriteBinfile(Slave, 0,BinFile);
                MyPage.MyControl.Sleep100Ms(2);
                SubWindow.PrintDebugMessage("Single-File : RTBB I2C Write & Save Waveform !!!");

                MyPage.MyControl.SendPercentage(1, BinFile);
                if (_scope != null)
                    _scope.SaveWaveform(WaveFormPath, Path.GetFileName(BinFile).Split('.')[0]);
            }
            else
            {
                double tempDelay = Delay * 1000;
                List<string> BinList = new List<string>();
                BinList = GetBinFileList(BinFile);

                SubWindow.PrintDebugMessage("Multi-File : RTBB I2C Write & Save Waveform !!!");
                SubWindow.PrintDebugMessage("Delay time(s) : " + Delay);

                for (int i = 0; i < BinList.Count; i++)
                {
                    if (MyPage.MyControl.IsPause) MyPage.MyControl.ManualReset.WaitOne();
                    if (MyPage.MyControl.IsStop) break;
                    MyPage.MyControl.SendPercentage(i, BinList[i]);
                    string file = BinFile + @"\" + BinList[i];

                    SubWindow.PrintDebugMessage("File Name :" + BinList[i]);

                    if (Path.HasExtension(file)) MyPage.RTControl.I2cWriteBinfile(Slave, 0, file);

                    if(Delay != 0) System.Threading.Thread.Sleep((int)tempDelay);

                    if (_scope != null)
                        _scope.SaveWaveform(WaveFormPath, Path.GetFileNameWithoutExtension(BinList[i]));
                }
            }
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            Index = e.ProgressPercentage;
            ProgressStatus = "File Name : " + Path.GetFileName(e.UserState.ToString());
            Console.WriteLine(Path.GetFileName(e.UserState.ToString()));
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (ShellView.IsRun) MyPage.MessageExt.Instance.ShowDialog("Catch Waveform Finish !!!", _Ver);
            ProgressStatus = " Test Finish !!!";

            SubWindow.PrintDebugMessage("-------------------------------------------");
            SubWindow.PrintDebugMessage("           Test Finish !!!   ");


            MyPage.MyControl.ManualReset.Reset();
            Console.WriteLine("index = {0}, promax = {1}", Index, ProMax);
            MyPage.MyControl.IsStop = false;
            BTPauseEn = false;
            BTStopEn = false;
            BTRunEn = true;
        }

        public async void Bt_Connect()
        {
            if (IsConnect == Visibility.Visible) return;
            IsConnect = Visibility.Visible;
            await TaskRun("Connect");
            IsConnect = Visibility.Collapsed;
        }

        public async void Bt_Save()
        {
            if (IsSave == Visibility.Visible || _scope.doQueryIDN() == "unlink Ins.")
            {
                MyPage.MessageExt.Instance.ShowDialog("Please Connect Scope !!!", _Ver);
                SubWindow.PrintDebugMessage("Please Connect Scope !!!");
                return;
            }

            IsSave = Visibility.Visible;
            await TaskRun("Save");
            IsSave = Visibility.Collapsed;
        }

        public void Bt_SaveInit()
        {
            MyPage.MessageExt.Instance.ShowDialog("Save GUI Setting !!!", _Ver);
            MyPage.LoadIni.SaveFile();
        }

        public void Bt_Run()
        {
            if (!Directory.Exists(BinFile))
            {
                if (!File.Exists(BinFile))
                {
                    MyPage.MessageExt.Instance.ShowDialog("\nPlease check your path !!!!", _Ver);
                    SubWindow.PrintDebugMessage("Please check your path !!!!");
                    return;
                }
            }

            if (!CheckBinInFile())
            {
                MyPage.MessageExt.Instance.ShowDialog("\nThe Folder no any Bin file !!!!", _Ver);
                SubWindow.PrintDebugMessage("The Folder no any Bin file !!!!");
                return;
            }

            if (_scope == null)
            {
                MyPage.MessageExt.Instance.ShowDialog("\nScope Connect Fail !!!!", _Ver);
                SubWindow.PrintDebugMessage("Scope Connect Fail !!!!");
                return;
            }

            if (_scope.doQueryIDN() == "")
            {
                MyPage.MessageExt.Instance.ShowDialog("\nScope Connect Fail !!!!", _Ver);
                SubWindow.PrintDebugMessage("Scope Connect Fail !!!!");
                return;
            }

            MyPage.MyControl.IsPause = false;
            MyPage.MyControl.IsStop = false;

            BTPauseEn = true;
            BTStopEn = true;
            BTRunEn = false;

            if (File.Exists(BinFile))
            {
                ProMax = 1;
            }
            else
            {
                ProMax = GetBinFileList(BinFile).Count - 1;
            }

            MyPage.MyControl.BW.RunWorkerAsync();
        }

        public void Bt_Pause()
        {
            ProgressStatus = " Pause !!!";
            if (!MyPage.MyControl.IsPause)
            {
                PauseOrResume = "Resume";
                SubWindow.PrintDebugMessage("Resume !!!!");
                MyPage.MyControl.IsPause = true;
                MyPage.MyControl.ManualReset.Reset();
                PauseColor = Brushes.Red;
            }
            else
            {
                PauseOrResume = "Pause";
                SubWindow.PrintDebugMessage("Pause !!!!");
                MyPage.MyControl.IsPause = false;
                MyPage.MyControl.ManualReset.Set();
                PauseColor = Brushes.Black;
            }
        }

        public void Bt_Stop()
        {
            SubWindow.PrintDebugMessage("Stop !!!");
            ProgressStatus = " Stop !!!";
            PauseOrResume = "Pause";
            MyPage.MyControl.IsStop = true;
            BTPauseEn = false;
            BTStopEn = false;
            BTRunEn = true;
            if (MyPage.MyControl.IsPause)
            {
                PauseOrResume = "Pause";
                MyPage.MyControl.IsPause = false;
                MyPage.MyControl.ManualReset.Set();
                PauseColor = Brushes.Black;
            }
        }

        public void Bt_DebugWindow()
        {
            if(SubWindow.IsOn)
            {
                SubWindow.view.DebugMessage = "Here show debug message !!!!" + Environment.NewLine;
                SubWindow.IsShow ^= true;

                if (SubWindow.IsShow)
                    sub.Show();
                else
                    sub.Hide();   
            }
            else
            {
                SubWindow.IsOn = true;
                SubWindow.view.DebugMessage = "Here show debug message !!!!" + Environment.NewLine;
                sub = new SubWindow();
                sub.Show();
            }
        }

        public void Bt_LuaWindow()
        {
            if (LuaWindow.IsOn)
            {
                LuaWindow.IsShow ^= true;
                if (LuaWindow.IsShow)
                {
                    luaWin.Show();
                    MyPage.ConsoleControl.Show_AppConsole();
                }
                else
                {
                    luaWin.Hide();
                    MyPage.ConsoleControl.Hide_AppConsole();
                }
                    
            }
            else
            {
                LuaWindow.IsOn = true;
                luaWin = new LuaWindow();
                luaWin.Show();
                MyPage.ConsoleControl.Show_AppConsole();
            }
        }

        public void Bt_RTBBWindow()
        {
            if (RTBBWindow.IsOn)
            {
                RTBBWindow.IsShow ^= true;
                if (RTBBWindow.IsShow)
                {
                    RTBBWin.Show();
                }
                else
                {
                    RTBBWin.Hide();
                }

            }
            else
            {
                RTBBWindow.IsOn = true;
                RTBBWin = new RTBBWindow();
                RTBBWin.Show();
            }
        }

        public Task<int> TaskRun(string symbol)
        {
            if (symbol == "Connect") return Task.Factory.StartNew(() => this.Task_Con2Device());
            else if (symbol == "Save") return Task.Factory.StartNew(() => this.Task_SaveWaveform());
            return null;
        }

        public int Task_Con2Device()
        {
            SubWindow.PrintDebugMessage("Scope Connecting ... ");
            System.Threading.Thread.Sleep(3000);
            _scope = new AgilentOSC(InsName);
            IDN = _scope.doQueryIDN();

            if (IDN == "unlink Ins.")
                SubWindow.PrintDebugMessage("Connect Scope Failed !!!");
            else
                SubWindow.PrintDebugMessage("Connect Scope Success !!!");

            return 0;
        }

        public int Task_SaveWaveform()
        {
            SubWindow.PrintDebugMessage("Save Scope Waveform !!! WaveformPath : " + _WaveFormPath + "File Name : " + _FileName);
            _scope.SaveWaveform(_WaveFormPath, _FileName);
            return 0;
        }

    }
}
