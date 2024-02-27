using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Windows.Input;
using System.IO;

namespace Scope_Simple_tool.Pages
{
    public class RTBBWindowViewModel : INotifyPropertyChanged
    {

        private byte[] _writeBuffer;
        private byte[] _readBuffer = new byte[255];


        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public RTBBWindowViewModel()
        {
            _slave = 0x62;
            _addr = 0x00;
            _wData = "62, 77, 88";
            _binEn = false;
        }

        private string      _filename;
        private byte        _slave;
        private byte        _addr;
        private string      _wData;
        private bool        _binEn;
        private ICommand    _openCommand;
        private ICommand    _writeCommand;
        private ICommand    _readCommand;
        private string      _log;
        private int         _len;


        public bool CanExecute
        {
            get
            {
                // check if executing is allowed, i.e., validate, check if a process is running, etc. 
                return true;
            }
        }

        public string FileName
        {
            get { return _filename; }
            set { _filename = value; RaisePropertyChanged("FileName"); }
        }

        public byte Slave
        {
            get { return _slave; }
            set { _slave = value; RaisePropertyChanged("Slave"); }
        }

        public byte Addr
        {
            get { return _addr; }
            set { _addr = value; RaisePropertyChanged("Addr"); }
        }

        public string WData
        {
            get { return _wData; }
            set { _wData = value; RaisePropertyChanged("WData"); }
        }

        public bool BinEn
        {
            get { return _binEn; }
            set { _binEn = value; RaisePropertyChanged("BinEn"); }
        }

        public string Log
        {
            get { return _log; }
            set { _log = value; RaisePropertyChanged("Log"); }
        }

        public int Length
        {
            get { return _len; }
            set { _len = value; RaisePropertyChanged("Length"); }
        }


        public ICommand OpenCommand
        {
            get { return _openCommand ?? (_openCommand = new MyPage.CommandHandler(() => OpenFile(), () => CanExecute)); }
        }

        public ICommand WriteCommand
        {
            get { return _writeCommand ?? (_writeCommand = new MyPage.CommandHandler(() => WriteData(), () => CanExecute)); }
        }

        public ICommand ReadCommand
        {
            get { return _readCommand ?? (_readCommand = new MyPage.CommandHandler(() => ReadData(), () => CanExecute)); }
        }


        public void OpenFile()
        {
            Microsoft.Win32.OpenFileDialog openFileDlg = new Microsoft.Win32.OpenFileDialog();
            openFileDlg.Filter = "Bin File(*.bin)|*.bin";
            Nullable<bool> result = openFileDlg.ShowDialog();
            if (result == true)
            {
                FileName = openFileDlg.FileName;
                SubWindow.PrintDebugMessage("RTBB Window Open Bin file ok !!!");
            }
            else
            {
                SubWindow.PrintDebugMessage("RTBB Window Open Bin file fail !!!");
            }
        }

        public void WriteData()
        {
            if(BinEn)
            {
                /* write bin file path */
                if (File.Exists(FileName))
                    MyPage.RTControl.I2cWriteBinfile(Slave, 0, FileName);
                else
                    SubWindow.PrintDebugMessage("Bin File isn't exisit");
            }
            else
            {
                /* write gui path */
                string[] buf = WData.Split(',');
                _writeBuffer = new byte[buf.Length];
                for(int i = 0; i < buf.Length; i++)
                {
                    _writeBuffer[i] = Convert.ToByte(buf[i].Trim(), 16);
                }

            }
        }

        public void ReadData()
        {

        }



    }
}
