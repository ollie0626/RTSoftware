using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.ComponentModel;
using System.Collections.ObjectModel;

namespace Scope_Simple_tool.Pages
{
    public class DebugModuel 
    {
        public string _debugMessage { get; set; }

        public string _title { get; set; }
    }

    public class SubWindowViewModel : INotifyPropertyChanged
    {
        public string _debugMessage;
        public string _title;
        public bool _isClose;

        public event PropertyChangedEventHandler PropertyChanged;

        private void RaisePropertyChanged(string propertyName)
        {
            PropertyChangedEventHandler handler = PropertyChanged;

            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        public SubWindowViewModel()
        {
            DebugMessage = "Here show debug message !!!!" + Environment.NewLine;
            Title = "Debug Window";
        }

        public string DebugMessage
        {
            get { return _debugMessage; }
            set { _debugMessage = value; RaisePropertyChanged("DebugMessage"); }
        }

        public string Title
        {
            get { return _title; }
            set { _title = value; RaisePropertyChanged("Title"); }
        }



    }

}

