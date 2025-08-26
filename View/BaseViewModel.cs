using ReverseGeoCoding.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace GTFS_Feed_Conversion.View
{
    public class BaseViewModel : INotifyPropertyChanged
    {

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        public void NotifyIt(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
        #endregion

        #region Window Open :: common class

        public ICommand ShowWindow
        {
            get { return new RelayCommand(DisplayIt); }
        }

        public void DisplayIt(object viewModelName)
        {
            if (viewModelName != null)
            {
                var win = new Window();
                win.Content = viewModelName;
                win.Show();
            }
        }

        #endregion        

        public string _Selectedfilepath { get; set; }
        public string Selectedfilepath
        {
            get { return _Selectedfilepath; }
            set
            {
                if (_Selectedfilepath != value)
                {
                    _Selectedfilepath = value;
                    NotifyIt("Selectedfilepath");
                    GlobalClass.InputFilepath = _Selectedfilepath;
                }
            }
        }
        public string _Outputfilepath { get; set; }
        public string Outputfilepath
        {
            get { return _Outputfilepath; }
            set
            {
                if (_Outputfilepath != value)
                {
                    _Outputfilepath = value;
                    NotifyIt("Outputfilepath");
                    GlobalClass.OutputFilepath = _Outputfilepath;
                }
            }
        }
    }

 

    #region  RelayCommand
    public class RelayCommand : ICommand
    {
        private Action<object> execute;
        private bool canExecute;
        public event EventHandler CanExecuteChanged;

        public RelayCommand(Action<object> whatToExecute, bool whenToExecute = true)
        {
            this.execute = whatToExecute;
            this.canExecute = whenToExecute;
        }


        public bool CanExecute(object parameter)
        {
            return canExecute;
        }

        public void Execute(object parameter)
        {

            this.execute(parameter);

        }
    }
    #endregion
}
