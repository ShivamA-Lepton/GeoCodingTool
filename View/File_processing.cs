using GTFS_Feed_Conversion.View;
using ReverseGeoCoding.Common;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Input;

namespace ReverseGeoCoding.View
{
    public class File_processing : ProgressVal
    {
        System.Threading.Thread ProgBar = null;
        bool Flag = true;

        public File_processing()
        {
            GlobalClass.ChangeForm.ChangeEvent += new GlobalClass.FormSelectIndex(ClosePanel);
        }

        #region for refreshing the value
        public void ClosePanel(int e)
        {
            try
            {
                if (e == 5)
                {
                    TaskStatus = ProgressVal.ActionType.Completed;
                }
            }
            catch (Exception ex) { }

        }
        #endregion

        public ICommand OpenBrowseWindow
        {
            get { return new RelayCommand(AddOpenFileInput); }
        }
        public ICommand OpenOutputWindow
        {
            get { return new RelayCommand(OpenOutputFilepath); }
        }
        public void AddOpenFileInput(object paramter)
        {
            try
            {
                OpenFileDialog Csvfileath = new OpenFileDialog();
                Csvfileath.Multiselect = true;
                Csvfileath.Filter = "*.xlsx*|*.*";
                Csvfileath.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                if (Csvfileath.ShowDialog() == DialogResult.OK)
                {
                    string ext = System.IO.Path.GetExtension(Csvfileath.FileName.ToString());
                    if ((ext.Equals(".xlsx")) | (ext.Equals(".xls")) | (ext.Equals(".html")))
                    {
                        GlobalClass.InputFilepath = Csvfileath.FileName.ToString();
                        GlobalClass.ChangeForm.OnChangeForm(2);
                    }
                }
            }
            catch { }
        }

        /// <summary>
        /// this method call after all the files are loaded from xlsx or csv
        /// </summary>
        /// <param name="paramter"></param>
        public void OpenOutputFilepath(object paramter)
        {
            try
            {
                System.Windows.Forms.FolderBrowserDialog OutputFile = new System.Windows.Forms.FolderBrowserDialog();
                if (OutputFile.ShowDialog() == DialogResult.OK)
                {
                    GlobalClass.OutputFilepath = OutputFile.SelectedPath.ToString()+"/";
                    GlobalClass.ChangeForm.OnChangeForm(2);
                }
            }
            catch { }
        }

    }

    #region Create Model Entities
    public class ProgressVal : BaseViewModel
    {

        public ProgressVal()
        {
            Thread myThread = new Thread(new ThreadStart(UpdateProgressBarStatus));
            myThread.Start();
        }

        private ActionType _StartProcess { get; set; }
        public ActionType TaskStatus
        {
            get { return _StartProcess; }
            set
            {
                if (_StartProcess != value)
                {
                    _StartProcess = value;
                    NotifyIt("TaskStatus");
                }
                switch (_StartProcess)
                {
                    case ActionType.Start:
                        {
                            ProgressBarMaxValue = 100;
                            ProgressBarSleepTime = 1000;
                            
                            StatusImage = "Images\\loading_2.gif";
                            StatusBarVisibility = "visible";
                            ProcessBarMessage = GlobalClass.messagevalue;
                            ProgressBarValue = GlobalClass.progressVaue;
                            if (ProgressBarValue != 0)
                            {
                                objStopWatch.Start();
                                TimeSpan objTimeSpan = TimeSpan.FromMilliseconds(objStopWatch.ElapsedMilliseconds);
                                ProcessingTime = String.Format(CultureInfo.CurrentCulture, "{0:00}:{1:00}:{2:00}", objTimeSpan.Hours, objTimeSpan.Minutes, objTimeSpan.Seconds);
                            }
                            if (ProgressBarValue == 100)
                            {
                                objStopWatch.Reset();
                            }
                        }
                        break;
                    case ActionType.Stop:
                        {
                            objStopWatch.Stop();
                            StatusBarVisibility = "collapsed";
                            ProgressBarValue = 0;
                            ProcessingTime = "00:00:00";
                            ProcessBarMessage = string.Empty;
                        }
                        break;

                    case ActionType.Completed:
                        {
                            objStopWatch.Stop();
                            StatusImage = "Images\\done.png";
                            ProcessBarMessage = "Processed successfully.";
                        }
                        break;
                }
            }
        }

        private string _StatusBarVisibility { get; set; }
        public string StatusBarVisibility
        {
            get { return _StatusBarVisibility; }
            set
            {
                if (_StatusBarVisibility != value)
                {
                    _StatusBarVisibility = value;
                    NotifyIt("StatusBarVisibility");
                }
            }
        }

        private int _ProgressBarMaxValue { get; set; }
        public int ProgressBarMaxValue
        {
            get { return _ProgressBarMaxValue; }
            set
            {
                if (_ProgressBarMaxValue != value)
                {
                    _ProgressBarMaxValue = value;
                    NotifyIt("ProgressBarMaxValue");
                }
            }
        }

        private double _ProgressBarValue { get; set; }
        public double ProgressBarValue
        {
            get { return _ProgressBarValue; }
            set
            {
                if (_ProgressBarValue != value)
                {
                    _ProgressBarValue = value;
                    NotifyIt("ProgressBarValue");
                }
            }
        }

        private int _ProgressBarSleepTime { get; set; }
        public int ProgressBarSleepTime
        {
            get { return _ProgressBarSleepTime; }
            set
            {
                if (_ProgressBarSleepTime != value)
                {
                    _ProgressBarSleepTime = value;
                    NotifyIt("ProgressBarSleepTime");
                }
            }
        }

        private string _StatusImage { get; set; }
        public string StatusImage
        {
            get { return _StatusImage; }
            set
            {
                if (_StatusImage != value)
                {
                    _StatusImage = value;
                    NotifyIt("StatusImage");
                }
            }
        }

        private string _ProcessBarMessage { get; set; }
        public string ProcessBarMessage
        {
            get { return _ProcessBarMessage; }
            set
            {
                if (_ProcessBarMessage != value)
                {
                    _ProcessBarMessage = value;
                    NotifyIt("ProcessBarMessage");
                }
            }
        }

        public DateTime ActionCompletedTime { get; set; }
        public DateTime ActionStartTime { get; set; }

        void UpdateProgressBarStatus()
        {
            while (true)
            {
                if (TaskStatus.Equals(ActionType.Completed))
                {
                    if (DateTime.Now.AddSeconds(-3) > ActionCompletedTime)
                        TaskStatus = ActionType.Stop;
                }
                else
                {
                    TaskStatus = ProgressVal.ActionType.Start;
                    ProgressBarValue = GlobalClass.progressVaue;
                }
                Thread.Sleep(10);
            }
        }

        private string _ProcessingTime { get; set; }
        public string ProcessingTime
        {
            get { return _ProcessingTime; }
            set
            {
                if (_ProcessingTime != value)
                {
                    _ProcessingTime = value;
                    NotifyIt("ProcessingTime");
                }
            }
        }

        Stopwatch objStopWatch = new Stopwatch();

        public enum ActionType
        {
            Start,
            Stop,
            Completed,
            None,
        }
    }
    #endregion
}
