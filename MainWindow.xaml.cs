using ReverseGeoCoding.Common;
using ReverseGeoCoding.Controller;
using ReverseGeoCoding.View;
using System;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Configuration;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Threading;
using Microsoft.Win32;

namespace ReverseGeoCoding
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        string WindowAtTop = WebConfigurationManager.AppSettings["WindowAtTop"];
        Thread processed = null; DownloadTemplate downloadTemplate = null; UploadTemplate uploadTemplate = null;
        CleanAddress CleanAddress= null;
        public MainWindow()
        {
            File_processing Fileprocessing = null;
            Fileprocessing = new File_processing();
            this.DataContext = Fileprocessing;
            InitializeComponent();
            ServicePointManager.Expect100Continue = true;
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                   | SecurityProtocolType.Tls11
                   | SecurityProtocolType.Tls12
                   | SecurityProtocolType.Ssl3;
            this.PreviewKeyDown += new KeyEventHandler(HandleEsc);
            DataObject.AddPastingHandler(txtInputFilepath, OnPasteInput);
            DataObject.AddPastingHandler(txtoutputdrive, OnPasteOutput);
            GlobalClass.ChangeForm.ChangeEvent += new GlobalClass.FormSelectIndex(ClosePanel);
        }

        private void btnClose_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (GlobalClass.BangValue.Equals(true))
                    GlobalClass.ChangeForm.OnChangeForm(8);
                else
                    GlobalClass.ChangeForm.OnChangeForm(12);
                Environment.Exit(0);
            }
            catch { }
        }
        private void HandleEsc(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                Dispatcher.BeginInvoke(
                   DispatcherPriority.ContextIdle,
                   new Action(delegate ()
                   {
                       ProgressBar(false); GlobalClass.messagevalue = string.Empty;
                       GlobalClass.progressVaue = 0; Notify.CommandValue = "0";
                   }));
            }
        }

        #region for refreshing the value
        public void ClosePanel(int e)
        {
            try
            {
                if (e == 12)
                {
                    Dispatcher.BeginInvoke(
                    DispatcherPriority.ContextIdle,
                    new Action(delegate ()
                    {
                        txtoutputdrive.Text = String.Empty;
                        txtInputFilepath.Text = String.Empty;
                        ResetAllControl();
                        //Environment.Exit(0);
                    }));
                }
                if (e == 2)
                {
                    txtInputFilepath.Text = GlobalClass.InputFilepath;
                    txtoutputdrive.Text = System.IO.Path.GetDirectoryName(GlobalClass.InputFilepath) + "\\";
                }
                if (e == 6)
                {
                    Environment.Exit(0);
                }
                if(e==9)
                {
                    ProgressBar(false);
                }
                if (e == 8)
                {
                    ProgressBar(false);
                }
            }
            catch (Exception ex) { }

        }
        public void ResetAllControl()
        {
            try
            {
                ProgressBar(false); GlobalClass.progressVaue = 0; GlobalClass.messagevalue = string.Empty;

            }
            catch { }
        }
        #endregion

        #region for cut, copy & paste
        /// <summary>
        /// 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void OnPasteInput(object sender, DataObjectPastingEventArgs e)
        {
            var isText = e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, true);
            if (!isText) return;

            var text = e.SourceDataObject.GetData(DataFormats.UnicodeText) as string;
            GlobalClass.InputFilepath = text;
        }
        private void OnPasteOutput(object sender, DataObjectPastingEventArgs e)
        {
            var isText = e.SourceDataObject.GetDataPresent(DataFormats.UnicodeText, true);
            if (!isText) return;

            var text = e.SourceDataObject.GetData(DataFormats.UnicodeText) as string;
            GlobalClass.OutputFilepath = text;
        }
        #endregion

        void ProgressBar(bool show)
        {
            Storyboard sb = Resources[show ? "ShowProgressBar" : "HideProgressBar"] as Storyboard;
            sb.Begin(pnlTopMenu);
        }
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Environment.Exit(0);
        }
        private void btnWindowShowHide_Click(object sender, MouseButtonEventArgs e)
        {
            if (this.Width < 250 || this.WindowState == WindowState.Maximized)
            {
                this.Width = 630;
                this.Height = 480;
                this.Topmost = WindowAtTop.Equals("1") ? true : false;
                this.WindowStartupLocation = System.Windows.WindowStartupLocation.CenterScreen;

            }
            else
            {
                this.Width = 210;
                this.Height = 30;
                this.Topmost = true;
                this.Top = 0;

            }
        }
        private void btnMinimized_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Minimized)
            {
                WindowState = WindowState.Normal;
            }
            else //if (this.WindowState == WindowState.Normal)
            {
                WindowState = WindowState.Minimized;
            }
        }
        public bool Validation()
        {
            try
            {
                if (txtoutputdrive.Text.Trim() == string.Empty)
                {
                    MessageBox.Show("Please provide appropriate details !", "InterStates", MessageBoxButton.OK, MessageBoxImage.Warning);
                    txtInputFilepath.Focus();
                    return false;
                }
                return true;
            }
            catch { return false; }
        }
        private void btnInputFile_Click_1(object sender, RoutedEventArgs e)
        { }
        private void btntemplate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                downloadTemplate = new DownloadTemplate(GlobalClass.InputFilepath, GlobalClass.OutputFilepath);
                processed = new Thread(downloadTemplate.DownloadTemplateFile);
                processed.Start();

            }
            catch (Exception ex) { ProgressBar(false); System.Windows.Forms.MessageBox.Show(ex.Message.ToString() + "" + ex.Source.ToString(), "ReverseGeoCoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }
        }
        private void btnReverseGeoCoding_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Validation())
                {
                    ProgressBar(true);
                    uploadTemplate = new UploadTemplate(GlobalClass.InputFilepath, GlobalClass.OutputFilepath);
                    processed = new Thread(uploadTemplate.ReverseGeoCoding);
                    processed.Start();
                }
            }
            catch (Exception ex) { ProgressBar(false); System.Windows.Forms.MessageBox.Show(ex.Message.ToString() + "" + ex.Source.ToString(), "ReverseGeoCoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }
        }
        private void btnForwardGeoCoding_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Validation())
                {
                    ProgressBar(true);
                    uploadTemplate = new UploadTemplate(GlobalClass.InputFilepath, GlobalClass.OutputFilepath);
                    processed = new Thread(uploadTemplate.ForwardGeoCoding_Merged);
                    processed.Start();
                }
            }
            catch (Exception ex) { ProgressBar(false); System.Windows.Forms.MessageBox.Show(ex.Message.ToString() + "" + ex.Source.ToString(), "ReverseGeoCoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }
        }
        private void btnCleanAddress_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (Validation())
                {
                    ProgressBar(true);
                    CleanAddress = new CleanAddress(GlobalClass.InputFilepath, GlobalClass.OutputFilepath);
                    processed = new Thread(CleanAddress.CleanCustomerAddress);
                    processed.Start();
                }
            }
            catch (Exception ex) { ProgressBar(false); System.Windows.Forms.MessageBox.Show(ex.Message.ToString() + "" + ex.Source.ToString(), "ReverseGeoCoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information); }
        }
    }
    public static class Notify
    {
        public static string CommandValue { get; set; }
    }
    public static class NotifyValue
    {
        public static string Paramter { get; set; }
    }
}
