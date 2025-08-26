using Microsoft.Win32;
using ReverseGeoCoding.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ReverseGeoCoding.Controller
{
    internal class DownloadTemplate
    {
        #region Variable Declaration
        string inputfilefolder = "";
        string outputfilefolder = "";
        #endregion
        public DownloadTemplate(string InputFilepath, string OutputFilepath)
        {
            inputfilefolder = InputFilepath;
            outputfilefolder = OutputFilepath;
        }
        public void DownloadTemplateFile()
        {
            // Get the base directory of the running application (e.g., bin\Debug\netX.X)
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

            // Go up one level from bin/ and enter Template/ folder
            string sourcePath = Path.Combine(baseDirectory, @"..\..\Template\Template.xlsx");
            sourcePath = Path.GetFullPath(sourcePath); // Normalize to full path

            // Target path - Save directly to Downloads folder
            string downloadsPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads");
            string targetPath = Path.Combine(downloadsPath, "Template.xlsx");

            if (File.Exists(sourcePath))
            {
                try
                {
                    File.Copy(sourcePath, targetPath, overwrite: true);
                    MessageBox.Show($"Template downloaded to:\n{targetPath}", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Error copying file: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            else
            {
                MessageBox.Show("Template file not found!", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}
