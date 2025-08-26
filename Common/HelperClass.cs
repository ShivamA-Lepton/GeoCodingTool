using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Documents;
using System.Xml;
using System.Xml.Linq;

namespace ReverseGeoCoding.Common
{
    public static class HelperClass
    {
        //For saving excel in .xlsx format with timestamp
        public static void CreateExcelFilewithTime(string outputPath, DataTable dataTable, string filename, string timestampp)
        {
            // Combine path components to create the full file path with the desired filename format
            string filePath = Path.Combine(outputPath, $"{filename}_{timestampp}.xlsx");

            // Ensure the directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            // Create a spreadsheet document
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Create a spreadsheet document
                //using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(Path.Combine(outputPath, filename, timestampp), SpreadsheetDocumentType.Workbook))
                //{
                // Add a WorkbookPart to the document
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();

                // Create a Sheets collection
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Add a new sheet and associate it with the WorksheetPart
                Sheet sheet = new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);

                // Get the sheetData element of the WorksheetPart
                SheetData sheetData = worksheetPart.Worksheet.AppendChild(new SheetData());

                // Add the header row to the sheetData
                Row headerRow = new Row();
                foreach (DataColumn column in dataTable.Columns)
                {
                    Cell cell = new Cell();
                    cell.DataType = CellValues.InlineString;

                    InlineString inlineString = new InlineString();
                    Text text = new Text { Text = column.ColumnName };
                    inlineString.AppendChild(text);

                    cell.AppendChild(inlineString);
                    headerRow.AppendChild(cell);
                }
                sheetData.AppendChild(headerRow);

                // Add the DataTable content to the sheetData
                foreach (DataRow row in dataTable.Rows)
                {
                    Row excelRow = new Row();

                    foreach (var cellValue in row.ItemArray)
                    {
                        Cell cell = new Cell();
                        cell.DataType = CellValues.InlineString;

                        InlineString inlineString = new InlineString();
                        Text text = new Text { Text = cellValue.ToString() };
                        inlineString.AppendChild(text);

                        cell.AppendChild(inlineString);
                        excelRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(excelRow);
                }

                // Save changes to the spreadsheet document
                workbookPart.Workbook.Save();
            }
        }
        public static Int32 UnixTimeStampUTC(string time, bool isNextDayTime = false)
        {
            Int32 unixTimeStamp;
            try
            {
                DateTime currentTime = Convert.ToDateTime(time);
                if (isNextDayTime.Equals(true))
                {
                    currentTime = Convert.ToDateTime(time).AddDays(1);
                }
                DateTime zuluTime = currentTime.ToUniversalTime();
                DateTime unixEpoch = new DateTime(1970, 1, 1);
                unixTimeStamp = (Int32)(zuluTime.Subtract(unixEpoch)).TotalSeconds;
            }
            catch (Exception ex) { return unixTimeStamp = 0; }
            return unixTimeStamp;
        }
        public static DataTable[] GetDataTableConvertion(string _pathvalue)
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            try
            {
                var fileName = _pathvalue;
                var connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;IMEX=1;HDR=YES;TypeGuessRows=0;ImportMixedTypes=Text\"";
                using (var conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    DataSet ds;
                    var sheets = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                    int len = sheets.Rows.Count;
                    int kk = 0;
                    for (int i = 0; i <= len - 1; i++)
                    {
                        ds = new DataSet();
                        using (var cmd = conn.CreateCommand())
                        {
                            if ((sheets.Rows[i][2].Equals("Template$")) )
                            {
                                cmd.CommandText = "SELECT * FROM [" + sheets.Rows[i][2].ToString() + "] ";
                                var adapter = new OleDbDataAdapter(cmd);
                                adapter.Fill(ds);
                                if (kk == 0)
                                {
                                    dt = ds.Tables[0];
                                    kk++;
                                    continue;
                                }
                                if (kk == 1)
                                {
                                    dt1 = ds.Tables[0];
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { }
            return new DataTable[] { dt, dt1 };
        }
        public static bool IsColumnExist(DataTable dt, string columnName, bool create = false)
        {
            bool isExist = false;
            try
            {
                string dtColumnName = ",";
                foreach (DataColumn column in dt.Columns)
                {
                    dtColumnName = dtColumnName + column.Caption.ToLower().Trim() + ",";
                }
                string[] columnList = columnName.Split(',');
                foreach (string column in columnList)
                {
                    if (!dtColumnName.Contains("," + column.ToLower().Trim() + ","))
                    {
                        if (create.Equals(true))
                        {
                            dt.Columns.Add(column.Trim(), typeof(String));
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                isExist = true;
            }
            catch
            {

                isExist = false;
            }
            return isExist;
        }
        public static void ClearFolder(DirectoryInfo folder)
        {
            foreach (FileInfo file in folder.GetFiles())
            { file.Delete(); }
            foreach (DirectoryInfo subfolder in folder.GetDirectories())
            { ClearFolder(subfolder); }
        }
        private enum WindowShowStyle : uint
        {
            /// <summary>Hides the window and activates another window.</summary>
            /// <remarks>See SW_HIDE</remarks>
            Hide = 0,
            /// <summary>Activates and displays a window. If the window is minimized 
            /// or maximized, the system restores it to its original size and 
            /// position. An application should specify this flag when displaying 
            /// the window for the first time.</summary>
            /// <remarks>See SW_SHOWNORMAL</remarks>
            ShowNormal = 1,
            /// <summary>Activates the window and displays it as a minimized window.</summary>
            /// <remarks>See SW_SHOWMINIMIZED</remarks>
            ShowMinimized = 2,
            /// <summary>Activates the window and displays it as a maximized window.</summary>
            /// <remarks>See SW_SHOWMAXIMIZED</remarks>
            ShowMaximized = 3,
            /// <summary>Maximizes the specified window.</summary>
            /// <remarks>See SW_MAXIMIZE</remarks>
            Maximize = 3,
            /// <summary>Displays a window in its most recent size and position. 
            /// This value is similar to "ShowNormal", except the window is not 
            /// actived.</summary>
            /// <remarks>See SW_SHOWNOACTIVATE</remarks>
            ShowNormalNoActivate = 4,
            /// <summary>Activates the window and displays it in its current size 
            /// and position.</summary>
            /// <remarks>See SW_SHOW</remarks>
            Show = 5,
            /// <summary>Minimizes the specified window and activates the next 
            /// top-level window in the Z order.</summary>
            /// <remarks>See SW_MINIMIZE</remarks>
            Minimize = 6,
            /// <summary>Displays the window as a minimized window. This value is 
            /// similar to "ShowMinimized", except the window is not activated.</summary>
            /// <remarks>See SW_SHOWMINNOACTIVE</remarks>
            ShowMinNoActivate = 7,
            /// <summary>Displays the window in its current size and position. This 
            /// value is similar to "Show", except the window is not activated.</summary>
            /// <remarks>See SW_SHOWNA</remarks>
            ShowNoActivate = 8,
            /// <summary>Activates and displays the window. If the window is 
            /// minimized or maximized, the system restores it to its original size 
            /// and position. An application should specify this flag when restoring 
            /// a minimized window.</summary>
            /// <remarks>See SW_RESTORE</remarks>
            Restore = 9,
            /// <summary>Sets the show state based on the SW_ value specified in the 
            /// STARTUPINFO structure passed to the CreateProcess function by the 
            /// program that started the application.</summary>
            /// <remarks>See SW_SHOWDEFAULT</remarks>
            ShowDefault = 10,
            /// <summary>Windows 2000/XP: Minimizes a window, even if the thread 
            /// that owns the window is hung. This flag should only be used when 
            /// minimizing windows from a different thread.</summary>
            /// <remarks>See SW_FORCEMINIMIZE</remarks>
            ForceMinimized = 11
        }
        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, WindowShowStyle nCmdShow);
    }
}
