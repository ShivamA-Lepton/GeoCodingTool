using DocumentFormat.OpenXml.Bibliography;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ReverseGeoCoding.Common;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ReverseGeoCoding.Controller
{
    internal class UploadTemplate
    {
        #region Variable Declaration
        string inputfilefolder = "";
        string outputfilefolder = "";
        DataTable dtFinalOutput = null; int ik = 0;
        string timestamp = ""; double result=0;
        #endregion
        public UploadTemplate(string InputFilepath, string OutputFilepath)
        {
            inputfilefolder = InputFilepath;
            outputfilefolder = OutputFilepath;
            GlobalClass.ChangeForm.ChangeEvent += new GlobalClass.FormSelectIndex(ClosePanel);
        }
        public void ClosePanel(int e)
        {
            try
            {
                if (e == 9)
                {
                    if ((outputfilefolder != string.Empty))
                    {
                        if (dtFinalOutput.Rows.Count > 0)
                        {
                            HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Reverse_Geocoding_With_Address", timestamp);
                        }
                        System.Windows.Forms.MessageBox.Show("Process was stoped !!! Please check at output path.", "Reverse Geocoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { }
        }
        public void ReverseGeoCoding()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
             dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Lat/Long,Input Address, Output Lat/Long,Output Address,Error";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            var _inputfolder = Path.GetDirectoryName(inputfilefolder);
            var _outputfolder = outputfilefolder;

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtFinalOutput = dtinputxlsx.Clone();

            if (!dtFinalOutput.Columns.Contains("Output Lat/Long"))
                dtFinalOutput.Columns.Add("Output Lat/Long", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Output Address"))
                dtFinalOutput.Columns.Add("Output Address", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Error"))
                dtFinalOutput.Columns.Add("Error", typeof(string));

            #region Required Column Validation
            string requiredColumns = "FEASIBILITY_ID,Input Lat/Long";
            string[] requiredCols = requiredColumns.Split(',');
            List<string> missingColumns = new List<string>();

            foreach (string col in requiredCols)
            {
                if (!dtinputxlsx.Columns.Contains(col.Trim()))
                {
                    missingColumns.Add(col.Trim());
                }
            }

            if (missingColumns.Count > 0)
            {
                MessageBox.Show(string.Join(", ", missingColumns) + " column(s) required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            #endregion

            dynamic root;
            string api_status = "";

            foreach (DataRow row in dtinputxlsx.Rows)
            {
                string errors = "";
                string address = "";
                string uniqueId = row["FEASIBILITY_ID"]?.ToString().Trim();
                string latLong = row["Input Lat/Long"]?.ToString().Trim();
                //string inputAddress = row["Address"]?.ToString().Trim();

                // Validation
                if (string.IsNullOrWhiteSpace(uniqueId))
                    errors += "FEASIBILITY_ID is empty. ";

                if (string.IsNullOrWhiteSpace(latLong))
                    errors += "Lat/Long is empty. ";

                //if (!string.IsNullOrWhiteSpace(inputAddress))
                //    errors += "Address is not empty. ";

                double lat = 0, lon = 0;
                bool latLongValid = false;

                if (!string.IsNullOrWhiteSpace(latLong))
                {
                    var parts = latLong.Split(',');
                    if (parts.Length == 2 &&
                        double.TryParse(parts[0], out lat) &&
                        double.TryParse(parts[1], out lon))
                    {
                        if (lat >= -90 && lat <= 90 && lon >= -180 && lon <= 180)
                        {
                            latLongValid = true;
                        }
                        else
                        {
                            errors += "Latitude or Longitude out of range. ";
                        }
                    }
                    else
                    {
                        errors += "Lat/Long format is invalid. ";
                    }
                }

                // Perform Reverse Geocoding if no error so far
                if (string.IsNullOrEmpty(errors) && latLongValid)
                {
                    try
                    {
                        string apiKey = ConfigurationManager.AppSettings["GoogleMapsApiKey"];
                        var url = $"https://maps.googleapis.com/maps/api/geocode/json?latlng={lat},{lon}&key={apiKey}";

                        var req1 = (HttpWebRequest)WebRequest.Create(url);
                        req1.Timeout = 30000;
                        req1.ReadWriteTimeout = 30000;

                        if (Convert.ToBoolean(ConfigurationManager.AppSettings["isProd"]))
                        {
                            WebProxy wp = new WebProxy("http://10.94.147.11:8080", true);
                            wp.Credentials = CredentialCache.DefaultCredentials;
                            wp.UseDefaultCredentials = true;
                            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                            req1.Proxy = wp;
                        }

                        var res1 = (HttpWebResponse)req1.GetResponse();

                        using (var streamreader = new StreamReader(res1.GetResponseStream()))
                        {
                            var result = streamreader.ReadToEnd();
                            if (!string.IsNullOrWhiteSpace(result))
                            {
                                root = JsonConvert.DeserializeObject(result);
                                if (root.status.ToString() == "OK")
                                {
                                    address = root.results[0].formatted_address.ToString();
                                    if (address.Contains("+"))
                                    {
                                        int plusIndex = address.IndexOf('+');
                                        address = address.Substring(plusIndex + 1).Trim();
                                    }
                                }
                                else
                                {
                                    errors += $"API returned status: {root.status}. ";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors += "Exception during API call. ";
                        // Optionally log: ex.Message
                    }
                }

                // Add record to final output
                DataRow newRow = dtFinalOutput.NewRow();
                foreach (DataColumn col in dtinputxlsx.Columns)
                {
                    newRow[col.ColumnName] = row[col.ColumnName];
                }
                newRow["Output Address"] = address;
                newRow["Error"] = errors.Trim();
                dtFinalOutput.Rows.Add(newRow);
            Progress:
                try
                {
                    ik++; double _value = Convert.ToDouble(100) / Convert.ToDouble(dtinputxlsx.Rows.Count); result = result == 0 ? _value : result + _value;
                    string remaining = Convert.ToString((dtinputxlsx.Rows.Count) - ik);
                    if (remaining == "0")
                    {
                        GlobalClass.progressVaue = 100;
                        GlobalClass.messagevalue = "Process Completed";
                    }
                    else
                    {
                        GlobalClass.progressVaue = result;
                        GlobalClass.messagevalue = ik + " records fetched. " + remaining + " records remaining.";
                    }
                }
                catch (Exception ex) { }
            }

            // Save the final output file (with error column)
            if (dtFinalOutput.Rows.Count > 0)
            {
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Reverse_Geocoding_With_Address", timestamp);
                MessageBox.Show("Reverse Geocoding completed. Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);

            }
        }
        public void ForwardGeoCoding_Merged()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
            dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Address,Output Lat/Long,Output Address,Error,Source";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtinputxlsx.Columns.Add("Output Address", typeof(string));
            dtFinalOutput = dtinputxlsx.Clone();

            // Ensure output columns exist
            if (!dtFinalOutput.Columns.Contains("Output Address"))
                dtFinalOutput.Columns.Add("Output Address", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Output Lat/Long"))
                dtFinalOutput.Columns.Add("Output Lat/Long", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Confidence_Radius"))
                dtFinalOutput.Columns.Add("Confidence_Radius", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Source"))
                dtFinalOutput.Columns.Add("Source", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Error"))
                dtFinalOutput.Columns.Add("Error", typeof(string));

            #region Required Column Validation
            string requiredColumns = "FEASIBILITY_ID,Input Address";
            string[] requiredCols = requiredColumns.Split(',');
            List<string> missingColumns = new List<string>();

            foreach (string col in requiredCols)
            {
                if (!dtinputxlsx.Columns.Contains(col.Trim()))
                {
                    missingColumns.Add(col.Trim());
                }
            }

            if (missingColumns.Count > 0)
            {
                MessageBox.Show(string.Join(", ", missingColumns) + " column(s) required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            #endregion


            // Process each row
            foreach (DataRow row in dtinputxlsx.Rows)
            {
                try
                {
                    string rawAddress = row["Input Address"].ToString();
                    //string cleanedAddress = RemoveDuplicates(rawAddress);

                    // Extract pincode
                    string pincode = System.Text.RegularExpressions.Regex.Match(rawAddress, @"\b\d{6}\b").Value;

                    string city = "", district = "", state = "";

                    // If pincode exists, lookup using API
                    if (!string.IsNullOrEmpty(pincode))
                    {
                        try
                        {
                            using (var client = new System.Net.WebClient())
                            {
                                string url = $"https://api.postalpincode.in/pincode/{pincode}";
                                string json = client.DownloadString(url);

                                var arr = Newtonsoft.Json.Linq.JArray.Parse(json);
                                var postOffice = arr[0]["PostOffice"]?.First;
                                if (postOffice != null)
                                {
                                    district = postOffice["District"]?.ToString() ?? "";
                                    state = postOffice["State"]?.ToString() ?? "";
                                    city = postOffice["Region"]?.ToString() ?? ""; // Region used as "City"
                                }
                            }
                        }
                        catch (Exception exApi)
                        {
                            row["Error"] = "API Error: " + exApi.Message;
                        }
                    }

                    
                    string cleanedagainAddress = RemoveDuplicates(rawAddress, city, district, state);
                    // Update row directly
                    row["Output Address"] = cleanedagainAddress;
                }
                catch (Exception ex)
                {
                    
                }
            }

            int ik = 0;
            double result = 0;

            // Read confidence threshold from app.config
            int confidenceThreshold = Convert.ToInt32(ConfigurationManager.AppSettings["ConfidenceThreshold"]);

            foreach (DataRow row in dtinputxlsx.Rows)
            {
                string errors = "";
                string latLong = "";
                string outputAddress = "";
                string source = "";
                string resp_result = "";
                int confidence = 0;
                string uniqueId = row["FEASIBILITY_ID"]?.ToString().Trim();
                string address = row["Output Address"]?.ToString().Trim();

                if (string.IsNullOrWhiteSpace(uniqueId))
                    errors += "FEASIBILITY_ID is empty. ";
                if (string.IsNullOrWhiteSpace(address))
                    errors += "Output Address is empty. ";

                if (string.IsNullOrEmpty(errors))
                {
                    resp_result = CallDLVApi(address);
                    var jsonObj = JObject.Parse(resp_result);

                     confidence = jsonObj["confidence_radius"]?.ToObject<int>() ?? 0;
                    // Select API based on confidence
                    if (confidence <= confidenceThreshold)
                    {
                       double lat = jsonObj["geocode"]?["lat"]?.ToObject<double>() ?? 0;
                        double lng = jsonObj["geocode"]?["lng"]?.ToObject<double>() ?? 0;
                        latLong = lat + " , " + lng;
                        source = "DLV API";
                    }
                    else
                    {
                        (latLong, outputAddress, errors, source) = CallGoogleApi(address);
                    }
                }

                // Add record to final output
                DataRow newRow = dtFinalOutput.NewRow();
                foreach (DataColumn col in dtinputxlsx.Columns)
                {
                    newRow[col.ColumnName] = row[col.ColumnName];
                }
                newRow["Confidence_Radius"] = confidence;
                newRow["Output Lat/Long"] = latLong;
                newRow["Output Address"] = address;
                newRow["Error"] = errors.Trim();
                newRow["Source"] = source;
                dtFinalOutput.Rows.Add(newRow); 

                #region Progress
                try
                {
                    ik++;
                    double _value = Convert.ToDouble(100) / Convert.ToDouble(dtinputxlsx.Rows.Count);
                    result = result == 0 ? _value : result + _value;
                    string remaining = Convert.ToString((dtinputxlsx.Rows.Count) - ik);

                    if (remaining == "0")
                    {
                        GlobalClass.progressVaue = 100;
                        GlobalClass.messagevalue = "Process Completed";
                    }
                    else
                    {
                        GlobalClass.progressVaue = result;
                        GlobalClass.messagevalue = ik + " records fetched. " + remaining + " records remaining.";
                    }
                }
                catch { }
                #endregion
            }

            // Save output file
            if (dtFinalOutput.Rows.Count > 0)
            {
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Geocoding_With_LatLong", timestamp);
                MessageBox.Show("Geocoding completed. Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);
            }
        }
        private static string RemoveDuplicates(string input, string city, string district, string state)
        {
            if (string.IsNullOrWhiteSpace(input)) return input;

            // Step 1: Insert space between letters & digits (e.g. "KHERI262723" -> "KHERI 262723")
            string normalized = System.Text.RegularExpressions.Regex.Replace(input, @"([A-Za-z])(\d)", "$1 $2");
            normalized = System.Text.RegularExpressions.Regex.Replace(normalized, @"(\d)([A-Za-z])", "$1 $2");

            // Step 2: Split into tokens
            var matches = System.Text.RegularExpressions.Regex.Matches(normalized, @"[A-Za-z\-]+|\d+");

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var result = new List<string>();
            string pincode = null;

            foreach (System.Text.RegularExpressions.Match m in matches)
            {
                string token = m.Value.Trim(new char[] { ',', '.', '-' });
                if (string.IsNullOrWhiteSpace(token)) continue;

                // Detect pincode
                if (int.TryParse(token, out _) && token.Length == 6)
                {
                    pincode = token;
                    continue; // will add later
                }

                // Normalize (remove suffix like "-I", "-II")
                string normToken = System.Text.RegularExpressions.Regex.Replace(token, @"[-_]\w+$", "");

                // Keep first occurrence only
                if (seen.Add(normToken.ToLower()))
                {
                    result.Add(token); // preserve original casing in output
                }
            }

            // Step 3: Build cleaned main address
            var finalAddress = string.Join(" ", result).Trim();

            // Step 4: Append city, district, state with a single comma before first one
            var locationParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(city) && !seen.Contains(city.ToLower())) locationParts.Add(city);
            if (!string.IsNullOrWhiteSpace(district) && !seen.Contains(district.ToLower())) locationParts.Add(district);
            if (!string.IsNullOrWhiteSpace(state) && !seen.Contains(state.ToLower())) locationParts.Add(state);

            if (locationParts.Count > 0)
            {
                finalAddress += ", " + string.Join(", ", locationParts);
            }

            // Step 5: Add pincode at the end with hyphen
            if (!string.IsNullOrWhiteSpace(pincode) && !finalAddress.Contains(pincode))
                finalAddress += " - " + pincode;

            return finalAddress.Trim();
        }
        /// <summary>
        /// Calls DLV API
        /// </summary>
        public string   CallDLVApi(string address)
        {
            string latLong = "";
            string outputAddress = "";
            string error = "";
            string source = "DLV";
            string resultres = "";
            try
            {
                string apiKey = ConfigurationManager.AppSettings["DLVApiKey"];
                string url = "https://api.getos1.com/locateone/v1/geocode";

                var payload = new { data = new { address = address } };
                string jsonPayload = JsonConvert.SerializeObject(payload);

                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                request.Method = "POST";
                request.ContentType = "application/json";
                request.Headers["X-api-key"] = apiKey;

                using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                {
                    streamWriter.Write(jsonPayload);
                }

                var response = (HttpWebResponse)request.GetResponse();
                using (var streamReader = new StreamReader(response.GetResponseStream()))
                {
                    string resultJson = streamReader.ReadToEnd();
                   
                    dynamic root = JsonConvert.DeserializeObject(resultJson);

                    if (root != null && root.success == true)
                    {
                        resultres = root.result.ToString();
                      // double lat = root.result.geocode.lat;
                      //  double lng = root.result.geocode.lng;
                        //latLong = $"{lat},{lng}";
                        //outputAddress = root.data.formatted_address?.ToString();
                    }
                    else
                    {
                        error += "DLV API returned error. ";
                    }
                }
            }
            catch (Exception ex)
            {
                error += $"DLV Exception: {ex.Message}. ";
            }

            return resultres;

        }
        /// <summary>
        /// Calls Google Geocoding API
        /// </summary>
        private (string latLong, string outputAddress, string error, string source) CallGoogleApi(string address)
        {
            string latLong = "";
            string outputAddress = "";
            string error = "";
            string source = "Geo-Coding API";

            try
            {
                string apiKey = ConfigurationManager.AppSettings["GoogleMapsApiKey"];
                string preferredType = ConfigurationManager.AppSettings["Rooftop"] ?? "ROOFTOP";
                var url = $"https://maps.googleapis.com/maps/api/geocode/json?address={Uri.EscapeDataString(address)}&key={apiKey}";

                var req1 = (HttpWebRequest)WebRequest.Create(url);
                req1.Timeout = 30000;
                req1.ReadWriteTimeout = 30000;

                if (Convert.ToBoolean(ConfigurationManager.AppSettings["isProd"]))
                {
                    WebProxy wp = new WebProxy("http://10.94.147.11:8080", true);
                    wp.Credentials = CredentialCache.DefaultCredentials;
                    wp.UseDefaultCredentials = true;
                    ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                    req1.Proxy = wp;
                }

                var res1 = (HttpWebResponse)req1.GetResponse();
                using (var streamreader = new StreamReader(res1.GetResponseStream()))
                {
                    var result = streamreader.ReadToEnd();
                    dynamic root = JsonConvert.DeserializeObject(result);

                    if (root.status.ToString() == "OK" && root.results != null && root.results.Count > 0)
                    {
                        dynamic selectedResult = null;
                        foreach (var item in root.results)
                        {
                            if (item.geometry.location_type.ToString().Equals(preferredType, StringComparison.OrdinalIgnoreCase))
                            {
                                selectedResult = item;
                                break;
                            }
                        }
                        if (selectedResult == null)
                            selectedResult = root.results[0];

                        double lat = selectedResult.geometry.location.lat;
                        double lng = selectedResult.geometry.location.lng;
                        latLong = $"{lat},{lng}";
                        outputAddress = selectedResult.formatted_address?.ToString();
                    }
                    else
                    {
                        error += $"Google API returned status: {root.status}. ";
                    }
                }
            }
            catch (Exception ex)
            {
                error += $"Google Exception: {ex.Message}. ";
            }

            return (latLong, outputAddress, error, source);
        }
        //DLV API
        public void ForwardGeoCoding()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
            dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Address,Output Lat/Long,Output Address,Error";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            var _inputfolder = Path.GetDirectoryName(inputfilefolder);
            var _outputfolder = outputfilefolder;

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtFinalOutput = dtinputxlsx.Clone();

            if (!dtFinalOutput.Columns.Contains("Output Lat/Long"))
                dtFinalOutput.Columns.Add("Output Lat/Long", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Output Address"))
                dtFinalOutput.Columns.Add("Output Address", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Error"))
                dtFinalOutput.Columns.Add("Error", typeof(string));

            #region Required Column Validation
            string requiredColumns = "FEASIBILITY_ID,Input Address";
            string[] requiredCols = requiredColumns.Split(',');
            List<string> missingColumns = new List<string>();

            foreach (string col in requiredCols)
            {
                if (!dtinputxlsx.Columns.Contains(col.Trim()))
                {
                    missingColumns.Add(col.Trim());
                }
            }

            if (missingColumns.Count > 0)
            {
                MessageBox.Show(string.Join(", ", missingColumns) + " column(s) required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            #endregion

            int ik = 0;
            double result = 0;
            dynamic root;

            foreach (DataRow row in dtinputxlsx.Rows)
            {
                string errors = "";
                string latLong = "";
                string outputAddress = "";
                string uniqueId = row["FEASIBILITY_ID"]?.ToString().Trim();
                string address = row["Input Address"]?.ToString().Trim();

                if (string.IsNullOrWhiteSpace(uniqueId))
                    errors += "FEASIBILITY_ID is empty. ";
                if (string.IsNullOrWhiteSpace(address))
                    errors += "Input Address is empty. ";

                if (string.IsNullOrEmpty(errors))
                {
                    try
                    {
                        string apiKey = "vXG4RKkd23c99fp4T5KNTknJ2K2p1wFKacnGTFxPh1nQeerO"; // Move to config
                        string url = "https://api.getos1.com/locateone/v1/geocode";

                        var payload = new
                        {
                            data = new
                            {
                                address = address
                            }
                        };

                        string jsonPayload = JsonConvert.SerializeObject(payload);

                        HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                        request.Method = "POST";
                        request.ContentType = "application/json";
                        request.Headers["X-api-key"] = apiKey;

                        using (var streamWriter = new StreamWriter(request.GetRequestStream()))
                        {
                            streamWriter.Write(jsonPayload);
                            streamWriter.Flush();
                            streamWriter.Close();
                        }

                        var response = (HttpWebResponse)request.GetResponse();
                        using (var streamReader = new StreamReader(response.GetResponseStream()))
                        {
                            string resultJson = streamReader.ReadToEnd();
                            if (!string.IsNullOrWhiteSpace(resultJson))
                            {
                                root = JsonConvert.DeserializeObject(resultJson);

                                if (root != null && root.success == true)
                                {
                                    try
                                    {
                                        double lat = root.result.geocode.lat;
                                        double lng = root.result.geocode.lng;
                                        latLong = $"{lat},{lng}";
                                        //outputAddress = root.data.formatted_address?.ToString();
                                    }
                                    catch
                                    {
                                        errors += "Invalid API response format. ";
                                    }
                                }
                                else
                                {
                                    errors += $"API returned error. ";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors += $"Exception during API call: {ex.Message}. ";
                    }
                }

                // Add record to final output
                DataRow newRow = dtFinalOutput.NewRow();
                foreach (DataColumn col in dtinputxlsx.Columns)
                {
                    newRow[col.ColumnName] = row[col.ColumnName];
                }
                newRow["Output Lat/Long"] = latLong;
                newRow["Output Address"] = outputAddress;
                newRow["Error"] = errors.Trim();
                dtFinalOutput.Rows.Add(newRow);

                #region Progress
                try
                {
                    ik++;
                    double _value = Convert.ToDouble(100) / Convert.ToDouble(dtinputxlsx.Rows.Count);
                    result = result == 0 ? _value : result + _value;
                    string remaining = Convert.ToString((dtinputxlsx.Rows.Count) - ik);

                    if (remaining == "0")
                    {
                        GlobalClass.progressVaue = 100;
                        GlobalClass.messagevalue = "Process Completed";
                    }
                    else
                    {
                        GlobalClass.progressVaue = result;
                        GlobalClass.messagevalue = ik + " records fetched. " + remaining + " records remaining.";
                    }
                }
                catch { }
                #endregion
            }

            // Save output file
            if (dtFinalOutput.Rows.Count > 0)
            {
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Geocoding_With_LatLong", timestamp);
                MessageBox.Show("Geocoding completed. Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);
            }
        }
        //Google API
        public void ForwardGeoCoding_()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
            dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Lat/Long,Input Address, Output Lat/Long,Output Address,Error";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            var _inputfolder = Path.GetDirectoryName(inputfilefolder);
            var _outputfolder = outputfilefolder;

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtFinalOutput = dtinputxlsx.Clone();

            if (!dtFinalOutput.Columns.Contains("Output Lat/Long"))
                dtFinalOutput.Columns.Add("Output Lat/Long", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Output Address"))
                dtFinalOutput.Columns.Add("Output Address", typeof(string));

            if (!dtFinalOutput.Columns.Contains("Error"))
                dtFinalOutput.Columns.Add("Error", typeof(string));

            #region Required Column Validation
            string requiredColumns = "FEASIBILITY_ID,Input Address";
            string[] requiredCols = requiredColumns.Split(',');
            List<string> missingColumns = new List<string>();

            foreach (string col in requiredCols)
            {
                if (!dtinputxlsx.Columns.Contains(col.Trim()))
                {
                    missingColumns.Add(col.Trim());
                }
            }

            if (missingColumns.Count > 0)
            {
                MessageBox.Show(string.Join(", ", missingColumns) + " column(s) required", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }
            #endregion

            dynamic root;
            string api_status = "";

            foreach (DataRow row in dtinputxlsx.Rows)
            {
                string errors = "";
                string latLong = "";
                string uniqueId = row["FEASIBILITY_ID"]?.ToString().Trim();
                string address = row["Input Address"]?.ToString().Trim();

                if (string.IsNullOrWhiteSpace(uniqueId))
                    errors += "FEASIBILITY_ID is empty. ";
                if (string.IsNullOrWhiteSpace(address))
                    errors += "Input Address is empty. ";

                if (string.IsNullOrEmpty(errors))
                {
                    try
                    {
                        string apiKey = ConfigurationManager.AppSettings["GoogleMapsApiKey"];
                        string preferredType = ConfigurationManager.AppSettings["Rooftop"] ?? "ROOFTOP";
                        var url = $"https://maps.googleapis.com/maps/api/geocode/json?address={Uri.EscapeDataString(address)}&key={apiKey}";

                        var req1 = (HttpWebRequest)WebRequest.Create(url);
                        req1.Timeout = 30000;
                        req1.ReadWriteTimeout = 30000;

                        if (Convert.ToBoolean(ConfigurationManager.AppSettings["isProd"]))
                        {
                            WebProxy wp = new WebProxy("http://10.94.147.11:8080", true);
                            wp.Credentials = CredentialCache.DefaultCredentials;
                            wp.UseDefaultCredentials = true;
                            ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;
                            req1.Proxy = wp;
                        }

                        var res1 = (HttpWebResponse)req1.GetResponse();

                        using (var streamreader = new StreamReader(res1.GetResponseStream()))
                        {
                            var result = streamreader.ReadToEnd();
                            if (!string.IsNullOrWhiteSpace(result))
                            {
                                root = JsonConvert.DeserializeObject(result);
                                if (root.status.ToString() == "OK" && root.results != null && root.results.Count > 0)
                                {
                                    dynamic selectedResult = null;

                                    foreach (var item in root.results)
                                    {
                                        if (item.geometry.location_type.ToString().Equals(preferredType, StringComparison.OrdinalIgnoreCase))
                                        {
                                            selectedResult = item;
                                            break;
                                        }
                                    }

                                    if (selectedResult == null)
                                        selectedResult = root.results[0];

                                    double lat = selectedResult.geometry.location.lat;
                                    double lng = selectedResult.geometry.location.lng;
                                    latLong = $"{lat},{lng}";
                                }
                                else
                                {
                                    errors += $"API returned status: {root.status}. ";
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        errors += "Exception during API call. ";
                    }
                }

                // Add record to final output
                DataRow newRow = dtFinalOutput.NewRow();
                foreach (DataColumn col in dtinputxlsx.Columns)
                {
                    newRow[col.ColumnName] = row[col.ColumnName];
                }
                newRow["Output Lat/Long"] = latLong;
                newRow["Error"] = errors.Trim();
                dtFinalOutput.Rows.Add(newRow);

            Progress:
                try
                {
                    ik++;
                    double _value = Convert.ToDouble(100) / Convert.ToDouble(dtinputxlsx.Rows.Count);
                    result = result == 0 ? _value : result + _value;
                    string remaining = Convert.ToString((dtinputxlsx.Rows.Count) - ik);

                    if (remaining == "0")
                    {
                        GlobalClass.progressVaue = 100;
                        GlobalClass.messagevalue = "Process Completed";
                    }
                    else
                    {
                        GlobalClass.progressVaue = result;
                        GlobalClass.messagevalue = ik + " records fetched. " + remaining + " records remaining.";
                    }
                }
                catch (Exception ex) { }
            }

            // Save output file
            if (dtFinalOutput.Rows.Count > 0)
            {
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Geocoding_With_LatLong", timestamp);
                MessageBox.Show("Geocoding completed. Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);
            }
        }
    }
}
