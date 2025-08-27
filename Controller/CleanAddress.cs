using Newtonsoft.Json.Linq;
using ReverseGeoCoding.Common;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace ReverseGeoCoding.Controller
{
    internal class CleanAddress
    {
        #region Variable Declaration
        string inputfilefolder = "";
        string outputfilefolder = "";
        DataTable dtFinalOutput = null;
        int ik = 0;
        string timestamp = "";
        double result = 0;
        #endregion

        public CleanAddress(string InputFilepath, string OutputFilepath)
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
                        System.Windows.Forms.MessageBox.Show("Process was stopped !!! Please check at output path.", "Reverse Geocoding", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex) { }
        }

        public void CleanCustomerAddress()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
            dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Address,Output Lat/Long,Output Address,Error,Source,City,District,State,Pincode";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtFinalOutput = dtinputxlsx.Clone();

            // Ensure output columns exist
           
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

                    // Construct final address
                    //string finalAddress = $"{rawAddress}, {city}, {district}, {state} - {pincode}"
                    //                      .Replace(",,", ",").Trim().Trim(',');
                    string cleanedagainAddress = RemoveDuplicates(rawAddress, city, district, state);
                    // Add into output table
                    DataRow newRow = dtFinalOutput.NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    newRow["Output Address"] = cleanedagainAddress;
                    //newRow["City"] = city;
                    //newRow["District"] = district;
                    //newRow["State"] = state;
                    //newRow["Pincode"] = pincode;
                    //newRow["Source"] = "API";

                    dtFinalOutput.Rows.Add(newRow);
                }
                catch (Exception ex)
                {
                    DataRow newRow = dtFinalOutput.NewRow();
                    newRow.ItemArray = row.ItemArray.Clone() as object[];
                    newRow["Error"] = "Processing Error: " + ex.Message;
                    dtFinalOutput.Rows.Add(newRow);
                }
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
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Clean_Address_Template", timestamp);
                MessageBox.Show("Clean Address Template Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);
            }
        }

        /// <summary>
        /// Removes duplicate words from an address string
        /// </summary>
        /// 

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

        public void CleanCustomerAddress_()
        {
            var time = DateTime.Now.ToString();
            timestamp = HelperClass.UnixTimeStampUTC(time).ToString();
            dtFinalOutput = new DataTable();

            string dtcolumnValues = "UniqueId,Input Address,Output Lat/Long,Output Address,Error,Source";
            HelperClass.IsColumnExist(dtFinalOutput, dtcolumnValues, true);

            DataTable[] dtinput = HelperClass.GetDataTableConvertion(inputfilefolder);
            DataTable dtinputxlsx = dtinput[0];
            dtFinalOutput = dtinputxlsx.Clone();

            // Ensure output columns exist
            if (!dtFinalOutput.Columns.Contains("Output Lat/Long"))
                dtFinalOutput.Columns.Add("Output Lat/Long", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Output Address"))
                dtFinalOutput.Columns.Add("Output Address", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Error"))
                dtFinalOutput.Columns.Add("Error", typeof(string));
            if (!dtFinalOutput.Columns.Contains("Source"))
                dtFinalOutput.Columns.Add("Source", typeof(string));

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

            // 🔹 Setup Google Maps API Key
            string googleApiKey = "YOUR_GOOGLE_MAPS_API_KEY";

            // 🔹 Process Each Row
            foreach (DataRow row in dtinputxlsx.Rows)
            {
                string inputAddress = row["Input Address"].ToString();
                string feasibilityId = row["FEASIBILITY_ID"].ToString();

                try
                {
                    // ✅ Step 1: Clean Address
                    string cleanedAddress = CleanRawAddress(inputAddress);

                    // ✅ Step 2: Get Lat/Long
                    var (formatted, lat, lng) = GeocodeAddress(cleanedAddress, googleApiKey);

                    // ✅ Step 3: Fill Final Output
                    DataRow newRow = dtFinalOutput.NewRow();
                    newRow["FEASIBILITY_ID"] = feasibilityId;
                    newRow["Input Address"] = cleanedAddress;
                    //newRow["Output Address"] = string.IsNullOrEmpty(formatted) ? cleanedAddress : formatted;
                    //newRow["Output Lat/Long"] = (lat != null && lng != null) ? $"{lat}, {lng}" : "";
                    //newRow["Error"] = (lat == null || lng == null) ? "Geocoding failed" : "";
                    //newRow["Source"] = "Google Maps API";

                    dtFinalOutput.Rows.Add(newRow);
                }
                catch (Exception ex)
                {
                    DataRow errorRow = dtFinalOutput.NewRow();
                    errorRow["FEASIBILITY_ID"] = feasibilityId;
                    errorRow["Input Address"] = inputAddress;
                    errorRow["Output Address"] = "";
                    errorRow["Output Lat/Long"] = "";
                    errorRow["Error"] = ex.Message;
                    errorRow["Source"] = "System";

                    dtFinalOutput.Rows.Add(errorRow);
                }

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
                HelperClass.CreateExcelFilewithTime(outputfilefolder, dtFinalOutput, "Clean_Address_Template", timestamp);
                MessageBox.Show("Clean Address Template Output saved.", "Info", MessageBoxButton.OK, MessageBoxImage.Information);
                GlobalClass.ChangeForm.OnChangeForm(12);
            }
        }

        private static readonly HttpClient httpClient = new HttpClient();

        // ✅ Get City + District + State from Pincode
        public static (string city, string district, string state, string pincode) GetLocationFromPincode(string pincode)
        {
            try
            {
                string url = $"https://api.postalpincode.in/pincode/{pincode}";
                var response = httpClient.GetStringAsync(url).Result;
                JArray json = JArray.Parse(response);

                if (json[0]["Status"].ToString() == "Success")
                {
                    var postOffice = json[0]["PostOffice"][0];
                    string district = postOffice["District"].ToString();
                    string state = postOffice["State"].ToString();
                    string city = postOffice["Block"]?.ToString();

                    if (string.IsNullOrEmpty(city))
                        city = postOffice["Division"].ToString();

                    return (city, district, state, pincode);
                }
            }
            catch { }
            return (null, null, null, pincode);
        }

        // ✅ Clean Address (Now includes City + District + State + Pincode)
        public static string CleanRawAddress(string rawAddress)
        {
            string address = rawAddress ?? "";

            // Extract pincode (6 digits)
            var match = Regex.Match(address, @"\b\d{6}\b");
            if (match.Success)
            {
                string pincode = match.Value;
                var (city, district, state, pin) = GetLocationFromPincode(pincode);

                if (!string.IsNullOrEmpty(district))
                {
                    // Case: Only pincode given
                    if (address.Trim() == pincode)
                    {
                        return $"{city}, {district}, {state} - {pin}";
                    }
                    else
                    {
                        // Replace pincode with "City, District, State - Pincode"
                        address = address.Replace(pincode, $"{city}, {district}, {state} - {pin}");
                    }
                }
            }

            // Remove duplicate words
            var words = address.Split(new[] { ' ', ',', '.', '-' }, StringSplitOptions.RemoveEmptyEntries);
            var uniqueWords = words
                .Select(w => w.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase);

            return string.Join(" ", uniqueWords);
        }

        // ✅ Geocode + Ensure City/District/State/Pincode
        public static (string formattedAddress, string lat, string lng) GeocodeAddress(string rawAddress, string apiKey)
        {
            try
            {
                string url = $"https://maps.googleapis.com/maps/api/geocode/json?address={Uri.EscapeDataString(rawAddress)}&key={apiKey}";
                var response = httpClient.GetStringAsync(url).Result;
                JObject json = JObject.Parse(response);

                if (json["status"].ToString() == "OK")
                {
                    string formatted = json["results"][0]["formatted_address"].ToString();
                    var location = json["results"][0]["geometry"]["location"];
                    string lat = location["lat"].ToString();
                    string lng = location["lng"].ToString();

                    // Extract pincode from Google response
                    var pinComponent = json["results"][0]["address_components"]
                        .FirstOrDefault(c => ((JArray)c["types"]).Any(t => t.ToString() == "postal_code"));

                    string pincode = pinComponent?["long_name"]?.ToString();

                    // If missing, try to get from input
                    var match = Regex.Match(rawAddress, @"\b\d{6}\b");
                    if (string.IsNullOrEmpty(pincode) && match.Success)
                        pincode = match.Value;

                    // Append City/District/State/Pincode if available
                    if (!string.IsNullOrEmpty(pincode))
                    {
                        var (city, district, state, pin) = GetLocationFromPincode(pincode);
                        string fullLocation = $"{city}, {district}, {state} - {pin}";

                        if (!formatted.Contains(pin))
                            formatted = $"{formatted}, {fullLocation}";
                    }

                    return (formatted, lat, lng);
                }
            }
            catch { }
            return (null, null, null);
        }
    }
}
