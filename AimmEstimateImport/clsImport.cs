using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aimm.Logging;
using System.IO;
using System.Xml;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Data;

namespace AimmEstimateImport
{
    public class clsImport
    {
        #region enums

        enum cellColors
        {
            errorColor = 3,
            warnColor = 55
        }

        #endregion

        #region objects

        clsExcel oXl = null;
        dynamic xlRange = null;
        dynamic xlCell = null;

        #endregion

        #region variables

        private string msg;
        private string connString;
        private string archivePath;
        private string errorPath;
        private string logPath;
        private bool showExcel = true;
        private string xlFile = "";
        private string xlSheet = "";
        private string destPath = "";
        private string destFile = "";
        private string destPathName = "";
        private string logPathName = "";
        private string[] projects;
        private bool isValid = false;

        #endregion

        public clsImport()
        {

        }

        ~clsImport()
        {

        }

        #region events

        /// <summary>
        /// for reporting status back to caller
        /// </summary>
        public event EventHandler<StatusChangedEventArgs> StatusChanged;
        protected virtual void OnStatusChanged(StatusChangedEventArgs e)
        {
            StatusChanged?.Invoke(this, e);
        }

        #endregion

        #region properties

        /// <summary>
        /// Full path to Excel file
        /// </summary>
        public string ExcelFile { get; set; }
        public string SourcePath { get; set; }

        public string Status
        {
            set { OnStatusChanged(new StatusChangedEventArgs(value)); }
        }

        #endregion



        #region methods

        public void InitClass(string settingsPath)
        {
            // get settings
            try
            {
                string settingsFile = Path.Combine(settingsPath, "Settings.xml");
                XmlDocument doc = new XmlDocument();
                doc.Load(settingsFile);
                connString = GetSetting(doc, "POLSQL");
                SourcePath = GetSetting(doc, "SourceFolder");
                archivePath = GetSetting(doc, "ArchiveFolder");
                errorPath = GetSetting(doc, "ErrorFolder");
                logPath = GetSetting(doc, "LogFolder");
                bool.TryParse(GetSetting(doc, "ShowExcel"), out showExcel);
                xlSheet = GetSetting(doc, "WorksheetName");
                LogIt.LogInfo("Got Settings");
            }
            catch(Exception ex)
            {
                msg = ex.Message;
                Status = msg;
                LogIt.LogError(msg);
            }
        }

        private string GetSetting(XmlDocument doc, string settingName)
        {
            string response = "";
            try
            {
                response = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");
            }
            catch(Exception)
            {
            }
            return response;
        }

        public void ImportExcel()
        {
            // continue if we can open excel file
            if(open_excel(ExcelFile, xlSheet))
            {
                xlFile = Path.GetFileName(ExcelFile);
                msg = $"Opened Excel file \"{xlFile}\"";
                Status = msg;
                LogIt.LogInfo(msg);

                // get total at default % and at 100%
                // used later to add extra module to end of AIMM job for extra profit
                float totalAtDefaultPercent = JobTotalAtCurrentPercent();
                SetJobPercent(100);
                float totalAt100Percent = JobTotalAtCurrentPercent();

                // continue if valid customer
                string jobID = "";
                int custID = 0;
                var id = oXl.GetSecondaryRange("ID").Value;
                int.TryParse((id ?? "").ToString(), out custID);
                if(custID != 0 && is_valid_aimm_cust(custID, connString))
                {

                    // continue if valid salesman, estimator, sale date, job description
                    int salesRep = 0;
                    var sr = oXl.GetSecondaryRange("SR").Value;
                    int.TryParse((sr ?? "").ToString().Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries)[1], out salesRep);
                    int estimator = 0;
                    var est = oXl.GetSecondaryRange("ESTIMATOR").Value;
                    int.TryParse((est ?? "").ToString().Split(new string[] { "(", ")" },StringSplitOptions.RemoveEmptyEntries)[1], out estimator);

                    bool isValidDate = false;
                    DateTime saleDate;
                    dynamic sd = oXl.GetSecondaryRange("sale_date");
                    try
                    {
                        // using value2 for date because it returns excel OA date
                        saleDate = DateTime.FromOADate(sd.Value2);
                        isValidDate = true;
                    }
                    catch(Exception)
                    {
                        isValidDate = DateTime.TryParse(sd.Value, out saleDate);
                    }

                    //var jd = oXl.GetSecondaryRange("job_description").Value;
                    //string jobDesc = (jd ?? "").ToString();
                    if(isValidDate && salesRep != 0 && estimator != 0)// && jobDesc != "")
                    {
                        // get job description from projects
                        string jobDesc = build_job_description();
                        
                        // iterate list of projects and options
                        List<string> projects = oXl.GetNamedRanges(new string[] { "specs_project_*", "specs_option_*" });
                        for(int i = 1; i <= projects.Count(); i++)
                        {
                            string projectRange = $"specs_project_{i}";
                            if(projects.Contains(projectRange))
                            {
                                // get the project range, continue if approved
                                oXl.GetRange(projectRange);
                                if(ProjectIsApproved())
                                {
                                    dynamic thisProject = oXl.Range;
                                    dynamic largestModule = null;
                                    bool aModuleHasNoManDays = false;
                                    float saveMtlCost = 0;
                                    foreach(dynamic module in thisProject.Rows)
                                    {
                                        // process this module if it has a total price
                                        if(RowHasTotal(module))
                                        {
                                            string modDesc = module.Cells(2).Value;
                                            float manDays = GetManDays(module);
                                            float mtlCost = GetMaterialCost(module);

                                            // if no man days, save material cost for later
                                            if(manDays == 0 && mtlCost != 0)
                                            {
                                                largestModule = GetLargestModule(thisProject);
                                                aModuleHasNoManDays = (largestModule != null);
                                                if(aModuleHasNoManDays)
                                                    saveMtlCost = mtlCost;
                                            }
                                            else
                                            {
                                                // if we're saving a cost and this is largest row, add it to this row's cost
                                                if(aModuleHasNoManDays && saveMtlCost != 0 && module.Address == largestModule.Address)
                                                {
                                                    mtlCost += saveMtlCost;
                                                    saveMtlCost = 0;
                                                    aModuleHasNoManDays = false;
                                                    largestModule = null;
                                                }

                                                // add job to aimm if we haven't already
                                                if(jobID == "")
                                                    jobID = add_aimm_job(custID, jobDesc, salesRep, estimator, saleDate, connString);

                                                // add the module
                                                //add_aimm_module();





                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        msg = $"Invalid Sale Date (\"{sd.Value ?? ""}\"), Sales Rep (\"{sr}\") or Estimator (\"{est}\"), estimate not imported";
                        LogIt.LogError(msg);
                        Status = msg;
                    }
                }
                else
                {
                    msg = $"Could not validate customer ID {custID}, estimate not imported";
                    LogIt.LogError(msg);
                    Status = msg;
                }



                isValid = oXl.CloseWorkbook();
                oXl.CloseExcel();
                oXl = null;

            }
            else
            {
                msg = $"Could not open Excel file \"{ExcelFile}\", estimate not imported";
                LogIt.LogError(msg);
                Status = msg;
            }

        }

        private bool add_aimm_module(string jobID, int custID, string modDesc, int estTypeID, int salesman, int estimator)
        {
            bool result = false;

            return result;
        }

        private string add_aimm_job(int custID, string jobDesc, int salesman, int estimator, DateTime saleDate, string connectionString)
        {
            string result = "";
            int jobStatus = 12; // div 4 pending
            int jobCoordinator = 251; // no jc assigned
            int jobType = 1; // construction
            int coID = 1;

            DateTime createDate = DateTime.Now;
            DateTime weDate = get_week_start_and_end(createDate, DayOfWeek.Wednesday).Value;

            float laborRate = get_labor_rate(createDate, connectionString);
            float burdenRate = get_burden_rate(createDate, connectionString);
            float commRate = get_commission_rate(createDate, connectionString);

            // get a random job id not used in the current year
            List<string> jobsForYear = get_jobs_for_year(connectionString);
            string jobID = get_unique_job_id(jobsForYear);

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    msg = $"Adding new AIMM job for customer ID {custID}";
                    LogIt.LogInfo(msg);
                    Status = msg;

                    string cmdText = "AddAIMMJob";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        cmd.Parameters.AddWithValue("@custID", custID);
                        cmd.Parameters.AddWithValue("@salesmanID", salesman);
                        cmd.Parameters.AddWithValue("@jobDesc", jobDesc);
                        cmd.Parameters.AddWithValue("@laborRate", laborRate);
                        cmd.Parameters.AddWithValue("@burdenRate", burdenRate);
                        cmd.Parameters.AddWithValue("@saleDate", saleDate);
                        cmd.Parameters.AddWithValue("@estimatorID", estimator);
                        cmd.Parameters.AddWithValue("@weDate", weDate);
                        cmd.Parameters.AddWithValue("@defCommRate", commRate);
                        cmd.Parameters.AddWithValue("@jobType", jobType);
                        cmd.Parameters.AddWithValue("@coID", coID);

                        conn.Open();
                        int rows = (int)cmd.ExecuteNonQuery();
                        if(rows == 1)
                            result = jobID;
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error validating customer ID {custID}: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Verify valid customer
        /// </summary>
        /// <param name="custID"></param>
        /// <param name="connectionString"></param>
        /// <returns>boolean indicating customer exists in database</returns>
        private bool is_valid_aimm_cust(int custID, string connectionString)
        {
            bool isValid = false;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Validating customer ID {custID}");
                    string cmdText = "SELECT COUNT(*) FROM MLG.dbo.[Customers And Prospects] WHERE [Customer ID] = @custID";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@custID", custID);
                        conn.Open();
                        int rows = (int)cmd.ExecuteScalar();
                        isValid = (rows > 0);
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error validating customer ID {custID}: {ex.Message}");
            }

            return isValid;
        }


        private float get_burden_rate(DateTime rateDate, string connectionString)
        {
            float result = 0;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting labor burden rate for {rateDate.ToShortDateString()}");
                    string cmdText = "SELECT BurdenRate FROM MLG.POL.tblBurdenRate WHERE @date BETWEEN StartDate and EndDate";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@date", rateDate);
                        conn.Open();
                        result = (float)cmd.ExecuteScalar();
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting labor burden rate for {rateDate.ToShortDateString()}: {ex.Message}");
            }
            return result;
        }

        private float get_commission_rate(DateTime rateDate, string connectionString)
        {
            float result = 0;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting commission rate for {rateDate.ToShortDateString()}");
                    string cmdText = "SELECT ComissRate FROM MLG.POL.tblComissRate WHERE @date BETWEEN StartDate and EndDate";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@date", rateDate);
                        conn.Open();
                        result = (float)cmd.ExecuteScalar();
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting commission rate for {rateDate.ToShortDateString()}: {ex.Message}");
            }
            return result;
        }

        private float get_labor_rate(DateTime rateDate, string connectionString)
        {
            float result = 0;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting labor rate for {rateDate.ToShortDateString()}");
                    string cmdText = "SELECT LaborRate FROM MLG.POL.tblLaborRate WHERE @date BETWEEN StartDate and EndDate";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@date", rateDate);
                        conn.Open();
                        result = (float)cmd.ExecuteScalar();
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting labor rate for {rateDate.ToShortDateString()}: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// get all timesheet IDs for the month supplied
        /// </summary>
        /// <param name="timesheetDate"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private List<string> get_jobs_for_year(string connectionString)
        {
            string yy = DateTime.Now.ToString("yy");
            string msg = $"Getting job IDs for current year";
            List<string> jobs = new List<string>();

            try
            {
                // get data using DataTable
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = $"SELECT JobID FROM MLG.POL.tblJobs WHERE JobID like '{yy}%'";
                    using(SqlDataAdapter dap = new SqlDataAdapter(cmdText, conn))
                    {
                        using(DataTable dt = new DataTable())
                        {
                            dap.Fill(dt);
                            foreach(DataRow row in dt.Rows)
                            {
                                jobs.Add(row[0].ToString());
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting timesheet IDs for current year: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return jobs;

        }

        /// <summary>
        /// build random timesheet id that hasn't been used in the current month
        /// </summary>
        /// <param name="timesheetDate"></param>
        /// <param name="existingJobs"></param>
        /// <returns></returns>
        private string get_unique_job_id(List<string> existingJobs)
        {
            string result = "";
            string yy = DateTime.Now.ToString("yy");
            const string pool = "0123456789";
            do
            {
                result = generate_random_string(6, yy, pool);
            } while(existingJobs.Contains(result));

            return result;
        }

        /// <summary>
        /// get week start and end dates for supplied date
        /// </summary>
        /// <param name="workDate">date to use to get start and end dates for week</param>
        /// <param name="wkEndDay"><see cref="DayOfWeek"/> enum member identifying the week-ending day</param>
        /// <returns></returns>
        private KeyValuePair<DateTime, DateTime> get_week_start_and_end(DateTime workDate, DayOfWeek wkEndDay)
        {
            int dow = (int)workDate.DayOfWeek;
            int wed = (int)wkEndDay;
            int daysToEow = (wed - dow) >= 0 ? wed - dow : wed - dow + 7;
            return new KeyValuePair<DateTime, DateTime>(workDate.AddDays(daysToEow - 6), workDate.AddDays(daysToEow));
        }

        /// <summary>
        /// Returns job description from 3 largest project types
        /// </summary>
        /// <returns></returns>
        private string build_job_description()
        {
            string result = "";
            List<KeyValuePair<string, float>> modList = new List<KeyValuePair<string, float>>();
            List<string> projects = oXl.GetNamedRanges("specs_project_*");
            foreach(string project in projects)
            {
                if(oXl.GetRange(project))
                {
                    float jobTTL = 0;
                    string ttl = (oXl.RangeOffset(0, -2, 1, 1).Value ?? "").ToString();
                    float.TryParse(ttl, out jobTTL);
                    string approved = (oXl.RangeOffset(1, -2, 1, 1).Value ?? "").ToString();
                    if(jobTTL > 0 && approved == "APPROVED")
                    {
                        string projDesc = oXl.RangeOffset(0, 0, 1, 1).Value;
                        modList.Add(new KeyValuePair<string, float>(projDesc, jobTTL));
                    }
                }

            }
            var sorted = from kvp in modList
                         orderby kvp.Value descending
                         select kvp.Key;
            result = string.Join(", ", sorted.Take(3).ToList());
            return result;
        }

        /// <summary>
        /// start ms excel and open supplied workbook name
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns>boolean indicating success status</returns>
        private bool open_excel(string fileName, string sheetName = "")
        {
            bool result = false;
            string msg = "";

            if(File.Exists(fileName))
            {
                var xlFile = Path.GetFileName(fileName);
                try
                {
                    oXl = new clsExcel();
                    oXl.Visible = showExcel;
                    result = oXl.OpenExcel(ExcelFile, sheetName);
                }
                catch(Exception ex)
                {
                    msg = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    Status = msg;
                    LogIt.LogError(msg);
                }
            }
            return result;
        }

        /// <summary>
        /// Returns whether selected project is approved (assumes a project is selected on SPECS sheet)
        /// </summary>
        /// <returns>Boolean</returns>
        private bool ProjectIsApproved()
        {
            bool result = false;
            dynamic cell = oXl.RangeOffset(1, -2, 1, 1);
            if(cell != null)
            {
                result = (cell.Value == "APPROVED");
            }
            cell = null;
            return result;
        }

        /// <summary>
        /// Gets job total at currently set percent level
        /// </summary>
        /// <returns></returns>
        private float JobTotalAtCurrentPercent()
        {
            float result = 0;
            dynamic cell = oXl.GetSecondaryRange("specs_job_total");
            if(cell != null)
            {
                var temp = cell.Value;
                float.TryParse(temp.ToString(), out result);
            }
            cell = null;
            return result;
        }

        /// <summary>
        /// Returns whether specs sheet row total price is non-zero
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private bool RowHasTotal(dynamic row)
        {
            bool result = false;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange("specs_proj_total_2"));
                result = (cell.Value ?? 0) != 0;
                cell = null;
            }
            catch(Exception ex)
            {
                msg = $"Error getting row total: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
                result = false;
            }
            return result;
        }

        /// <summary>
        /// Returns specs sheet row man days
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float GetManDays(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange("specs_proj_man_days"));
                float.TryParse((cell.Value ?? "").ToString(), out result);
                cell = null;
            }
            catch(Exception ex)
            {
                msg = $"Error getting row man days: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// Returns specs sheet row man days
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float GetMaterialCost(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange("specs_proj_material_cost"));
                float.TryParse((cell.Value ?? "").ToString(), out result);
                cell = null;
            }
            catch(Exception ex)
            {
                msg = $"Error getting row material cost: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// Returns the project row containing the largest total
        /// </summary>
        /// <param name="projectRange"></param>
        /// <returns></returns>
        private dynamic GetLargestModule(dynamic projectRange)
        {
            dynamic result = null;
            try
            {
                // get the totals column for the project
                dynamic totals = oXl.GetIntersect(projectRange.EntireRow, oXl.GetSecondaryRange("specs_proj_total_2"));
                if(totals != null)
                {
                    // get the cell containing the largest amount
                    dynamic bigCell = oXl.GetMaxCellInRange(totals);
                    if(bigCell != null)
                    {
                        // get the project row for the cell
                        result = oXl.GetIntersect(bigCell.EntireRow, projectRange);
                        bigCell = null;
                    }
                    totals = null;
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting largest row: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// Sets the current job percent to the supplied percentage
        /// </summary>
        /// <param name="newPercent">Desired percent expressed as percentage (ex: 100, 98.5, etc.)</param>
        /// <remarks><see cref="newPercent"/> supplied will be divided by 100 before setting cell value</remarks>
        private void SetJobPercent(float newPercent)
        {

            if(newPercent != 0)
            {
                dynamic cell = oXl.GetSecondaryRange("specs_percent_hours");
                if(cell != null)
                {
                    try
                    {
                        cell.Value = newPercent / 100;
                    }
                    catch(Exception ex)
                    {
                        msg = $"Error setting job percent: {ex.Message}";
                        Status = msg;
                        LogIt.LogError(msg);
                    }
                }
                cell = null;
            }
        }

        /// <summary>
        /// generates a random string of characters
        /// </summary>
        /// <param name="length">length of string to generate</param>
        /// <param name="prefix">text to start the random string with</param>
        /// <param name="pool">list of available characters to choose from</param>
        /// <returns></returns>
        private string generate_random_string(int length, string prefix, string pool)
        {
            Random rand = new Random();
            var sb = new StringBuilder(prefix);

            for(var i = sb.Length; i < length; i++)
            {
                var c = pool[rand.Next(0, pool.Length - 1)];
                sb.Append(c);
            }

            return sb.ToString();
        }


        #endregion


    }

    /// <summary>
    /// for reporting status back to caller
    /// </summary>
    public class StatusChangedEventArgs : EventArgs
    {
        private string e;
        public StatusChangedEventArgs(string e)
        {
            Status = e;
        }
        public string Status { get; set; }
    }


}
