using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Aimm.Logging;
using System.IO;
using System.Xml;
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

        enum rateTypes
        {
            burdenRate,
            commissionRate,
            laborRate
        }

        enum moduleTypes
        {
            all,
            demo,
            poolDemo
        }

        #endregion

        #region objects

        clsExcel oXl = null;
        dynamic xlRange = null;
        dynamic xlCell = null;
        Dictionary<string, string> rangeNames = new Dictionary<string, string>
        {
            { "cust_id", "ID" },
            { "sales_rep", "SR" },
            { "estimator", "ESTIMATOR" },
            { "sale_date", "sale_date" },
            { "job_total", "specs_job_total" },
            { "percent_hours" ,"specs_percent_hours" },
            { "man_days_range", "specs_proj_man_days" },
            { "material_cost_range", "specs_proj_material_cost" },
            { "project_total_range", "specs_proj_total_2" },
            { "project_approval_range", "specs_proj_approvals" },
            { "aimm_category_range", "specs_proj_aimm_category" },
            { "sub_cost_range", "specs_proj_sub_cost" }
        };


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
        private int modulesInJob = 0;
        private int subModulesInJob = 0;
        private string jobID = "";
        private string demo_types = "";
        private string pool_demo_types = "";
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

        #region public methods

        public void ImportECM()
        {
            // continue if we can open excel file
            if(open_excel(ExcelFile, xlSheet))
            {
                string adminModuleID = "";

                xlFile = Path.GetFileName(ExcelFile);
                msg = $"Opened Excel file \"{xlFile}\"";
                Status = msg;
                LogIt.LogInfo(msg);

                // get job total, continue if valid
                float totalAtDefaultPercent = get_job_total();
                if(totalAtDefaultPercent > 0)
                {

                    // get total at 100%
                    // later we will add an extra module to AIMM job for difference
                    if(set_job_percent(100))
                    {
                        float totalAt100Percent = get_job_total();

                        // continue if valid customer
                        int custID = 0;
                        var id = oXl.GetSecondaryRange(rangeNames["cust_id"]).Value;
                        int.TryParse((id ?? "").ToString(), out custID);
                        if(custID != 0 && is_valid_aimm_cust(custID, connString))
                        {

                            // continue if valid salesman, estimator, sale date
                            int salesRep = 0;
                            int estimator = 0;
                            var sr = oXl.GetSecondaryRange(rangeNames["sales_rep"]).Value ?? "";
                            var est = oXl.GetSecondaryRange(rangeNames["estimator"]).Value ?? "";
                            int.TryParse(get_last_paren_segment(sr), out salesRep);
                            int.TryParse(get_last_paren_segment(est), out estimator);

                            bool isValidDate = false;
                            DateTime saleDate;
                            dynamic sd = oXl.GetSecondaryRange(rangeNames["sale_date"]);
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

                            if(isValidDate && salesRep != 0 && estimator != 0)
                            {
                                // get job description from projects
                                string jobDesc = build_job_description();

                                // get rates
                                float laborRate = get_rate(rateTypes.laborRate, DateTime.Now, connString);
                                float burdenRate = get_rate(rateTypes.burdenRate, DateTime.Now, connString);
                                float commRate = get_rate(rateTypes.commissionRate, DateTime.Now, connString);

                                // add job, import projects and options
                                bool isOK = import_projects(custID, jobDesc, salesRep, estimator, saleDate, laborRate, burdenRate, commRate);

                                // add demo work order
                                int woNumber = 0;
                                using(DataTable demoModules = get_job_modules(moduleTypes.demo, connString))
                                {
                                    if(demoModules.Rows.Count > 0)
                                    {
                                        woNumber++;
                                        isOK = add_aimm_work_order(custID, woNumber, demoModules, laborRate, burdenRate, commRate, connString);
                                    }
                                }

                                // add pool demo work order
                                using(DataTable poolDemoModules = get_job_modules(moduleTypes.poolDemo, connString))
                                {
                                    if(poolDemoModules.Rows.Count > 0)
                                    {
                                        woNumber++;
                                        isOK = add_aimm_work_order(custID, woNumber, poolDemoModules, laborRate, burdenRate, commRate, connString);
                                    }
                                }

                                // add admin module for difference between totals
                                // use .01 man-days for time
                                modulesInJob++;
                                float modulePrice = totalAtDefaultPercent - totalAt100Percent;
                                adminModuleID = add_aimm_module(custID, modulesInJob, "ADMIN MODULE", (float).01, 0, modulePrice, 4, salesRep, estimator, laborRate, burdenRate, commRate, connString);

                                // update job totals from modules
                                isOK = update_aimm_job_totals(connString);
                                if(isOK)
                                {
                                    msg = $"AIMM job {jobID} created";
                                    LogIt.LogInfo(msg);
                                    Status = msg;
                                }
                                else
                                {
                                    msg = $"AIMM job {jobID} created, but totals were NOT updated. Please load job in AIMM, edit and save to refresh totals.";
                                    LogIt.LogError(msg);
                                    Status = msg;
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
                    }
                    else
                    {
                        msg = "Could not set estimate to 100%, estimate not imported";
                        LogIt.LogError(msg);
                        Status = msg;
                    }

                }
                else
                {
                    msg = $"Invalid job total ({totalAtDefaultPercent.ToString("$#0.00")}), estimate not imported";
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

        public void InitClass(string settingsPath)
        {
            // get settings
            try
            {
                string settingsFile = Path.Combine(settingsPath, "Settings.xml");
                XmlDocument doc = new XmlDocument();
                doc.Load(settingsFile);
                SourcePath = get_setting(doc, "SourceFolder");
                archivePath = get_setting(doc, "ArchiveFolder");
                errorPath = get_setting(doc, "ErrorFolder");
                logPath = get_setting(doc, "LogFolder");
                bool.TryParse(get_setting(doc, "ShowExcel"), out showExcel);
                xlSheet = get_setting(doc, "WorksheetName");
                demo_types = get_setting(doc, "DemoWorkTypes");
                pool_demo_types = get_setting(doc, "PoolDemoWorkTypes");

                string cs = get_setting(doc, "POLSQL");
                string usr = get_string_segment(cs, "User ID=", ";");
                string pwd = get_string_segment(cs, "Password=", ";");
                connString = cs.Replace(usr, clsCrypto.Decrypt(usr)).Replace(pwd, clsCrypto.Decrypt(pwd));

                LogIt.LogInfo("Got Settings");
            }
            catch(Exception ex)
            {
                msg = ex.Message;
                Status = msg;
                LogIt.LogError(msg);
            }
        }

        #endregion

        #region private methods

        /// <summary>
        /// Adds job to AIMM tables, returns job ID or "" if unsuccessful.
        /// </summary>
        /// <param name="custID"></param>
        /// <param name="jobDesc"></param>
        /// <param name="salesman"></param>
        /// <param name="estimator"></param>
        /// <param name="saleDate"></param>
        /// <param name="laborRate"></param>
        /// <param name="burdenRate"></param>
        /// <param name="commRate"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private string add_aimm_job(int custID, string jobDesc, int salesman, int estimator, DateTime saleDate,
                                    float laborRate, float burdenRate, float commRate, string connectionString)
        {
            string result = "";
            int jobType = 1; // construction
            int coID = 1;

            DateTime createDate = DateTime.Now;
            DateTime weDate = get_week_start_and_end(createDate, DayOfWeek.Wednesday).Value;

            // get a random job id not used in the current year
            List<string> jobsForYear = get_jobs_for_year(connectionString);
            string newJobID = get_unique_job_id(jobsForYear);

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    msg = $"Adding new AIMM job for customer ID {custID}";
                    LogIt.LogInfo(msg);
                    Status = msg;

                    string cmdText = "POL.AddAIMMJob";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@jobID", newJobID);
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
                            result = newJobID;
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error adding AIMM job for customer ID {custID}: {ex.Message}");
            }

            return result;
        }

        private string add_aimm_module(int custID, int moduleNumber, string modDesc, float manDays, float mtlCost, float modulePrice,
                                       int estTypeID, int salesman, int estimator, float laborRate, float burdenRate, float commRate, string connectionString)
        {
            string result = "";
            string estCommonID = $"{jobID}-{moduleNumber.ToString("00")}";
            float laborHours = manDays * 8;
            float laborCost = laborHours * laborRate;
            float burden = laborHours * burdenRate;
            float gp = (modulePrice - mtlCost - laborCost - burden) / (1 + commRate);
            float commission = gp * commRate;

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    msg = $"Adding AIMM module {estCommonID} ({modDesc})";
                    LogIt.LogInfo(msg);
                    Status = msg;

                    string cmdText = "POL.AddAimmModule";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        cmd.Parameters.AddWithValue("@projectEstCommonID", estCommonID);
                        cmd.Parameters.AddWithValue("@custID", custID);
                        cmd.Parameters.AddWithValue("@moduleDesc", modDesc);
                        cmd.Parameters.AddWithValue("@estTypeID", estTypeID);
                        cmd.Parameters.AddWithValue("@salesmanID", salesman);
                        cmd.Parameters.AddWithValue("@estimatorID", estimator);
                        cmd.Parameters.AddWithValue("@manDays", manDays);
                        cmd.Parameters.AddWithValue("@materialsCost", mtlCost);
                        cmd.Parameters.AddWithValue("@gp", gp);
                        cmd.Parameters.AddWithValue("@commisRate", commRate);
                        cmd.Parameters.AddWithValue("@commis", commission);
                        cmd.Parameters.AddWithValue("@price", modulePrice);
                        cmd.Parameters.AddWithValue("@laborRate", laborRate);
                        cmd.Parameters.AddWithValue("@burdenRate", burdenRate);
                        cmd.Parameters.AddWithValue("@laborCost", laborCost);
                        cmd.Parameters.AddWithValue("@burden", burden);

                        conn.Open();
                        int rows = (int)cmd.ExecuteNonQuery();
                        if(rows == 1)
                            result = estCommonID;
                    }
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error adding AIMM module {estCommonID}: {ex.Message}");
            }
            return result;
        }


        private string add_aimm_sub_module(int custID, int moduleNumber, string modDesc, float mtlCost, float subCost,
                                           float modulePrice, int estTypeID, int salesman, float commRate, string connectionString)
        {
            string result = "";
            string subModuleID = $"{jobID}-{moduleNumber.ToString("00")}S";
            float gp = (modulePrice - subCost - mtlCost) / (1 + commRate);
            float commission = gp * commRate;
            int modType = 1; // normal module
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    msg = $"Adding AIMM subcontractor module {subModuleID} ({modDesc}).";
                    LogIt.LogInfo(msg);
                    Status = msg;

                    string cmdText = "POL.AddAimmSubModule";
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        cmd.Parameters.AddWithValue("@subModID", subModuleID);
                        cmd.Parameters.AddWithValue("@custID", custID);
                        cmd.Parameters.AddWithValue("@moduleDesc", modDesc);
                        cmd.Parameters.AddWithValue("@salesmanID", salesman);
                        cmd.Parameters.AddWithValue("@subCost", subCost);
                        cmd.Parameters.AddWithValue("@materialsCost", mtlCost);
                        cmd.Parameters.AddWithValue("@gp", gp);
                        cmd.Parameters.AddWithValue("@commisRate", commRate);
                        cmd.Parameters.AddWithValue("@commis", commission);
                        cmd.Parameters.AddWithValue("@price", modulePrice);
                        cmd.Parameters.AddWithValue("@workType", estTypeID);
                        cmd.Parameters.AddWithValue("@modType", modType);

                        conn.Open();
                        int rows = (int)cmd.ExecuteNonQuery();
                        if(rows == 1)
                            result = subModuleID;
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error adding AIMM subcontractor module {subModuleID}: {ex.Message}";
                LogIt.LogError(msg);
                Status = msg;
            }
            return result;
        }

        /// <summary>
        /// Adds a DEMO work order to database
        /// </summary>
        /// <param name="moduleType"><see cref="moduleTypes"/> enum value indicating what type of module</param>
        /// <param name="custID"></param>
        /// <param name="woID">Work order number to build job-work order key</param>
        /// <param name="jobModules">DataTable of modules for the job</param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool add_aimm_work_order(int custID, int woID, DataTable jobModules, float laborRate, float burdenRate, float commRate, string connectionString)
        {
            bool result = false;
            msg = $"Adding AIMM work order to job {jobID}";
            LogIt.LogInfo(msg);
            Status = msg;

            string cust = get_job_customer_name(connectionString);
            string jobWorkOrderKey = $"{jobID}-W{woID.ToString("00")}";
            int coID = 1;
            DateTime createDate = DateTime.Today;

            int workType = (int)jobModules.Rows[0]["EstimateTypeID"];
            int broadCat = get_broad_category(workType, connectionString);
            string woDesc = broadCat == 16 ? "POOL DEMO" : "DEMO";

            // get estimate totals from modules
            var estMtlsCost = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotaEquipMaterialsCost"));
            var estLbrHrs = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotalLaborHours"));
            var estLbrCost = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotalLaborCost"));
            var estBurden = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("ProjEstBurden"));
            var estManDays = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("ManDays"));
            var estGPMD = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("GPMD"));
            var estGP = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("GP"));
            var estCommis = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("SalesCommission"));
            var estPrice = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("EstimatePrice"));

            // insert work order
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "POL.AddAimmWorkOrder";

                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.AddWithValue("@woID", jobWorkOrderKey);
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        cmd.Parameters.AddWithValue("@custID", custID);
                        cmd.Parameters.AddWithValue("@cust", cust);
                        cmd.Parameters.AddWithValue("@woDesc", woDesc);
                        cmd.Parameters.AddWithValue("@coID", coID);
                        cmd.Parameters.AddWithValue("@broadCatID", broadCat);
                        cmd.Parameters.AddWithValue("@workTypeID", workType);
                        cmd.Parameters.AddWithValue("@mtlsCost", estMtlsCost);
                        cmd.Parameters.AddWithValue("@lbrHours", estLbrHrs);
                        cmd.Parameters.AddWithValue("@lbrCost", estLbrCost);
                        cmd.Parameters.AddWithValue("@estPrice", estPrice);
                        cmd.Parameters.AddWithValue("@manDays", estManDays);
                        cmd.Parameters.AddWithValue("@laborRate", laborRate);
                        cmd.Parameters.AddWithValue("@burdenRate", burdenRate);
                        cmd.Parameters.AddWithValue("@estBurden", estBurden);
                        cmd.Parameters.AddWithValue("@GPMD", estGPMD);
                        cmd.Parameters.AddWithValue("@GP", estGP);
                        cmd.Parameters.AddWithValue("@commRate", commRate);
                        cmd.Parameters.AddWithValue("@commis", estCommis);
                        cmd.Parameters.AddWithValue("@moduleCount", jobModules.Rows.Count);

                        conn.Open();
                        int rows = cmd.ExecuteNonQuery();
                        if(rows == 1)
                        {
                            //msg = $"Added work order to job {jobID}";
                            //LogIt.LogInfo(msg);
                            result = true;
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error adding work order to job {jobID}: {ex.Message}";
                LogIt.LogError(msg);
                Status = msg;
            }

            // add the materials
            if(result)
                result = add_aimm_work_order_materials(jobWorkOrderKey, jobModules, connectionString);

            // update the modules with work order ID
            if(result)
                result = assign_work_order_to_modules(jobWorkOrderKey, jobModules, connectionString);

            return result;
        }

        /// <summary>
        /// Add dump and base materials for work order demo modules
        /// </summary>
        /// <param name="workOrderID"></param>
        /// <param name="modules"><see cref="DataTable"/> containing demo modules for work order</param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool add_aimm_work_order_materials(string workOrderID, DataTable modules, string connectionString)
        {
            bool result = false;
            bool thisOne = false;
            int dumpAndBaseID = 1139;
            DateTime createDate = DateTime.Today;

            msg = $"Adding materials for work order {workOrderID} to job {jobID}";
            LogIt.LogInfo(msg);
            Status = msg;

            foreach(DataRow row in modules.Rows)
            {
                float mtlCost;
                //mtlCost = (float)row["TotaEquipMaterialsCost"];
                mtlCost = row.Field<float>("TotaEquipMaterialsCost");
                string moduleID;
                //moduleID = (string)row["ProjectEstimateCommonID"];
                moduleID = row.Field<string>("ProjectEstimateCommonID");
                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        string cmdText = "INSERT INTO MLG.POL.tblProjectFinalMatEquip (ProjectFinalID, JobID, BuildingMaterialID, "
                                       + "OtherMaterial, CostEach, Quantity, TotalCost, Notes, EnteredDate, Correction, JobErrorID) "
                                       + $"VALUES (@woID, @jobID, @bldgMatID, null, @cost, 1, @cost, null, @entDate, 0, null)";

                        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                        {
                            cmd.CommandType = CommandType.Text;
                            cmd.Parameters.AddWithValue("@woID", workOrderID);
                            cmd.Parameters.AddWithValue("@jobID", jobID);
                            cmd.Parameters.AddWithValue("@bldgMatID", dumpAndBaseID);
                            cmd.Parameters.AddWithValue("@cost", mtlCost);
                            cmd.Parameters.AddWithValue("@entDate", createDate);
                            conn.Open();
                            int rows = cmd.ExecuteNonQuery();
                            if(rows == 1)
                            {
                                msg = $"Added materials for work order {workOrderID} to job {jobID}";
                                //LogIt.LogInfo(msg);
                                thisOne = true;
                                Status = msg;
                            }
                        }
                    }
                }
                catch(Exception ex)
                {
                    msg = $"Error adding materials for work order {workOrderID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                    thisOne = false;
                }

                // if any one is true, return true
                result = result || thisOne;
            }
            return result;
        }

        /// <summary>
        /// Returns module price allocated between in-house and contractor portions
        /// </summary>
        /// <param name="projectRange"></param>
        /// <param name="manDays"></param>
        /// <param name="mtlCost"></param>
        /// <param name="subCost"></param>
        /// <param name="ihModulePrice"></param>
        /// <param name="subModulePrice"></param>
        /// <returns>KeyValuePair of in-house and sub prices</returns>
        private KeyValuePair<float, float> allocate_ih_and_sub(dynamic projectRange, float modulePrice, float manDays, float mtlCost, float subCost)
        {
            KeyValuePair<float, float> result = new KeyValuePair<float, float>(0F, 0F);
            dynamic row = get_first_unused_project_line(projectRange);

            if(row != null)
            {
                // get the in-house portion
                var mdOK = set_man_days(row, manDays);
                var mcOK = set_material_cost(row, mtlCost);
                var ihPrice = 0F;
                if(mdOK == manDays && mcOK == mtlCost)
                {
                    ihPrice = get_row_total(row);
                    set_man_days(row, null);
                    set_material_cost(row, null);
                }

                // get the subcontractor portion
                var subOK = set_sub_cost(row, subCost);
                var subPrice = 0F;
                if(subOK == subCost)
                {
                    subPrice = get_row_total(row);
                    set_sub_cost(row, null);
                }

                // if IH + SUB prices don't equal original price, adjust sub price to correct
                if(ihPrice + subPrice != modulePrice)
                    subPrice = modulePrice - ihPrice;

                result = new KeyValuePair<float, float>(ihPrice, subPrice);
            }
            return result;
        }

        /// <summary>
        /// Sets ProjectFinalID to workOrderID in work order demo modules
        /// </summary>
        /// <param name="workOrderID"></param>
        /// <param name="jobModules"></param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool assign_work_order_to_modules(string workOrderID, DataTable jobModules, string connectionString)
        {
            bool result = false;
            msg = $"Assigning work order {workOrderID} to modules for job {jobID}";
            LogIt.LogInfo(msg);
            Status = msg;

            if(jobModules.Rows.Count != 0)
            {
                // convert module IDs to comma-separated string
                string[] strArr = jobModules.AsEnumerable().Select(r => "'" + Convert.ToString(r["ProjectEstimateCommonID"]) + "'").ToArray();
                string modIDs = string.Join(",", strArr);

                try
                {
                    using(SqlConnection conn = new SqlConnection(connectionString))
                    {
                        string cmdText = "UPDATE MLG.POL.tblProjectEstimateCommon "
                                       + $"SET ProjectFinalID = '{workOrderID}' WHERE ProjectEstimateCommonID IN ({modIDs})";

                        conn.Open();
                        using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                        {
                            cmd.CommandType = CommandType.Text;
                            int rows = cmd.ExecuteNonQuery();
                            if(rows >= 1)
                            {
                                msg = $"Added work order {workOrderID} to demo module(s) for job {jobID}";
                                LogIt.LogInfo(msg);
                                Status = msg;
                                result = true;
                            }
                        }
                        conn.Close();
                    }
                }
                catch(Exception ex)
                {
                    msg = $"Error updating demo module(s) for work order {workOrderID}: {ex.Message}";
                    LogIt.LogError(msg);
                    Status = msg;
                }
            }
            return result;
        }

        /// <summary>
        /// Returns job description from 3 largest APPROVED project types
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
                    if(project_is_approved())
                    {
                        float ttl = get_project_total();
                        if(ttl > 0)
                        {
                            string projDesc = oXl.RangeOffset(0, 1, 1, 1).Value;
                            modList.Add(new KeyValuePair<string, float>(projDesc, ttl));
                        }
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

        /// <summary>
        /// Returns broad category from supplied work type
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private int get_broad_category(int workType, string connectionString)
        {
            int response = 0;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting broad category for work type {workType}");
                    string cmdText = "SELECT WorkTypeBroadCatergoryID FROM MLG.POL.tblEstimateTypes where EstimateTypeID = @estType";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@estType", workType);
                        response = (int)cmd.ExecuteScalar();
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting broad category for estimate type {workType}: {ex.Message}");
            }
            return response;
        }


        /// <summary>
        /// Returns work type from project row AIMM category
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private int get_estimate_type(dynamic row)
        {
            int response = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["aimm_category_range"]));
                string temp = (cell.Value ?? "").ToString();
                int.TryParse(get_last_paren_segment(temp), out response);
                cell = null;
            }
            catch(Exception ex)
            {
                msg = $"Error getting row estimate type: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return response;
        }

        private dynamic get_first_unused_project_line(dynamic projectRange)
        {
            dynamic result = null;
            int row = projectRange.Rows.Count + 1;
            string stepText = "";
            try
            {
                while(stepText == "")
                {
                    row--;
                    stepText = projectRange.Rows[row].Cells[2].Value ?? "";
                }
                if(row < projectRange.Rows.Count)
                {
                    row++;
                    result = projectRange.Rows[row];
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting unused project row: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// get customer name for supplied job number
        /// </summary>
        /// <param name="connectionString"></param>
        /// <returns>customer for job or null if not found</returns>
        private string get_job_customer_name(string connectionString)
        {
            string cust = "";
            msg = $"Getting customer for job {jobID}";
            LogIt.LogError(msg);
            Status = msg;
            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "SELECT TheCustomerSimple FROM MLG.dbo.vJobs where JobID = @jobID";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        cust = (string)cmd.ExecuteScalar();
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting customer for job {jobID}: {ex.Message}";
                LogIt.LogError(msg);
                Status = msg;
            }
            return cust;
        }

        private DataTable get_job_modules(moduleTypes moduleType, string connectionString)
        {
            DataTable result = new DataTable();
            string demoTypeWhere = "";
            switch(moduleType)
            {
                case moduleTypes.demo:
                    demoTypeWhere = $" AND EstimateTypeID in ({demo_types})";
                    break;
                case moduleTypes.poolDemo:
                    demoTypeWhere = $" AND EstimateTypeID in ({pool_demo_types})";
                    break;
                default:
                    break;
            }
            try
            {
                // get data using DataTable
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "SELECT * FROM MLG.POL.tblProjectEstimateCommon "
                                   + $"WHERE JobID = '{jobID}'{demoTypeWhere}";
                    conn.Open();
                    using(SqlDataAdapter dap = new SqlDataAdapter(cmdText, conn))
                    {
                        dap.Fill(result);
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting modules for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }

            return result;
        }

        private DataTable get_job_sub_modules(string connectionString)
        {
            DataTable result = new DataTable();
            LogIt.LogInfo($"Getting sub modules for job {jobID}");
            try
            {
                // get data using DataTable
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "SELECT * FROM MLG.POL.tblSubModuleWOSubModule WHERE JobID = @jobID";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        using(SqlDataAdapter dap = new SqlDataAdapter(cmd))
                        {
                            dap.Fill(result);
                        }
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting sub modules for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }

            return result;
        }

        private DataTable get_job_work_orders(string connectionString)
        {
            DataTable result = new DataTable();
            LogIt.LogInfo($"Getting work orders for job {jobID}");
            try
            {
                // get data using DataTable
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "SELECT * FROM MLG.POL.tblProjectFinal WHERE JobID = @jobID";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@jobID", jobID);
                        using(SqlDataAdapter dap = new SqlDataAdapter(cmd))
                        {
                            dap.Fill(result);
                        }
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting work orders for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
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
                    conn.Open();
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
                    conn.Close();
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
        /// Gets job total at currently set percent level
        /// </summary>
        /// <returns></returns>
        private float get_job_total()
        {
            float result = 0;
            dynamic cell = oXl.GetSecondaryRange(rangeNames["job_total"]);
            if(cell != null)
            {
                var temp = cell.Value;
                float.TryParse(temp.ToString(), out result);
            }
            cell = null;
            return result;
        }

        /// <summary>
        /// Returns the project row containing the largest IN-HOUSE total
        /// </summary>
        /// <param name="projectRange"></param>
        /// <returns></returns>
        private dynamic get_largest_in_house_module(dynamic projectRange)
        {
            dynamic result = null;
            try
            {
                // get the totals, sub cost & aimm category columns for the project
                dynamic totals = oXl.GetIntersect(projectRange.EntireRow, oXl.GetSecondaryRange(rangeNames["project_total_range"]));
                dynamic subCosts = oXl.GetIntersect(projectRange.EntireRow, oXl.GetSecondaryRange(rangeNames["sub_cost_range"]));
                dynamic aimmCats = oXl.GetIntersect(projectRange.Entirerow, oXl.GetSecondaryRange(rangeNames["aimm_category_range"]));
                if(totals != null && subCosts != null && aimmCats != null)
                {
                    // get the cell containing the largest amount (exclude rows with sub cost)
                    int maxCell = 0;
                    float maxVal = 0;
                    for(int i = 1; i <= totals.Cells.Count; i++)
                    {
                        var price = (totals.Cells[i].Value ?? 0).ToString();
                        var sub = (subCosts.Cells[i].Value ?? 0).ToString();
                        var cat = get_last_paren_segment((aimmCats.Cells[i].Value ?? ""));

                        float val = 0;
                        float subCost = 0;
                        float.TryParse(price, out val);
                        float.TryParse(sub, out subCost);

                        if(subCost == 0)
                        {
                            if(cat == "" | !demo_types.Contains(cat))
                            {
                                if(cat == "" | !pool_demo_types.Contains(cat))
                                {
                                    if(val > maxVal)
                                    {
                                        maxCell = i;
                                        maxVal = Convert.ToSingle(val);
                                    }
                                }
                            }
                        }

                        //    if(subCost == 0 && !demo_types.Contains(cat) && !pool_demo_types.Contains(cat))
                        //{
                        //    if(val > maxVal)
                        //    {
                        //        maxCell = i;
                        //        maxVal = Convert.ToSingle(val);
                        //    }
                        //}

                    }
                    dynamic bigCell = totals.Cells[maxCell];
                    if(bigCell != null)
                    {
                        // get the project row for the cell
                        result = oXl.GetIntersect(bigCell.EntireRow, projectRange);
                        bigCell = null;
                    }
                    totals = null;
                    subCosts = null;
                    aimmCats = null;


                    //dynamic bigCell = oXl.GetMaxCellInRange(totals, subCosts, "=0");
                    //if(bigCell != null)
                    //{
                    //    // get the project row for the cell
                    //    result = oXl.GetIntersect(bigCell.EntireRow, projectRange);
                    //    bigCell = null;
                    //}
                    //totals = null;
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

        private string get_last_paren_segment(string textToSearch)
        {
            try
            {
                string[] temp = textToSearch.Split(new string[] { "(", ")" }, StringSplitOptions.RemoveEmptyEntries);
                return temp[temp.GetUpperBound(0)];
            }
            catch(Exception)
            {
                return "";
            }
        }

        /// <summary>
        /// Returns project row man days
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float get_man_days(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["man_days_range"]));
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
        /// Returns project row material cost
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float get_material_cost(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["material_cost_range"]));
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
        /// Returns total for selected project
        /// </summary>
        /// <remarks>
        /// - a project must be selected in Excel class
        /// - assumes project total cell is on same row as first project row
        /// </remarks>
        /// <returns>float</returns>
        private float get_project_total()
        {
            float result = 0;
            dynamic cell = oXl.GetIntersect(oXl.Range.Rows(1).EntireRow,
                                            oXl.GetSecondaryRange(rangeNames["project_approval_range"]));
            if(cell != null)
            {
                string ttl = (cell.Value ?? "").ToString();
                float.TryParse(ttl, out result);
            }
            cell = null;
            return result;
        }

        /// <summary>
        /// Get rate for supplied type and date
        /// </summary>
        /// <param name="rateType"><see cref="rateTypes"/> enum value indicating what rate to retrieve</param>
        /// <param name="rateDate">Date to retrieve rate for</param>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private float get_rate(rateTypes rateType, DateTime rateDate, string connectionString)
        {
            float result = 0;
            string rateDesc = "";
            string rateTable = "";
            string rateField = "";
            switch(rateType)
            {
                case rateTypes.burdenRate:
                    rateDesc = "burden";
                    rateTable = "tblBurdenRate";
                    rateField = "BurdenRate";
                    break;
                case rateTypes.commissionRate:
                    rateDesc = "commission";
                    rateTable = "tblComissRate";
                    rateField = "ComissRate";
                    break;
                case rateTypes.laborRate:
                    rateDesc = "labor";
                    rateTable = "tblLaborRate";
                    rateField = "LaborRate";
                    break;
                default:
                    break;
            }

            try
            {
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    LogIt.LogInfo($"Getting {rateDesc} rate for {rateDate.ToShortDateString()}");
                    string cmdText = $"SELECT {rateField} FROM MLG.POL.{rateTable} WHERE @date BETWEEN StartDate and EndDate";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.Parameters.AddWithValue("@date", rateDate);
                        result = (float)cmd.ExecuteScalar();
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                LogIt.LogError($"Error getting {rateDesc} rate for {rateDate.ToShortDateString()}: {ex.Message}");
            }
            return result;
        }

        /// <summary>
        /// Returns amount of specs sheet row total price
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float get_row_total(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["project_total_range"]));
                float.TryParse((cell.Value ?? "0").ToString(), out result);
                cell = null;
            }
            catch(Exception ex)
            {
                msg = $"Error getting row total: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
                result = 0;
            }
            return result;
        }

        private string get_setting(XmlDocument doc, string settingName)
        {
            string result = "";
            try
            {
                result = ((XmlElement)doc.SelectSingleNode($"/Settings/setting[@name='{settingName}']")).GetAttribute("value");
            }
            catch(Exception)
            {
            }
            return result;
        }

        private string get_string_segment(string textToSearch, string keyText, string delimiter)
        {
            string result = "";
            if(textToSearch.Contains(keyText))
            {
                result = textToSearch.Substring(textToSearch.IndexOf(keyText) + keyText.Length);
                if(result.Contains(delimiter))
                    result = result.Remove(result.IndexOf(delimiter));
            }
            return result;
        }

        /// <summary>
        /// Returns project row subcontractor cost
        /// </summary>
        /// <param name="row"></param>
        /// <returns></returns>
        private float get_sub_cost(dynamic row)
        {
            float result = 0;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["sub_cost_range"]));
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
        /// Import projects for a job from Excel spreadsheet
        /// </summary>
        /// <param name="custID"></param>
        /// <param name="jobDesc"></param>
        /// <param name="salesRep"></param>
        /// <param name="estimator"></param>
        /// <param name="saleDate"></param>
        /// <param name="laborRate"></param>
        /// <param name="burdenRate"></param>
        /// <param name="commRate"></param>
        /// <returns></returns>
        private bool import_projects(int custID, string jobDesc, int salesRep, int estimator, DateTime saleDate, float laborRate, float burdenRate, float commRate)
        {
            bool result = false;

            // get and sort a list of projects and options
            List<string> projects = oXl.GetNamedRanges(new string[] { "specs_project_*", "specs_option_*" });
            List<string> projects_and_options = sort_projects_and_options(projects);

            // iterate list of projects and options, add job and modules
            foreach(string rngName in projects_and_options)
            {
                // get the project/option range, continue if approved
                oXl.GetRange(rngName);
                if(project_is_approved())
                {
                    dynamic thisProject = oXl.Range;
                    dynamic largestModule = null;
                    bool aModuleHasNoManDays = false;
                    float saveMtlCost = 0;
                    float saveModulePrice = 0;
                    string pNo = rngName.Substring(rngName.LastIndexOf("_") + 1);
                    string modPrefix = rngName.Contains("project") ? $"P{pNo}: " : $"O{pNo}: ";

                    foreach(dynamic module in thisProject.Rows)
                    {
                        // process this module if it has a total price
                        float modulePrice = get_row_total(module);
                        if(modulePrice != 0)
                        {
                            string modDesc = modPrefix + module.Cells(2).Value;
                            float manDays = get_man_days(module);
                            float mtlCost = get_material_cost(module);
                            float subCost = get_sub_cost(module);
                            string moduleID = "";
                            string subModuleID = "";

                            // where does this module go?
                            bool isIhOnly = (subCost == 0);
                            bool isSubOnly = (subCost != 0 && manDays == 0);
                            bool isIhAndSub = (subCost != 0 && manDays != 0);

                            // if in-house and no man days, save material cost for later
                            if(isIhOnly && manDays == 0 && mtlCost != 0)
                            {
                                largestModule = get_largest_in_house_module(thisProject);
                                aModuleHasNoManDays = (largestModule != null);
                                if(aModuleHasNoManDays)
                                {
                                    saveMtlCost = mtlCost;
                                    saveModulePrice = modulePrice;
                                    msg = $"Saving material cost {mtlCost} and module price {modulePrice} for later entry";
                                    LogIt.LogInfo(msg);
                                    Status = msg;
                                }
                            }
                            else
                            {
                                // if we're saving a materials-only cost/price and this is 
                                // the largest row, add them to this row's cost/price
                                if(aModuleHasNoManDays && saveMtlCost != 0 && saveModulePrice != 0
                                   && module.Address == largestModule.Address)
                                {
                                    msg = $"Applying saved material cost/price ({saveMtlCost} / {saveModulePrice}) to module cost/price ({mtlCost} / {modulePrice}";
                                    mtlCost += saveMtlCost;
                                    modulePrice += saveModulePrice;
                                    msg += $", new cost/price = {mtlCost} / {modulePrice}";
                                    LogIt.LogInfo(msg);
                                    Status = msg;
                                    saveMtlCost = 0;
                                    saveModulePrice = 0;
                                    aModuleHasNoManDays = false;
                                    largestModule = null;
                                }

                                // add job to aimm if we haven't already
                                if(jobID == "")
                                {
                                    jobID = add_aimm_job(custID, jobDesc, salesRep, estimator, saleDate, laborRate, burdenRate, commRate, connString);
                                    if(jobID == "")
                                    {
                                        msg = "Error adding AIMM job, estimate NOT imported.";
                                        Status = msg;
                                        LogIt.LogError(msg);
                                        return false;
                                    }
                                    else
                                    {
                                        msg = $"Added AIMM job, id = {jobID}.";
                                        Status = msg;
                                        LogIt.LogInfo(msg);
                                    }
                                }

                                // get estimate type from aimm category
                                int estTypeID = get_estimate_type(module);


                                // if in-house AND sub, allocate price between them,
                                // with IH portion getting the materials cost
                                float ihModulePrice = isIhOnly ? modulePrice : 0;
                                float subModulePrice = isSubOnly ? modulePrice : 0;
                                float subMtlCost = isSubOnly ? mtlCost : 0;
                                if(isIhAndSub)
                                {
                                    KeyValuePair<float, float> prices =
                                        allocate_ih_and_sub(thisProject, modulePrice, manDays, mtlCost, subCost);
                                    ihModulePrice = prices.Key;
                                    subModulePrice = prices.Value;
                                }

                                // add the standard module if needed
                                if(isIhOnly || isIhAndSub)
                                {
                                    modulesInJob++;
                                    moduleID = add_aimm_module(custID, modulesInJob, modDesc, manDays, mtlCost, ihModulePrice, estTypeID, salesRep, estimator, laborRate, burdenRate, commRate, connString);

                                }
                                // add the sub module if needed
                                if(isSubOnly || isIhAndSub)
                                {
                                    // if this module has any in-house material cost, none of it goes here, it all goes to IH module
                                    subModulesInJob++;
                                    subModuleID = add_aimm_sub_module(custID, subModulesInJob, modDesc, subMtlCost, subCost, subModulePrice, estTypeID, salesRep, commRate, connString);

                                }
                            }
                        }
                    }
                    thisProject = null;
                    largestModule = null;
                    result = true;
                }
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
        /// Returns whether selected project is approved
        /// </summary>
        /// <remarks>
        /// - a project must be selected in Excel class
        /// - assumes approval cell is 1 row down from first project row
        /// </remarks>
        /// <returns>Boolean</returns>
        private bool project_is_approved()
        {
            bool result = false;
            dynamic cell = oXl.GetIntersect(oXl.Range.Rows(2).EntireRow,
                                            oXl.GetSecondaryRange(rangeNames["project_approval_range"]));
            if(cell != null)
            {
                result = (cell.Value == "APPROVED");
            }
            cell = null;
            return result;
        }

        /// <summary>
        /// Sets the current job percent to the supplied percentage
        /// </summary>
        /// <param name="newPercent">Desired percent expressed as number (ex: 100, 98.5). Will be divided by 100 before setting cell value.</param>
        private bool set_job_percent(float newPercent)
        {
            bool response = false;
            if(newPercent != 0)
            {
                dynamic cell = oXl.GetSecondaryRange(rangeNames["percent_hours"]);
                if(cell != null)
                {
                    try
                    {
                        cell.Value = newPercent / 100;
                        response = true;
                    }
                    catch(Exception ex)
                    {
                        msg = $"Error setting job percent: {ex.Message}";
                        Status = msg;
                        LogIt.LogError(msg);
                    }
                    cell = null;
                }
            }
            return response;
        }

        /// <summary>
        /// Sets project row man days, returns value set if successful.
        /// </summary>
        /// <param name="row">The project row to set</param>
        /// <param name="value">The value to set or null to clear value</param>
        /// <returns></returns>
        private float? set_man_days(dynamic row, float? value)
        {
            float? result = -.01234F;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["man_days_range"]));
                if(value == null)
                    cell.ClearContents();
                else
                    cell.Value = value;
                cell = null;
                result = value;
            }
            catch(Exception ex)
            {
                msg = $"Error setting row man days: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// Sets project row material cost, returns value set if successful.
        /// </summary>
        /// <param name="row">The project row to set</param>
        /// <param name="value">The value to set or null to clear value</param>
        /// <returns></returns>
        private float? set_material_cost(dynamic row, float? value)
        {
            float? result = -.01234F;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["material_cost_range"]));
                if(value == null)
                    cell.ClearContents();
                else
                    cell.Value = value;
                cell = null;
                result = value;
            }
            catch(Exception ex)
            {
                msg = $"Error setting row material cost: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }

        /// <summary>
        /// Sets project row subcontractor cost, returns value set if successful.
        /// </summary>
        /// <param name="row">The project row to set</param>
        /// <param name="value">The value to set or null to clear value</param>
        /// <returns></returns>
        private float? set_sub_cost(dynamic row, float? value)
        {
            float? result = -.01234F;
            try
            {
                dynamic cell = oXl.GetIntersect(row.EntireRow, oXl.GetSecondaryRange(rangeNames["sub_cost_range"]));
                if(value == null)
                    cell.ClearContents();
                else
                    cell.Value = value;
                cell = null;
                result = value;
            }
            catch(Exception ex)
            {
                msg = $"Error setting row material cost: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
            }
            return result;
        }


        /// <summary>
        /// Sort list of projects and options
        /// </summary>
        /// <param name="mixedup_list"></param>
        /// <returns>List of projects and options in proper order</returns>
        private List<string> sort_projects_and_options(List<string> mixedup_list)
        {
            List<string> newList = new List<string>();

            // first get all the projects
            for(int i = 1; i <= mixedup_list.Count; i++)
            {
                string projectRange = $"specs_project_{i}";
                if(mixedup_list.Contains(projectRange))
                    newList.Add(projectRange);
            }

            // now get all the options
            for(int i = 1; i <= mixedup_list.Count; i++)
            {
                string optionRange = $"specs_option_{i}";
                if(mixedup_list.Contains(optionRange))
                    newList.Add(optionRange);
            }
            return newList;
        }

        /// <summary>
        /// Sets job totals fields from modules
        /// </summary>
        /// <param name="connectionString"></param>
        /// <returns></returns>
        private bool update_aimm_job_totals(string connectionString)
        {
            bool result = false;
            float jobTotalPrice = 0;
            float jobTotalPriceMinusLossAndNegMods = 0;

            // module fields
            float modsPrice = 0;
            float modsMtls = 0;
            float modsLbrHrs = 0;
            float modsLbrCost = 0;
            float modsLbrBurden = 0;
            float modsCommis = 0;
            float modsGp = 0;
            float modsManDays = 0;
            float effModCommisRate = 0;
            float modsGpmd = 0;

            // work order fields
            float wosPrice = 0;
            float wosMtls = 0;
            float wosLbrHrs = 0;
            float wosLbrCost = 0;
            float wosLbrBurden = 0;
            float wosManDays = 0;
            float wosCommis = 0;
            float wosGp = 0;
            float wosGpmd = 0;

            // sub module fields
            float subsPrice = 0;
            float subsCost = 0;
            float subsMtlsCost = 0;
            float subsGp = 0;
            float subsCommis = 0;
            float effSubModCommisRate = 0;

            // get totals for modules
            try
            {
                LogIt.LogInfo($"Getting module totals for job {jobID}");
                using(DataTable jobModules = get_job_modules(moduleTypes.all, connectionString))
                {
                    if(jobModules.Rows.Count > 0)
                    {
                        //var estMtlsCost = jobModules.Rows.Cast<DataRow>().Sum(dr => (float)dr[jobModules.Columns["TotalEquipMaterialsCost"]]);
                        //var estMtlsCost2 = jobModules.Rows.Cast<DataRow>().Sum(dr => dr.Field<float>("TotalEquipMaterialsCost"));
                        //var estMtlsCost = jobModules.AsEnumerable().Sum(dr => dr.Field<float?>("TotalEquipMaterialsCost"));
                        modsPrice = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("EstimatePrice"));
                        modsMtls = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotaEquipMaterialsCost"));
                        modsLbrHrs = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotalLaborHours"));
                        modsLbrCost = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("TotalLaborCost"));
                        modsLbrBurden = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("ProjEstBurden"));
                        modsCommis = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("SalesCommission"));
                        modsGp = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("GP"));
                        modsManDays = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("ManDays"));
                        effModCommisRate = modsGp > 0 ? modsCommis / modsGp : 0;
                        modsGpmd = modsManDays > 0 ? modsGp / modsManDays : 0;
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting module totals for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
                return false;
            }

            // get totals for work orders
            try
            {
                LogIt.LogInfo($"Getting work order totals for job {jobID}");
                using(DataTable jobModules = get_job_work_orders(connectionString))
                {
                    if(jobModules.Rows.Count > 0)
                    {
                        wosPrice = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjPrice"));
                        wosMtls = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjMATCost"));
                        wosLbrHrs = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjLABHrs"));
                        wosLbrCost = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjLABCost"));
                        wosLbrBurden = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjBurdCost"));
                        wosCommis = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjComiss"));
                        wosGp = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjGP"));
                        wosManDays = jobModules.AsEnumerable().Sum(dr => dr.Field<float>("AltProjMD"));
                        wosGpmd = wosManDays > 0 ? wosGp / wosManDays : 0;
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting work order totals for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
                return false;
            }

            // get totals for sub modules
            try
            {
                LogIt.LogInfo($"Getting sub module totals for job {jobID}");
                using(DataTable subModules = get_job_sub_modules(connectionString))
                {
                    if(subModules.Rows.Count > 0)
                    {
                        subsPrice = subModules.AsEnumerable().Sum(dr => dr.Field<float>("EstimatePrice"));
                        subsCost = subModules.AsEnumerable().Sum(dr => dr.Field<float>("SubCost"));
                        subsMtlsCost = subModules.AsEnumerable().Sum(dr => dr.Field<float>("MATCost"));
                        subsGp = subModules.AsEnumerable().Sum(dr => dr.Field<float>("GP"));
                        subsCommis = subModules.AsEnumerable().Sum(dr => dr.Field<float>("SalesCommission"));
                        effSubModCommisRate = subsGp > 0 ? subsCommis / subsGp : 0;
                    }
                }
            }
            catch(Exception ex)
            {
                msg = $"Error getting sub module totals for job {jobID}: {ex.Message}";
                Status = msg;
                LogIt.LogError(msg);
                return false;
            }

            jobTotalPrice = modsPrice + subsPrice;
            jobTotalPriceMinusLossAndNegMods = jobTotalPrice;

            // update the job totals
            try
            {
                LogIt.LogInfo($"Updating job totals for job {jobID}");
                using(SqlConnection conn = new SqlConnection(connectionString))
                {
                    string cmdText = "UPDATE MLG.POL.tblJobs SET SumEstimatePrice = @modsPrice, SumMaterialsCost = @modsMtls, SumLaborHours = @modsLbrHrs, "
                                   + "SumLaborCost = @modsLbrCost, SumComiss = @modsCommis, SumGP = @modsGp, SumManDays = @modsManDays, SumGPMD = @modsGpmd, "
                                   + "SumProjEstBurden = @modsLbrBurden, GrandTotalJobPrice = @jobTotalPrice, SumTargetMAT = @wosMtls, SumTargetLABHours = @wosLbrHrs, "
                                   + "SumTargetLAB = @wosLbrCost, SumTargetBurden = @wosLbrBurden, SumTargetManDays = @wosManDays, SumTargetGPMD = @wosGpmd, "
                                   + "SumTargetGP = @wosGp, SumTargetCommiss = @wosCommis, SumTargetPrice = @wosPrice, SubModCostTotal = @subsCost, "
                                   + "SubModMATTotal = @subsMtlsCost, SubModGPTotal = @subsGp, SubModTotalCom = @subsCommis, SubModModTotalPrice = @subsPrice, "
                                   + "EffectiveModCommRate = @effModCommisRate, EffectiveSubModCommRate = @effSubModCommisRate, "
                                   + "TotalPriceMinusLossAndNegMods = @jobTotalPriceMinusLossAndNegMods WHERE JobID = @jobID;";
                    conn.Open();
                    using(SqlCommand cmd = new SqlCommand(cmdText, conn))
                    {
                        cmd.CommandType = CommandType.Text;
                        cmd.Parameters.AddWithValue("@modsPrice", modsPrice);
                        cmd.Parameters.AddWithValue("@modsMtls", modsMtls);
                        cmd.Parameters.AddWithValue("@modsLbrHrs", modsLbrHrs);
                        cmd.Parameters.AddWithValue("@modsLbrCost", modsLbrCost);
                        cmd.Parameters.AddWithValue("@modsCommis", modsCommis);
                        cmd.Parameters.AddWithValue("@modsGp", modsGp);
                        cmd.Parameters.AddWithValue("@modsManDays", modsManDays);
                        cmd.Parameters.AddWithValue("@modsGpmd", modsGpmd);
                        cmd.Parameters.AddWithValue("@modsLbrBurden", modsLbrBurden);
                        cmd.Parameters.AddWithValue("@jobTotalPrice", jobTotalPrice);
                        cmd.Parameters.AddWithValue("@wosMtls", wosMtls);
                        cmd.Parameters.AddWithValue("@wosLbrHrs", wosLbrHrs);
                        cmd.Parameters.AddWithValue("@wosLbrCost", wosLbrCost);
                        cmd.Parameters.AddWithValue("@wosLbrBurden", wosLbrBurden);
                        cmd.Parameters.AddWithValue("@wosManDays", wosManDays);
                        cmd.Parameters.AddWithValue("@wosGpmd", wosGpmd);
                        cmd.Parameters.AddWithValue("@wosGp", wosGp);
                        cmd.Parameters.AddWithValue("@wosCommis", wosCommis);
                        cmd.Parameters.AddWithValue("@wosPrice", wosPrice);
                        cmd.Parameters.AddWithValue("@subsCost", subsCost);
                        cmd.Parameters.AddWithValue("@subsMtlsCost", subsMtlsCost);
                        cmd.Parameters.AddWithValue("@subsGp", subsGp);
                        cmd.Parameters.AddWithValue("@subsCommis", subsCommis);
                        cmd.Parameters.AddWithValue("@subsPrice", subsPrice);
                        cmd.Parameters.AddWithValue("@effModCommisRate", effModCommisRate);
                        cmd.Parameters.AddWithValue("@effSubModCommisRate", effSubModCommisRate);
                        cmd.Parameters.AddWithValue("@jobTotalPriceMinusLossAndNegMods", jobTotalPriceMinusLossAndNegMods);
                        cmd.Parameters.AddWithValue("@jobID", jobID);

                        //conn.Open();
                        int rows = cmd.ExecuteNonQuery();
                        if(rows == 1)
                        {
                            msg = $"Updated job totals for job {jobID}";
                            LogIt.LogInfo(msg);
                            Status = msg;
                            result = true;
                        }
                    }
                    conn.Close();
                }
            }
            catch(Exception ex)
            {
                msg = $"Error updating job {jobID}: {ex.Message}";
                LogIt.LogError(msg);
            }
            return result;
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
