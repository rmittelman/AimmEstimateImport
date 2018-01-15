using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
//using Excel = Microsoft.Office.Interop.Excel;

namespace AimmEstimateImport
{
    class clsExcel
    {
        Type xlType;
        dynamic xlApp;
        dynamic xlWorkbook;
        dynamic xlWorksheet;
        dynamic xlRange;
        dynamic xlRange2;

        public enum XlFileFormat
        {
            xlOpenXMLWorkbook = 51,
            xlOpenXMLWorkbookMacroEnabled = 52,
            xlExcel8 = 56
        }

        private bool workbookOpen = false;
        public string[] parms;

        public clsExcel(string workbookName, bool visible = true)
        {
            WorkbookName = workbookName;
            Visible = visible;
        }

        public clsExcel()
        {

        }

        ~clsExcel()
        {
            release_objects();
        }

        #region properties

        public bool Visible { get; set; }
        public string WorkbookName { get; set; }
        public string LastError { get; set; }
        public dynamic WorkBook { get { return xlWorkbook; } }
        public dynamic WorkSheet { get { return xlWorksheet; } }
        public dynamic Range { get { return xlRange; } }
        public int Rows { get { return xlRange.Rows.Count; } }
        public int Columns { get { return xlRange.Columns.Count; } }
        public int ActiveRows { get { return (int)xlApp.WorksheetFunction.CountA(xlRange.Columns(1)); } }
        public int ActiveColumns { get { return (int)xlApp.WorksheetFunction.CountA(xlRange.Rows(1)); } }
        public string RangeAddress { get { return xlRange.Address; } }

        #endregion


        #region methods

        /// <summary>
        /// close excel file, save if needed
        /// </summary>
        /// <param name="save"></param>
        /// <returns>boolean indicating success status</returns>
        public bool CloseWorkbook(bool save = false)
        {
            bool result = false;
            try
            {
                try
                {
                    xlWorkbook.Close(save);
                    xlWorkbook = null;
                }
                catch(COMException)
                {
                    // ignore if already closed
                }
                workbookOpen = false;
                LastError = "";
                result = true;
            }
            catch(Exception ex)
            {
                LastError = $"Error closing excel file: {ex.Message}";
                result = false;
            }
            return result;
        }

        /// <summary>
        /// Save workbook to new file name (optionally to different format)
        /// </summary>
        /// <param name="fileName">Full Path Name of file to save</param>
        /// <param name="fileFormat">File format desired.</param>
        /// <returns></returns>
        public bool SaveWorkbookAs(string fileName, XlFileFormat fileFormat = XlFileFormat.xlOpenXMLWorkbook)
        {
            bool result = false;
            try
            {
                xlWorkbook.SaveAs(fileName, fileFormat);
                result = true;

            }
            catch(Exception ex )
            {
                LastError = $"Error saving excel file: {ex.Message}";
                result = false;
            }
            return result;
       }

        /// <summary>
        /// Start Excel and open supplied workbook name, optionally activate requested sheet
        /// </summary>
        /// <param name="fileName">Full path-name of file to open</param>
        /// <param name="sheet">Name or number of worksheet to activate</param>
        /// <returns>boolean indicating success status</returns>
        public bool OpenExcel(string fileName, string sheet = "")
        {
            var xlFile = Path.GetFileName(fileName);
            bool result = false;
            if(File.Exists(fileName))
            {
                int sheetNo = 0;
                try
                {
                    xlType = Type.GetTypeFromProgID("Excel.Application");
                    xlApp = Activator.CreateInstance(xlType);
                    xlApp.Visible = Visible;
                    xlWorkbook = xlApp.Workbooks.Open(fileName);
                    if(sheet == "")
                        xlWorksheet = xlWorkbook.Sheets[1];
                    else if(int.TryParse(sheet, out sheetNo))
                        xlWorksheet = xlWorkbook.Sheets[sheetNo];
                    else
                    {
                        try
                        {
                            xlWorksheet = xlWorkbook.Sheets[sheet];
                        }
                        catch(Exception)
                        {
                            xlWorksheet = xlWorkbook.Sheets[1];
                        }
                    }
                    xlWorksheet.Activate();
                    workbookOpen = true;
                    LastError = "";
                    result = true;
                }
                catch(Exception ex)
                {
                    LastError = $"Error opening Excel file \"{xlFile}\": {ex.Message}";
                    result = false;
                }
            }
            else
            {
                LastError = $"Could not find Excel file \"{xlFile}\"";
                result = false;
            }
            return result;
        }

        public bool CloseExcel()
        {
            bool result = true;
            if(workbookOpen)
                result = CloseWorkbook();

            release_objects();
            return result;
        }

        /// <summary>
        /// Get Range for supplied name or address, optionally select the range
        /// </summary>
        /// <param name="nameOrAddress">Named range or range address</param>
        /// <param name="select">If true, range is selected in workbook</param>
        /// <returns>Boolean indicating success</returns>
        /// <remarks>Sets class' internal range object which other methods depend on</remarks>
        public bool GetRange(string nameOrAddress, bool select = false)
        {
            bool result = false;
            if(xlApp != null && xlWorkbook != null)
            {
                try
                {
                    xlRange = xlApp.Range(nameOrAddress);
                    if(select)
                        xlRange.Activate();

                    result = true;
                    LastError = "";
                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                }
            }
            return result;
        }

        /// <summary>
        /// Return Intersect range of supplied ranges
        /// </summary>
        /// <param name="range1"></param>
        /// <param name="range2"></param>
        /// <returns></returns>
        public dynamic GetIntersect(dynamic range1, dynamic range2)
        {
            dynamic result = null;
            try
            {
                result = xlApp.Intersect(range1, range2);
            }
            catch(Exception ex)
            {
                LastError = ex.Message;
                result = null;
            }
            return result;
        }

        /// <summary>
        /// Get Range for supplied name or address (without affecting current range)
        /// </summary>
        /// <param name="nameOrAddress">Named range or range address</param>
        /// <returns>A dynamic range object</returns>
        public dynamic GetSecondaryRange(string nameOrAddress)
        {
            dynamic result = null;
            if(xlApp != null && xlWorkbook != null)
            {
                try
                {
                    result = xlApp.Range(nameOrAddress);
                    LastError = "";
                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                    result = null;
                }
            }
            return result;
        }

        /// <summary>
        /// Get list of range names matching any of the supplied patterns
        /// </summary>
        /// <param name="namesLike">Array of patterns to match</param>
        /// <returns>List of named range names</returns>
        public List<string> GetNamedRanges(string[] namesLike)
        {
            List<string> results = new List<string>();
            foreach(var nameLike in namesLike)
            {
                results.AddRange(GetNamedRanges(nameLike));
            }
            return results;
        }


        /// <summary>
        /// Get list of range names matching supplied pattern
        /// </summary>
        /// <param name="nameLike">Optional pattern to match</param>
        /// <returns>List of named range names</returns>
        public List<string> GetNamedRanges(string nameLike = "")
        {
            List<string> results = new List<string>();

            if(xlApp != null && xlWorkbook != null)
            {
                try
                {
                    var ranges = xlWorkbook.Names;
                    foreach(var rng in ranges)
                    {
                        string rngName = rng.Name;
                        if(rngName == "" || Regex.IsMatch(rngName, nameLike))
                        {
                            results.Add(rngName);
                        }
                    }
                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                }
            }
            return results;
        }

        /// <summary>
        /// Get an offset range from current range (without affecting current range)
        /// </summary>
        /// <param name="rowOffset">Rows to offset from current range (0 or missing for same as current)</param>
        /// <param name="colOffset">Columns to offset from current range (0 or missing for same as current)</param>
        /// <param name="rows">Rows to return (0 or missing for same as current)</param>
        /// <param name="cols">Columns to return (0 or missing for same as current)</param>
        /// <returns>Excel range, offset from current range</returns>
        public dynamic RangeOffset(int rowOffset = 0, int colOffset = 0, int rows = 0, int cols = 0)
        {
            dynamic result = null;
            if(xlApp != null && xlWorkbook != null)
            {
                try
                {
                    result = xlRange.Offset(rowOffset, colOffset);
                    if(rows != 0 & cols != 0)
                        result = result.Resize(rows, cols);
                    else if(rows != 0)
                        result = result.Resize(rows);
                    else if(cols != 0)
                        result = result.Resize(ColumnCount: cols);
                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                    result = null;
                }
            }
            return result;
        }

        /// <summary>
        /// Resize the current range by supplied rows and columns
        /// </summary>
        /// <param name="rows">Number of rows the range will contain (0 or missing for no change)</param>
        /// <param name="cols">Number of columns the range will contain (0 or missing for no change)</param>
        /// <returns>Boolean indicating success of failure</returns>
        public bool ResizeRange(int rows = 0, int cols = 0)
        {
            bool result = false;
            if(xlApp != null && xlWorkbook != null)
            {
                try
                {
                    if(rows != 0 & cols != 0)
                    {
                        xlRange = xlRange.Resize(rows, cols);
                        result = true;
                    }
                    else if(rows != 0)
                    {
                        xlRange = xlRange.Resize(rows);
                        result = true;
                    }
                    else if(cols != 0)
                    {
                        xlRange = xlRange.Resize(ColumnCount: cols);
                        result = true;
                    }

                }
                catch(Exception ex)
                {
                    LastError = ex.Message;
                    result = false;
                }
            }
            return result;
        }

        /// <summary>
        /// Returns the cell containing the maximum value in the supplied range
        /// </summary>
        /// <param name="rangeToCheck">Range of cells to search</param>
        /// <returns></returns>
        public dynamic GetMaxCellInRange(dynamic rangeToCheck)
        {
            dynamic result = null;
            try
            {
                result = rangeToCheck.Cells(xlApp.WorksheetFunction.Match(xlApp.WorksheetFunction.Max(rangeToCheck), rangeToCheck, 0));
            }
            catch(Exception ex)
            {
                LastError = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// Returns the cell containing the maximum value in the supplied range meeting criteria supplied
        /// </summary>
        /// <param name="rangeToCheck">Range of cells to search</param>
        /// <param name="criteriaRange">Range of cells containing criteria (must be same size as rangeToCheck)</param>
        /// <param name="criteria">Criteria to apply</param>
        /// <returns></returns>
        public dynamic GetMaxCellInRange(dynamic rangeToCheck, dynamic criteriaRange, string criteria)
        {
            dynamic result = null;
            try
            {
                // can't use MAXIFS for older excel compatibility, so iterate range 
                // and save largest cell that meets criteria
                int maxCell = 0;
                float maxVal = 0;
                for(int i = 1; i <= rangeToCheck.Cells.Count; i++)
                {
                    var val = rangeToCheck.Cells[i].Value;
                    string critVal = (criteriaRange.Cells[i].Value ?? 0).ToString();
                    bool isTrue = false;

                    // criteria
                    using(var parser = new DataTable())
                    {
                        try
                        {
                            isTrue = (bool)parser.Compute($"{ critVal.ToString()}{criteria}", string.Empty);
                        }
                        catch(Exception)
                        {
                        }
                    }
                    //var parsingEngine = new DataTable(); //in System.Data
                    //int i = (int)parsingEngine.Compute("3 + 4", String.Empty);
                    //decimal d = (decimal)parsingEngine.Compute("3.45 * 76.9/3", String.Empty);



                    if(isTrue && val > maxVal)
                    {
                        maxCell = i;
                        maxVal = Convert.ToSingle(val);
                    }
                }
                result = rangeToCheck.Cells[maxCell];
            }
            catch(Exception ex)
            {
                LastError = ex.Message;
            }
            return result;
        }

        /// <summary>
        /// release com objects to fully kill excel process from running in the background
        /// </summary>
        private void release_objects()
        {
            //try
            //{
            //    Marshal.ReleaseComObject(xlCell);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlCol);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlRow);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlRange);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //try
            //{
            //    Marshal.ReleaseComObject(xlWorksheet);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            //// quit and release
            //try
            //{
            //    xlApp.Quit();
            //    Marshal.ReleaseComObject(xlApp);
            //}
            //catch(NullReferenceException)
            //{
            //    // ignore if not yet instantiated
            //}

            try
            {
                xlRange = null;
                xlWorksheet = null;
                xlWorkbook = null;
                xlApp.quit();
                xlApp = null;

            }
            catch(Exception)
            {
            }
            finally
            {
                GC.Collect();
            }

        }

        #endregion

    }
}
