using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;


namespace BusinessUnitExcel
{
    public partial class BusinessUnitOrganizerForm : Form
    {


        private Excel.Application o_application = null;
        private Excel.Workbook o_workbook = null;
        private Excel._Worksheet o_first_sheet = null;

        //Get Populated at Runtime
        Dictionary<string, int> dictionary_headers = new Dictionary<string, int>();
        Dictionary<string, BusinessSegment> dictionary_business_segments = new Dictionary<string, BusinessSegment>();
        int num_columns = 0;
        int column_end = 0;
        int num_rows_of_data = 0;

        // Must know header row and start column
        int header_row = 4;
        int column_start = 2;
        int last_row = 4;

        //flags 
        bool config_loaded = false;

        //Totals headers dictionary
        Dictionary<string, SortedSet<string>> dict_total = new Dictionary<string, SortedSet<string>>();

        // formating large boxes max width
        private const int MAX_COLUMN_WIDTH = 30;
        
        /// <summary>
        /// Get LogBox object
        /// </summary>
        public TextBox LogBox
        {
            get { return textbox_log; }
        }

        /// <summary>
        /// Initialize BUO Form
        /// </summary>
        public BusinessUnitOrganizerForm()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Open Excel File using path
        /// </summary>
        /// <param name="path">The abs. path to the Excel File</param>
        private void Setup_Workbook(string path)
        {
            try
            {
                // if excel is already running use that instance
                o_application = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception)
            {
                // open a new instance of excel
                o_application = new Excel.Application();
            }

            // set visibility of application
            o_application.Visible = true;

            // initialize workbook and 'open' file designated by given path
            o_workbook = (Excel.Workbook)(o_application.Workbooks.Open(path));
        }

        /// <summary>
        /// Load event for the form
        /// </summary>
        /// <param name="sender">the form</param>
        /// <param name="e">event args</param>
        private void BusinessUnitOrganizerForm_Load(object sender, EventArgs e)
        {
            //set utility reference for logging 
            Utility.form_ref = this;
            //set defaults
            textbox_column_start.Text = Utility.ConvertNumToColumnLetters(column_start);
            textbox_row_start.Text = header_row.ToString();
            textbox_last_row.Text = last_row.ToString();

            //setup tooltips
            ToolTip t = new ToolTip();
            t.InitialDelay = 100;
            t.ReshowDelay = 100;
            t.SetToolTip(textbox_column_start, Properties.Resources.column_start_tooltip);
            t.SetToolTip(textbox_last_row, Properties.Resources.row_end_tooltip);
            t.SetToolTip(textbox_row_start, Properties.Resources.row_start_tooltip);

            //load config and set flag
            config_loaded = ConfigLoader.LoadConfig();
        }

        /// <summary>
        /// Open File Button Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_open_file_Click(object sender, EventArgs e)
        {
            //Show open file dialog
            openFileDialog.FileName = "";
            DialogResult result = openFileDialog.ShowDialog();
            if (result.Equals(DialogResult.OK))
            {
                Utility.Log("File name:", openFileDialog.FileName);
                //File picked to get here

                //Setup Excel Workbook using Path
                Setup_Workbook(openFileDialog.FileName);
                button_process_data.Enabled = true;
            }


        }

        /**  <summary>Tests if workbook is valid</summary>
         *   <param name="workbook">The Excel Wokbook</param>
         *   <returns>True if workbook is valid False otherwise</returns>
         */
        private bool IsValidWorkbook(Excel.Workbook workbook)
        {
            return workbook != null && workbook.Sheets.Count > 0;
        }

        /// <summary>
        /// Free Memory of Excel Objects
        /// </summary>
        private void FreeComObjects()
        {
            //Clear Up Memory Of COM Objects to be safe
            if (o_application != null)
            {
                Marshal.FinalReleaseComObject(o_application.Workbooks);

                if (o_workbook != null)
                {
                    Marshal.FinalReleaseComObject(o_workbook.Sheets);
                    Marshal.FinalReleaseComObject(o_workbook);
                    o_workbook = null;
                    if (o_first_sheet != null)
                    {
                        Marshal.FinalReleaseComObject(o_first_sheet);
                        o_first_sheet = null;
                    }

                }
            }
        }

        /// <summary>
        /// Form Closed Event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void BusinessUnitOrganizerForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            //FreeComObjects();
            if (o_application != null)
            {
                o_application.Quit();
                Marshal.FinalReleaseComObject(o_application);
                o_application = null;
            }
            //Garbage Collect Anything Necessary
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        /// <summary>
        /// Column Start Textbox Text-changed event;
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textbox_column_start_TextChanged(object sender, EventArgs e)
        {
            TextBox box = (TextBox)sender;
            box.Text = box.Text.ToUpper();
            // Update Value
            if (box.Text == string.Empty)
            {
                column_start = 0;
            }
            else
            {
                // make sure input is valid
                if (!Utility.IsValidColumnLetter(box.Text))
                {
                    box.Text = box.Text.Substring(0, box.Text.Length - 1);
                }
                column_start = Utility.ConvertColumnLetterToNum(box.Text);
            }
            box.Select(box.Text.Length, 0);
            Utility.Log("column starting at:", Utility.ConvertNumToColumnLetters(column_start));
        }

        /// <summary>
        /// row start Textbox Text-changed event;
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textbox_row_start_TextChanged(object sender, EventArgs e)
        {
            TextBox box = (TextBox)sender;
            //Format value with commas
            box.Text = Utility.Format_Int(box.Text);
            box.Select(box.Text.Length, 0);
            //Update value
            if (box.Text == string.Empty)
            {
                header_row = 0;
            }
            else
            {
                header_row = int.Parse(box.Text.Replace(",", ""));
            }
            Utility.Log("header row on:", header_row);
        }

        /// <summary>
        /// last row Textbox Text-changed event;
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void textbox_last_row_TextChanged(object sender, EventArgs e)
        {
            TextBox box = (TextBox)sender;
            //Format number with commas
            box.Text = Utility.Format_Int(box.Text);
            box.Select(box.Text.Length, 0);
            //Update value
            if (box.Text == string.Empty)
            {
                last_row = 0;
            }
            else
            {
                last_row = int.Parse(box.Text.Replace(",", ""));
            }
            Utility.Log("last row at:", last_row);
        }

        /// <summary>
        /// Reads in headers from Excel File
        /// </summary>
        private void LoadHeaders()
        {
            if (IsValidWorkbook(o_workbook))
            {
                // workbook is valid
                // get first sheet (not zero based index so use 1 as first element)
                o_first_sheet = o_workbook.Sheets[1];
                Utility.Log("First Sheet Name:", o_first_sheet.Name);
                o_first_sheet.Activate();

                // Must know header row and start column

                // calculate number of columns of data
                int column_index = column_start;
                string value;
                while ((value = o_first_sheet.Cells[header_row, column_index].Value2) != null)
                {
                    num_columns++;
                    // add header names to list
                    dictionary_headers[value] = column_index;
                    column_index++;
                }
                column_end = column_start + num_columns - 1;
                Utility.Log("Column end for Headers:", Utility.ConvertNumToColumnLetters(column_end));

            }
        }

        /// <summary>
        /// Loads business segment dictionary
        /// </summary>
        private void ReadBusinessSegments()
        {
            // Get Business Segments in the file
            int business_seg_column = ConfigLoader.headerinfo[HeaderConstants.BusinessSegment];

            int row_index = header_row + 1;
            string value = null;
            //get column of business segments
            Excel.Range business_seg = o_first_sheet.get_Range(Utility.GetColumnRange(business_seg_column, header_row + 1, last_row));
            //get data from column
            object[,] data = business_seg.Value2;

            //go through column and pull out any strings 
            for (int i = 0; i < data.Length; i++)
            {
                value = data[i + 1, 1] as string;
                //if no value use '-'
                if (value == null)
                {
                    value = "-";
                }

                //add business segment to dictionary if it doesn't already exist
                BusinessSegment bs;
                if (!dictionary_business_segments.TryGetValue(value, out bs))
                {
                    dictionary_business_segments[value] = new BusinessSegment(value);
                }
                row_index++;
            }

            num_rows_of_data = row_index - header_row - 1;

            Utility.Log("num rows of data: ", num_rows_of_data);

            foreach (string s in dictionary_business_segments.Keys)
            {
                Utility.Log("dictionary business segment key:", s);
            }
        }

        /// <summary>
        /// Writes Individual design review type summary pages
        /// </summary>
        /// <param name="sheet">the sheet to write to</param>
        /// <param name="design_review_type">the design review type (CDR,DDR,...)</param>
        private void WriteDesRevPage(Excel.Worksheet sheet, string design_review_type)
        {
            // get headers from config
            List<string> headers_to_write = ConfigLoader.drtheaders[design_review_type];
            int NUM_INDIVIDUAL_PAGE_HEADERS = headers_to_write.Count;
            object[,] headers = new object[1, NUM_INDIVIDUAL_PAGE_HEADERS];
            // fill headers
            for (int i = 0; i < NUM_INDIVIDUAL_PAGE_HEADERS; i++)
            {
                headers[0, i] = headers_to_write[i];
            }

            // Write headers
            Excel.Range header_row = sheet.get_Range(Utility.GetRowRange(1, 1, NUM_INDIVIDUAL_PAGE_HEADERS));
            header_row.Value2 = headers;
            header_row.WrapText = false;

            long write_row = 2;

            //Write data 
            // foreach business segment
            foreach (BusinessSegment bs in dictionary_business_segments.Values)
            {
                //for each product line in that business segment
                foreach (ProductLine pl in bs.ProductLines)
                {
                    // for each project in that product line
                    foreach (Project proj in pl.Projects)
                    {
                        // create array to hold data
                        object[,] data = new object[1, NUM_INDIVIDUAL_PAGE_HEADERS];
                        // get the range for the row of data
                        Excel.Range data_write = sheet.get_Range(Utility.GetRowRange(write_row, 1, NUM_INDIVIDUAL_PAGE_HEADERS));
                        // get the project data for the specific design review type
                        ProjectData pd = proj[design_review_type];
                        // if that design review type exists
                        if (pd != null)
                        {
                            // for each header entry
                            for (int i = 0; i < headers_to_write.Count; i++)
                            {
                                // get the data associated with that header
                                string key = headers_to_write[i];
                                object val = pd[key];

                                data[0, i] = val;
                            }
                            // set range to have data
                            data_write.Value = data;
                            data_write.WrapText = false;
                            // increment row count
                            write_row++;
                        }

                    }
                }
            }

            sheet.Columns.AutoFit();
            //Shrink Large Columns
            for (int i = 1; i < header_row.Count; i++)
            {
                if (header_row.Item[i].ColumnWidth > MAX_COLUMN_WIDTH)
                {
                    header_row.Item[i].ColumnWidth = MAX_COLUMN_WIDTH;
                }
            }
            sheet.Rows.UseStandardHeight = true;
        }

        /// <summary>
        /// Writes The Summary Page by copying the original data
        /// </summary>
        /// <param name="from">The sheet to copy from</param>
        /// <param name="to">The sheet to write to</param>
        private void CopyOriginalSheet(Excel.Worksheet from, Excel.Worksheet to)
        {
            //Copy header row
            Excel.Range header_f = from.get_Range(Utility.GetRowRange(header_row, column_start, column_end));

            Excel.Range header_t = to.get_Range(Utility.GetRowRange(1, 1, num_columns));

            header_t.Value2 = header_f.Value2;

            // List of design review types
            List<string> design_review_types = ConfigLoader.list_drt;
            // List of extra headers for the summary page
            List<string> extra_summary_headers = ConfigLoader.summary_headers;
            // List of extra headers for each design review type
            List<string> summary_foreach_headers = ConfigLoader.summary_foreach_headers;

            // Calculate size of array of data
            int NUM_EXTRA_SUMMARY_HEADERS = design_review_types.Count * summary_foreach_headers.Count + extra_summary_headers.Count;

            //Add extra headers
            object[,] oaeheader = new object[1, NUM_EXTRA_SUMMARY_HEADERS];
            
            // Fill header with extra summary headers
            for (int i = 0; i < extra_summary_headers.Count; i++)
            {
                oaeheader[0, i] = extra_summary_headers[i];
            }

            // Fill header with extra headers for each design review type
            for (int i = 0; i < design_review_types.Count * summary_foreach_headers.Count; i++)
            {
                oaeheader[0, i + extra_summary_headers.Count] = design_review_types[i / summary_foreach_headers.Count] + " " + summary_foreach_headers[i % summary_foreach_headers.Count];
            }

            // set header data in sheet
            Excel.Range eheader = to.get_Range(Utility.GetRowRange(1, num_columns + 1, num_columns + NUM_EXTRA_SUMMARY_HEADERS));
            eheader.Value2 = oaeheader;
            eheader.WrapText = false;

            //Copy Data
            Excel.Range data_f;
            Excel.Range data_t;
            string old_project_num = "";

            // go through every row in the original sheet that has data
            for (int i = 0, current_row_read = header_row + 1, current_row_write = 2; i < num_rows_of_data; i++, current_row_read++)
            {
                //get the current row on 1st page
                data_f = from.get_Range(Utility.GetRowRange(current_row_read, column_start, column_end));
                object[,] oa_data_f = data_f.Value2;
                //get the project number
                string current_project_num = oa_data_f[1, ConfigLoader.headerinfo[HeaderConstants.ProjectNumber] - column_start + 1] as string;

                //if the project number is different than the previous, create a new entry.
                // this only lets one project number show up
                if (!current_project_num.Equals(old_project_num))
                {
                    // copy original data over
                    data_t = to.get_Range(Utility.GetRowRange(current_row_write, 1, num_columns));
                    data_t.Value = data_f.Value;
                    data_t.WrapText = false;

                    //get business segment
                    string business_segment = Utility.AvoidNull(oa_data_f[1, ConfigLoader.headerinfo[HeaderConstants.BusinessSegment] - column_start + 1] as string);
                    //get product line
                    string product_line = Utility.AvoidNull(oa_data_f[1, ConfigLoader.headerinfo[HeaderConstants.ProductLine] - column_start + 1] as string);

                    // Add Extra data for decisions and dates
                    object[,] oaedata = new object[1, NUM_EXTRA_SUMMARY_HEADERS];

                    //get business segment from map
                    BusinessSegment bs = dictionary_business_segments[business_segment];
                    //get product line from business line
                    ProductLine pl = bs[product_line];
                    //get project from product line by project number
                    Project proj = pl[current_project_num];

                    // fill in extra data for general project
                    for (int k = 0; k < extra_summary_headers.Count; k++)
                    {
                        ProjectData pd = proj[design_review_types[0]];
                        // data will get first design review type
                        if (pd != null)
                        {
                            oaedata[0, k] = pd[extra_summary_headers[k]];
                        }
                        else
                        {
                            oaedata[0, k] = "-";
                        }
                    }

                    //for each design review type, add the extra specified headers
                    for (int k = 0; k < design_review_types.Count * summary_foreach_headers.Count; k++)
                    {
                        ProjectData pd = proj[design_review_types[k / summary_foreach_headers.Count]];
                        if (pd != null)
                        {
                            oaedata[0, k + extra_summary_headers.Count] = pd[summary_foreach_headers[k % summary_foreach_headers.Count]];
                        }
                        else
                        {
                            oaedata[0, k + extra_summary_headers.Count] = "-";
                        }
                    }


                    //write data to sheet
                    string write_range = Utility.GetRowRange(current_row_write, num_columns + 1, num_columns + NUM_EXTRA_SUMMARY_HEADERS);
                    Excel.Range edata = to.get_Range(write_range);
                    edata.Value = oaedata;
                    edata.WrapText = false;

                    old_project_num = current_project_num;
                    current_row_write++;

                }
            }

            //Auto fit columns to appropiate width
            to.Columns.AutoFit();
            //Shrink Large Columns
            for (int i = 1; i < header_t.Count; i++)
            {
                if (header_t.Item[i].ColumnWidth > MAX_COLUMN_WIDTH)
                {
                    header_t.Item[i].ColumnWidth = MAX_COLUMN_WIDTH;
                }
            }
            //Shrink Large Columns
            for (int i = 1; i < eheader.Count; i++)
            {
                if (eheader.Item[i].ColumnWidth > MAX_COLUMN_WIDTH)
                {
                    eheader.Item[i].ColumnWidth = MAX_COLUMN_WIDTH;
                }
            }
            to.Rows.UseStandardHeight = true;
        }

        /// <summary>
        /// Creates extra sheet of summary data
        /// </summary>
        private void CreateSheets()
        {
            // only create sheets if there is only 1 sheet
            if (o_workbook.Sheets.Count < 2)
            {
                int sheet_creation_num = 2;
                Excel._Worksheet o_sheet;

                //Summary (Only one row of each project #)
                o_workbook.Sheets.Add(After: o_first_sheet);
                o_sheet = o_workbook.Sheets[sheet_creation_num++];
                o_sheet.Activate();
                o_sheet.Name = Properties.Resources.sheet_name_summary;
                CopyOriginalSheet((Excel.Worksheet)o_first_sheet, (Excel.Worksheet)o_sheet);

                // create a sheet for each design review type
                List<string> design_review_types = ConfigLoader.list_drt;

                foreach (string drt in design_review_types)
                {
                    o_workbook.Sheets.Add(After: o_sheet);

                    Marshal.ReleaseComObject(o_sheet);
                    o_sheet = null;

                    o_sheet = o_workbook.Sheets[sheet_creation_num++];
                    o_sheet.Activate();
                    o_sheet.Name = drt;
                    WriteDesRevPage((Excel.Worksheet)o_sheet, drt);
                }

                FillTotalsKeys();

                ////////////Totals
                o_workbook.Sheets.Add(After: o_sheet);
                Marshal.ReleaseComObject(o_sheet);
                o_sheet = null;
                o_sheet = o_workbook.Sheets[sheet_creation_num++];
                o_sheet.Activate();
                o_sheet.Name = Properties.Resources.sheet_name_totals;
                WriteTotals((Excel.Worksheet)o_sheet);
                ////////////End Totals
            }

        }


        /// <summary>
        /// Goes through header columns and calculates all values present in that column
        /// </summary>
        /// <param name="headers">The list of headers</param>
        private void FillTotalsDictionary(List<string> headers)
        {
            // go through all headers
            foreach (string header in headers)
            {
                //get column that the header is in
                int column = -1;
                dictionary_headers.TryGetValue(header, out column);
                if (column != -1)
                {
                    // create set to store values
                    SortedSet<string> set = new SortedSet<string>();
                    dict_total.Add(header, set);
                    //get entire column
                    object[,] column_data = o_first_sheet.get_Range(Utility.GetColumnRange(column, header_row + 1, header_row + num_rows_of_data)).Value2;
                    //go through the entire column
                    for (int i = 0; i < column_data.Length; i++)
                    {
                        //add value to set
                        string value = Utility.AvoidNull(column_data[i + 1, 1] as string);
                        set.Add(value);
                    }
                }else
                {
                    // header not good
                    Utility.Log("Error", string.Format("Header '{0}' not valid: look for typos", header));
                }
            }
        }

        /// <summary>
        /// Fills totals headers to eventually calculate
        /// </summary>
        private void FillTotalsKeys()
        {
            List<string> totals_headers = ConfigLoader.totals_headers;
            List<string> totals_foreach_headers = ConfigLoader.totals_foreach_headers;

            // could crash if header is not in dictionary header
            FillTotalsDictionary(totals_headers);
            FillTotalsDictionary(totals_foreach_headers);

        }

        /// <summary>
        /// Writes totals page
        /// </summary>
        /// <param name="sheet">the sheet to write the totals on</param>
        private void WriteTotals(Excel.Worksheet sheet)
        {
            List<string> totals_headers = ConfigLoader.totals_headers;
            List<string> totals_foreach_headers = ConfigLoader.totals_foreach_headers;
            List<string> list_drt = ConfigLoader.list_drt;

            int NUM_TOTAL_PAGE_HEADERS = 0;
            //Calculate total headers 
            if (totals_foreach_headers.Count > 0)
            {
                //includes foreach headers if present
                NUM_TOTAL_PAGE_HEADERS = list_drt.Count;
                for (int j = 0; j < totals_foreach_headers.Count; j++)
                {
                    NUM_TOTAL_PAGE_HEADERS += (dict_total[totals_foreach_headers[j]].Count * list_drt.Count);
                }
                for (int j = 0; j < totals_headers.Count; j++)
                {
                    NUM_TOTAL_PAGE_HEADERS += dict_total[totals_headers[j]].Count;
                }
            }
            else
            {
                // no foreach headers included
                for (int j = 0; j < totals_headers.Count; j++)
                {
                    NUM_TOTAL_PAGE_HEADERS += dict_total[totals_headers[j]].Count;
                }
            }

            object[,] feheaders = new object[1, NUM_TOTAL_PAGE_HEADERS];

            //Write Business Segments Vertically
            object[,] headers = new object[dictionary_business_segments.Keys.Count, 1];
            sheet.Cells[1, 1] = HeaderConstants.BusinessSegment;
            int i = 0;
            foreach (string business_segment in dictionary_business_segments.Keys)
            {
                headers[i++, 0] = business_segment;
            }
            string range = Utility.GetColumnRange(1, 2, i + 1);
            Excel.Range rbusiness_seg = sheet.get_Range(range);
            rbusiness_seg.Value2 = headers;

            //Write ForEach Data
            int col_start = Utility.ConvertColumnLetterToNum("B");

            if (totals_foreach_headers.Count > 0)
            {
                // write each design review type total
                foreach (string design_review_type in list_drt)
                {
                    //write and calculate total columns written
                    int columns_written = WriteDesignTotals(sheet, design_review_type, col_start, totals_foreach_headers, feheaders, headers);
                    col_start += columns_written;
                }
            }

            //Write other data
            WriteAllTotal(sheet, col_start, totals_headers, feheaders, headers);

            //Autofit columns
            sheet.Columns.AutoFit();
        }

        /// <summary>
        /// Writes design review totals
        /// </summary>
        /// <param name="sheet">The sheet to write to</param>
        /// <param name="design_review_type">the design review type</param>
        /// <param name="col_start">the column to start on</param>
        /// <param name="total_foreach_headers">the list of headers</param>
        /// <param name="feheaders">the data array for headers</param>
        /// <param name="vert_business_seg">the array of business segments</param>
        /// <returns></returns>
        private int WriteDesignTotals(Excel.Worksheet sheet, string design_review_type, int col_start, List<string> total_foreach_headers, object[,] feheaders, object[,] vert_business_seg)
        {
            int NUM_TOTAL_PAGE_HEADERS = 0;
            // fill in headers
            feheaders[0, col_start - 2] = Properties.Resources.string_design_review_type;
            int feheader_index = col_start - 1;
            for (int i = 0; i < total_foreach_headers.Count; i++)
            {
                SortedSet<string> set = dict_total[total_foreach_headers[i]];
                foreach (string h in set)
                {
                    feheaders[0, feheader_index++] = total_foreach_headers[i] + " " + h;
                    NUM_TOTAL_PAGE_HEADERS++;
                }
            }

            //write headers to sheet
            string range_text = Utility.GetRowRange(1, col_start, col_start + NUM_TOTAL_PAGE_HEADERS);
            //Utility.Log("Total Range:", range_text);
            Excel.Range rheader = sheet.get_Range(range_text);
            rheader.Value2 = feheaders;
            rheader.WrapText = false;

            int row_start = 2;
            //Go through all business segments
            for (int i = 0; i < vert_business_seg.GetLength(0); i++)
            {
                //get business segment string
                string business_seg = vert_business_seg[i, 0] as string;
                //create array to store data
                object[,] oadata = new object[1, NUM_TOTAL_PAGE_HEADERS + 1];
                //set column design review type
                oadata[0, 0] = design_review_type;

                int col_index = 1;
                // calculate totals for each header key
                for (int j = 0; j < total_foreach_headers.Count; j++)
                {
                    string key = total_foreach_headers[j];
                    SortedSet<string> set = dict_total[total_foreach_headers[j]];
                    // for each header value calculate totals
                    foreach (string value in set)
                    {
                        BusinessSegment bs = dictionary_business_segments[business_seg];
                        oadata[0, col_index++] = bs.CalculateTotal(design_review_type, key, value);
                    }
                }

                // set data
                Excel.Range r_total = sheet.get_Range(Utility.GetRowRange(row_start + i, col_start, col_start + NUM_TOTAL_PAGE_HEADERS));
                r_total.Value2 = oadata;
            }
            //return columns written
            return NUM_TOTAL_PAGE_HEADERS + 1;
        }
        
        /// <summary>
        /// Writes all generic totals information
        /// </summary>
        /// <param name="sheet">the sheet to write to</param>
        /// <param name="col_start">the column to start on</param>
        /// <param name="totals_headers">the list of headers</param>
        /// <param name="feheaders">the array that stores the headers</param>
        /// <param name="vert_business_seg">the business segment array</param>
        /// <returns></returns>
        private int WriteAllTotal(Excel.Worksheet sheet, int col_start, List<string> totals_headers, object[,] feheaders, object[,] vert_business_seg)
        {
            int NUM_TOTAL_HEADERS = 0;
            List<string> list_drt = ConfigLoader.list_drt;

            //fill in header info
            int feheader_index = col_start - 2;
            for (int i = 0; i < totals_headers.Count; i++)
            {
                SortedSet<string> set = dict_total[totals_headers[i]];
                foreach (string h in set)
                {
                    feheaders[0, feheader_index++] = totals_headers[i] + " " + h;
                    NUM_TOTAL_HEADERS++;
                }
            }

            //correct offset by copying header into new array
            object[,] feheaders_actual = new object[1, NUM_TOTAL_HEADERS];
            for (int i = 0; i < NUM_TOTAL_HEADERS; i++)
            {
                feheaders_actual[0, i] = feheaders[0, i + feheader_index - NUM_TOTAL_HEADERS];
            }

            //set headers in sheet
            string range_text = Utility.GetRowRange(1, col_start, col_start + NUM_TOTAL_HEADERS - 1);
            Excel.Range rheader = sheet.get_Range(range_text);
            rheader.Value2 = feheaders_actual;
            rheader.WrapText = false;

            int row_start = 2;
            //for each business segment
            for (int i = 0; i < vert_business_seg.GetLength(0); i++)
            {
                string business_seg = vert_business_seg[i, 0] as string;
                object[,] oadata = new object[1, NUM_TOTAL_HEADERS];

                //for each header
                int col_index = 0;
                for (int j = 0; j < totals_headers.Count; j++)
                {
                    //for each value associated with that header
                    string key = totals_headers[j];
                    SortedSet<string> set = dict_total[totals_headers[j]];
                    foreach (string value in set)
                    {
                        BusinessSegment bs = dictionary_business_segments[business_seg];
                        int sum = 0;
                        // include all design review types in this total
                        foreach (string drt in list_drt)
                        {
                            sum += bs.CalculateTotal(drt, key, value);
                        }
                        oadata[0, col_index++] = sum;
                    }
                }
                //set data in sheet
                Excel.Range r_total = sheet.get_Range(Utility.GetRowRange(row_start + i, col_start, col_start + NUM_TOTAL_HEADERS - 1));
                r_total.Value2 = oadata;
            }
            // return columns written
            return NUM_TOTAL_HEADERS;
        }

        /// <summary>
        /// Fill in data from sheet into mempry
        /// </summary>
        private void ReadRows()
        {
            int offset = 1;

            // read headers
            object[,] headers_2d = o_first_sheet.get_Range(Utility.GetRowRange(header_row, column_start, column_end)).Value2;

            string[] headers = new string[headers_2d.GetLength(1)];
            for (int i = 0; i < headers_2d.GetLength(1); i++)
            {
                headers[i] = headers_2d[1, i + 1] as string;
            }

            object[,] data = o_first_sheet.get_Range(Utility.GetRowRange(header_row + (offset++), column_start, num_columns + 1)).Value;
            //get data necessary for keys in dictionaries
            string business_segment_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.BusinessSegment] - column_start + 1] as string);
            string product_line_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProductLine] - column_start + 1] as string);
            string project_number = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProjectNumber] - column_start + 1] as string);
            string project_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProjectName] - column_start + 1] as string);
            string design_review_type = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.DesignReviewType] - column_start + 1] as string);

            // while projects still left to read
            while (project_name != "-")
            {
                //get business segment
                BusinessSegment bs = dictionary_business_segments[business_segment_name];
                // add product line
                bs.AddProductLine(product_line_name);
                //get product line
                ProductLine pl = bs[product_line_name];
                //add project
                pl.AddProject(project_name, project_number);
                // get project
                Project proj = pl[project_number];
                //add project data
                proj.AddProjectData(design_review_type);
                //get project data
                ProjectData pd = proj[design_review_type];

                //set project data key value pairs
                for (int i = 0; i < headers.Length; i++)
                {
                    pd[headers[i] as string] = data[1, i + 1];
                }

                //get data again
                data = o_first_sheet.get_Range(Utility.GetRowRange(header_row + (offset++), column_start, num_columns + 1)).Value;

                business_segment_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.BusinessSegment] - column_start + 1] as string);
                product_line_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProductLine] - column_start + 1] as string);
                project_number = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProjectNumber] - column_start + 1] as string);
                project_name = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.ProjectName] - column_start + 1] as string);
                design_review_type = Utility.AvoidNull(data[1, ConfigLoader.headerinfo[HeaderConstants.DesignReviewType] - column_start + 1] as string);
                
            }
            
        }

        /// <summary>
        /// Process Excel File
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button_process_data_Click(object sender, EventArgs e)
        {
            // Verify config loaded properly
            if (config_loaded)
            {
                //Verify last row is set and makes sense
                if (last_row > header_row)
                {
                    //if workbook actually exists
                    if (o_workbook != null)
                    {
                        //used to calclulate time elapsed
                        DateTime before = DateTime.Now;
                        //Process data
                        LoadHeaders();
                        ReadBusinessSegments();
                        ReadRows();
                        CreateSheets();
                        //calculate elapsed time
                        DateTime after = DateTime.Now;
                        Utility.Log("Done Writing. Took", (after - before).ToString("T") + " to complete");
                        button_process_data.Enabled = false;
                        // free objects to ensure no memory leaks
                        FreeComObjects();
                    }
                    else
                    {
                        Utility.Log("Error:", "An Excel Document must be opened first");
                    }
                }
                else
                {
                    Utility.Log("Error:", "Last-Row number must be > Header-Row");
                }
            }
            else
            {
                Utility.Log("Error:", "Config.xml not found, restart program with Config file in same directory");
            }
        }

    }
}
