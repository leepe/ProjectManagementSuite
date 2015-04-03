using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WriteOutLongTermProjectFcsts
{
    public class MainProgramLogic
    {
        public class ProjectFcst
        {
            public int ProjID { get; set; }
            public string ProjNum { get; set; }
            public string ProjName { get; set; }
            public string ProjDesc { get; set; }
            public string State { get; set; }
            public string ProjManFlag { get; set; }
            public List<string> MVXProjNum { get; set; }
        }

        public class projectFcstLines
        {
            public int ProjId { get; set; }
            public string ProjItem { get; set; }
            public string Whse { get; set; }
            public int Month { get; set; }
            public decimal Qty { get; set; }
        }
        //
        public static void Main(string[] args)
        {
            // clean out all folders that will contain workbooks
            Dictionary<string, string> dfp = loadStatePathDictionary();
            foreach (var k in dfp)
            {
                Array.ForEach(Directory.GetFiles(k.Value), File.Delete);
            }
            // connect to sql express - get project header details
            //SqlConnection cnxn = setUpEXPRconn();
            SqlConnection cnxn = setUpDSXconn();
            DataTable dtHead = GetProjectHeaders(cnxn);
            // load up reference dictionaries
            Dictionary<string, string> dcn = getItemMaster(cnxn);
            Dictionary<string, string> diw = getItemWhseMaster(cnxn);
            Dictionary<string, string> dbu = getBUMaster(cnxn);
            //
            IEnumerable<ProjectFcst> pset = from g in dtHead.AsEnumerable()
                                            select new ProjectFcst
                                            {
                                                ProjID = g.Field<int>("ProjId"),
                                                ProjNum = g.Field<string>("ProjNum").Trim(),
                                                ProjName = g.Field<string>("ProjName").Trim(),
                                                ProjDesc = string.IsNullOrWhiteSpace(g.Field<string>("ProjDesc")) ? string.Empty :
                                                           g.Field<string>("ProjDesc").Trim(),
                                                State = g.Field<string>("State").Trim(),
                                                ProjManFlag = g.Field<string>("ProjManFlag").Trim(),
                                                MVXProjNum = string.IsNullOrEmpty(g.Field<string>("MVXProjNum")) ? new List<string>() :
                                                             Enumerable.Range(0, g.Field<string>("MVXProjNum").Trim().Length/10)
                                                             .Select(i => g.Field<string>("MVXProjNum").Trim().Substring(i * 10, 10)).ToList()
                                            };
            // 
            // now start to loop through the ienumerable
            foreach (ProjectFcst g in pset)
            {

                string tabName = g.ProjName.Replace(" ", string.Empty).Trim().Replace(",", string.Empty);          // compress name
                string fn0 = string.Format("{0}_{1}_{2}.xls", DateTime.Now.ToString("yyyyMMdd")                    // now date
                                            , g.ProjID.ToString("D0").Trim()                                       // project number
                                            , tabName.Substring(0, Math.Min(26, tabName.Length)));                 // compressed name
                string fpath = dfp[g.State.Trim()];
                generateProjectWorkBook(g, fpath+fn0);
                DataTable dtLines = GetProjectLines(cnxn, g.ProjID);
                // now add in project lines
                addProjectWorkBookLines(cnxn, dcn,diw,dbu,dtLines, fpath + fn0);
            }

            //Console.ReadKey();
        }

        //------------------------------------------------------------------------
        // retrieve project line data from express
        // modified to pass sales state through to datatable
        //------------------------------------------------------------------------

        private static DataTable GetProjectLines(SqlConnection conn, int ProjNo)
        {
            string sql = "select a.ProjID,a.ProjItem,a.Whse,b.[State],a.[Month],a.Qty " +
                         "from dbo.[ProjectItems] a inner join dbo.[Projects] b on (a.ProjID=b.ProjID) " +
                         "where a.ProjID=" + ProjNo;
            //string sql = "select distinct a.ProjID,b.ProjItem,b.Whse,a.[State],b.[Month],b.Qty from dbo.[Projects] a inner join dbo.[ProjectItems] b " +
            //             "on (a.ProjID=b.ProjID) where a.ProjManFlag in ('Yes') and a.ProjID=" + ProjNo +
            //             "union " +
            //             "select distinct a.ProjID,b.ProjItem,b.Whse,a.[State],b.[Month],b.[Qty] from dbo.[Projects] a inner join dbo.[ProjectMVXRemaining] b " +
            //             "on (a.ProjID=b.ProjID) where a.ProjManFlag in ('No') and a.ProjID=" + ProjNo;
            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(sql, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }

        //------------------------------------------------------------------------
        // load dictionary of file paths 
        //------------------------------------------------------------------------
        private static Dictionary<string, string> loadStatePathDictionary()
        {
            Dictionary<string, string> dcn = new Dictionary<string, string>();
            // add file paths to dictionary for each state
            dcn.Add("NSW", @"C:\ProjectData\ProjectFiles\LONGTERM\NSW\");
            dcn.Add("VIC", @"C:\ProjectData\ProjectFiles\LONGTERM\VIC\");
            dcn.Add("QLD", @"C:\ProjectData\ProjectFiles\LONGTERM\QLD\");
            dcn.Add("SA", @"C:\ProjectData\ProjectFiles\LONGTERM\SA\");
            dcn.Add("WA", @"C:\ProjectData\ProjectFiles\LONGTERM\WA\");
            dcn.Add("NZ", @"C:\ProjectData\ProjectFiles\LONGTERM\NZ\");
            //
            return dcn;
        }

        //------------------------------------------------------------------------
        // retrieve project header data from express
        //------------------------------------------------------------------------

        private static DataTable GetProjectHeaders(SqlConnection conn)
        {
            //string sql = "select ProjID,ProjNum,ProjName,ProjDesc,State,ProjManFlag," +
            //             "MVXProjNum from dbo.[Projects] where ProjManFlag='Yes'";
            string sql = "select projid,projnum,projname,projdesc,state,projmanflag," +
                         "mvxprojnum from dbo.[projects]";
            //string sql = "with unPrj as " +
            //             "( " +
            //             "select distinct a.ProjID,sum(b.[Qty]) Qty from dbo.[Projects] a inner join dbo.[ProjectItems] b " +
            //             "on (a.ProjID=b.ProjID) where a.ProjManFlag in ('Yes') " +
            //             "group by a.ProjID " +
            //             "union " +
            //             "select distinct a.ProjID,SUM(b.[Qty]) Qty from dbo.[Projects] a inner join dbo.[ProjectMVXRemaining] b " +
            //             "on (a.ProjID=b.ProjID) where a.ProjManFlag in ('No') " +
            //             "group by a.ProjID " +
            //             ") " +
            //             "select distinct a.ProjID,a.ProjNum,a.ProjName,a.ProjDesc,a.State,a.ProjManFlag,a.MVXProjNum  " +
            //             "from dbo.[Projects] a inner join dbo.[ProjectItems] b on (a.ProjID=b.ProjID) " +
            //             "inner join unPrj on (a.ProjID=unPrj.ProjID)";
            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(sql, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }
        //-------------------------------------------------------------------
        // open SQL client connection
        //-------------------------------------------------------------------
        private static SqlConnection setUpDSXconn()
        {
            // create connection object
            string strSQLsrvr = "SERVER=WETNT260;USER ID=#DSXdbadmin;PASSWORD=F0res!3R;" +
                    "DATABASE=STAGING;CONNECTION TIMEOUT=30;";
            SqlConnection SqlConn = new SqlConnection(strSQLsrvr);
            return SqlConn;
        }
        //-------------------------------------------------------------------
        // set up connection to SQL-EXPRESS - thinkpad
        //-------------------------------------------------------------------
        private static SqlConnection setUpEXPRconn()
        {
            // create connection object
            string strSQLsrvr = @"Data Source=.\SQLEXPRESS;Initial Catalog=STAGING;Integrated Security=True";
            SqlConnection SqlConn = new SqlConnection(strSQLsrvr);
            return SqlConn;
        }
        //--------------------------------------------------------------
        // open project workbook and write out current project lines
        //--------------------------------------------------------------
        public static void addProjectWorkBookLines(SqlConnection cn,Dictionary<string,string> d1,
                                                                    Dictionary<string,string> d2,
                                                                    Dictionary<string,string> d3, DataTable dt, string f0)
        {
            // open file back up and write out lines then 
            using (FileStream fprojects = new FileStream(f0, FileMode.Open, FileAccess.Read))
            {
                // set up NPOI objects
                IWorkbook xWrite = new HSSFWorkbook(fprojects);
                Dictionary<string, ICellStyle> sd = createStyles(xWrite);    // load dictionary of excel styles
                HSSFSheet xSheet1 = (HSSFSheet)xWrite.GetSheetAt(0);         // main excel project sheet
                // set up highlight font
                IFont iWarn = xWrite.CreateFont();
                iWarn.Color = HSSFColor.Red.Index;
                iWarn.Boldweight = (short)FontBoldWeight.Bold;
                //
                // loop through excel sheet pick up the dates on the sheet
                string[] dte = new string[23];
                for (int z = 0; z < 23; z++) dte[z] = xSheet1.GetRow(8).GetCell(6 + z).DateCellValue.ToString("yyyyMMdd");
                // run out data onto spreadsheet
                var numitem = (from h in dt.AsEnumerable() 
                               select new { 
                                   Item = h.Field<string>("ProjItem").Trim(), 
                                   Whse = h.Field<string>("Whse").Trim() 
                               }).Distinct();             // number of distinct items ie number of lines to write
                int ncount = 0;
                foreach (var j in numitem)
                {
                    var subqry = (from w in dt.AsEnumerable() select w).Where(z => z.Field<string>("ProjItem").Trim().Equals(j.Item));
                    string stID = (from x in dt.AsEnumerable() select x.Field<string>("State").Trim()).Last();
                    xSheet1.GetRow(ncount + 9).GetCell(1).SetCellValue(j.Item.Trim());
                    xSheet1.GetRow(ncount + 9).GetCell(3).SetCellValue(j.Whse.Trim());
                    //...............................................................................................................
                    // add in item description  - first dictionary D C N 
                    //...............................................................................................................
                    xSheet1.GetRow(ncount + 9).CreateCell(2);                                         // need to create cell first
                    string kval = string.Empty;
                    if (d1.TryGetValue(j.Item.Trim(), out kval))                                      // protect dictionary
                    {
                        xSheet1.GetRow(ncount+9).GetCell(2).SetCellValue(kval);                       // if the item exists in dictionary
                    }
                    else
                    {
                        xSheet1.GetRow(ncount+9).GetCell(2).SetCellValue("ITEM DOES NOT EXIST");      // if the item doesn't exist
                        xSheet1.GetRow(ncount+9).GetCell(2).RichStringCellValue.ApplyFont(iWarn);     // highlight in red
                    }
                    //...............................................................................................................
                    // add in item warehouse data  - second dictionary D I W 
                    //...............................................................................................................
                    xSheet1.GetRow(ncount + 9).CreateCell(4);                                           // create new cell first
                    string iwval = string.Empty;
                    if (d2.TryGetValue(j.Item.Trim() + j.Whse.Trim() + stID, out iwval))                // testing itme X whse X state combination
                    {
                        xSheet1.GetRow(ncount+9).GetCell(4).SetCellValue(iwval);
                    }
                    else
                    {
                        xSheet1.GetRow(ncount+9).GetCell(4).SetCellValue("COMBINATION DOES NOT EXIST");   // if the Item-Warehouse doesn't exist
                        xSheet1.GetRow(ncount+9).GetCell(4).RichStringCellValue.ApplyFont(iWarn);
                    }
                    //...............................................................................................................
                    // add in business  - third dictionary D B U 
                    //...............................................................................................................
                    xSheet1.GetRow(ncount+9).CreateCell(5);
                    string ibval = string.Empty;
                    if (d3.TryGetValue(j.Item.Trim(), out ibval))
                    {
                        xSheet1.GetRow(ncount+9).GetCell(5).SetCellValue(ibval);
                    }
                    else
                    {
                        xSheet1.GetRow(ncount+9).GetCell(5).SetCellValue("BU DOESNT EXIST");                  // if the BUShort doesn't exist
                        xSheet1.GetRow(ncount+9).GetCell(5).RichStringCellValue.ApplyFont(iWarn);
                    }
                    //................................................................................................................
                    // now write out the monthly data
                    //................................................................................................................
                    foreach (var s in subqry)
                    {
                        if (Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()) > -1)  // if the Month is found in the date array
                        {
                            //// write out the monthly data
                            xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim())).SetCellType(CellType.Numeric);
                            xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()))
                                                      .SetCellValue(Double.Parse(s.Field<decimal>("Qty").ToString("F0")));
                            // write out the monthly data
                            if (Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()) < 6)
                            {
                                xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim())).CellStyle = sd["Late_Qty"];
                            }
                            else
                            {
                                xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim())).CellStyle = sd["OK_Qty"];
                            }

                        }
                    }
                    ncount += 1;
                }
                // create new file and write out xRead
                FileStream f1 = new FileStream(f0, FileMode.Create);
                xWrite.Write(f1);
                f1.Close();

            }
        }
        //--------------------------------------------------------------
        // create project workbook to standard format - passed as nph
        //--------------------------------------------------------------
        public static void generateProjectWorkBook(ProjectFcst nph, string f0)
        {

            String[] titles = { "Item", "Description", "Whse", "Item Status", "Business", "End" };
            String[] verticalTitles = { "Project Number", "Project Name", "Description", "Manual Fcst", "Movex CO Number" };

            // location of template file with VBA module
            string ftempl = @"C:\ProjectData\ProjectFiles\LONGTERM\TEMPLATE\ManualFcstLogic2.xls";
            //--------------------------------------------------------
            // create workbook and define sheets
            //--------------------------------------------------------
            using (FileStream ftemplate = new FileStream(ftempl, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(ftemplate);
                //---------------------------------------------------------
                // get dictionaries
                //---------------------------------------------------------
                Dictionary<string, ICellStyle> sd = createStyles(workbook);
                //
                //---------------------------------------------------------
                // produce a safe project sheet name
                //---------------------------------------------------------
                string tabName = Regex.Replace(nph.ProjName.Substring(0, Math.Min(25, nph.ProjName.Length)).Trim(), @"[^a-zA-Z0-9 -]", "");
                ISheet sheet = workbook.GetSheetAt(0);
                workbook.SetSheetName(0, "Proj~" + tabName);
                //---------------------------------------------------------
                // turn on sheet print set up specific things
                //---------------------------------------------------------
                sheet.DisplayGridlines = true;
                sheet.IsPrintGridlines = true;
                sheet.HorizontallyCenter = true;
                sheet.PrintSetup.Landscape = true;
                sheet.PrintSetup.PaperSize = (short)PaperSize.A4;
                sheet.Autobreaks = true;
                sheet.PrintSetup.FitHeight = (short)1;
                sheet.PrintSetup.FitWidth = (short)1;
                //---------------------------------------------------------
                // set up header row
                //---------------------------------------------------------
                IRow headerRow = sheet.CreateRow(8); // this is zero based row numbering
                headerRow.HeightInPoints = 19.00f;
                for (int i = 0; i < titles.Length; i++)
                {
                    ICell cell = headerRow.CreateCell(i + 1);
                    cell.SetCellValue(titles[i]);
                    cell.CellStyle = sd["cell_normal"];
                }
                //---------------------------------------------------------
                // columns for 24 months starting from today
                //---------------------------------------------------------
                //DateTime d0 = DateTime.Now;
                //DateTime startMonth = new DateTime(d0.Year, d0.Month, 1);
                //for (int i = 0; i < 23; i++)
                //{
                //    ICell cell = headerRow.CreateCell(titles.Length + i);
                //    cell.SetCellValue(startMonth.AddMonths(i));
                //    cell.CellStyle = sd["cell_normal_date"];
                //}
                //---------------------------------------------------------
                // roll out months back 6 and forward 18
                //---------------------------------------------------------
                DateTime d0 = DateTime.Now.AddMonths(-6);
                DateTime c0 = DateTime.Now;
                DateTime curr = new DateTime(c0.Year, c0.Month, 1);
                DateTime startMonth = new DateTime(d0.Year, d0.Month, 1);
                for (int i = 0; i < 23; i++)
                {
                    ICell cell = headerRow.CreateCell(titles.Length + i);
                    cell.SetCellValue(startMonth.AddMonths(i));
                    if (startMonth.AddMonths(i) < curr)
                    {
                        cell.CellStyle = sd["cell_late_date"];
                    }
                    else
                    {
                        cell.CellStyle = sd["cell_ok_date"];
                    }
                }
                //----------------------------------------------------------
                //freeze the first row
                //----------------------------------------------------------
                sheet.CreateFreezePane(6, 9);
                // output header rows
                IRow row = sheet.CreateRow(0);
                ICell cell0 = row.CreateCell(0);
                cell0.SetCellValue("Project Data");
                cell0.CellStyle = sd["cell_normal"];
                ICell cell1 = row.CreateCell(1);
                cell1.SetCellValue("Project ID");
                cell1.CellStyle = sd["cell_normal"];
                ICell cell2 = row.CreateCell(2);
                cell2.SetCellValue(nph.ProjID);
                cell2.CellStyle = sd["cell_green"];
                // roll down and output column values
                for (int i = 0; i < verticalTitles.Length; i++)
                {
                    IRow vtitleCol = sheet.CreateRow(i + 1);
                    ICell cell = vtitleCol.CreateCell(1);
                    ICell cell_1 = vtitleCol.CreateCell(2);
                    // check for nulls
                    string tmpName = string.IsNullOrEmpty(nph.ProjName) ?
                                    string.Empty : Regex.Replace(nph.ProjName.Trim(), @"[^a-zA-Z0-9 -]", "");
                    string tmpDesc = string.IsNullOrEmpty(nph.ProjDesc) ?
                                    string.Empty : Regex.Replace(nph.ProjDesc.Trim(), @"[^a-zA-Z0-9 -]", "");
                    //
                    if (i == 0) cell_1.SetCellValue(nph.ProjNum.Trim());
                    else if (i == 1) cell_1.SetCellValue(tmpName);
                    else if (i == 2) cell_1.SetCellValue(tmpDesc);
                    else if (i == 3 && nph.MVXProjNum.Count > 0) cell_1.SetCellValue("No");
                    else if (i == 3 && nph.MVXProjNum.Count == 0) cell_1.SetCellValue("Yes");
                    else if (i == 4)
                    {
                        if (nph.MVXProjNum.Count > 0)
                        {
                            StringBuilder sb = new StringBuilder();
                            for (int j = 0; j < nph.MVXProjNum.Count; j++)
                            {
                                sb.Append(nph.MVXProjNum[j]);
                                sb.Append("\n");
                            }
                            cell_1.SetCellValue(sb.ToString());
                        }
                    }
                    cell.SetCellValue(verticalTitles[i]);
                    cell.CellStyle = sd["cell_normal"];
                    if (i != 3) cell_1.CellStyle = sd["cell_green"];
                    cell_1.CellStyle.Alignment = HorizontalAlignment.Left;
                    cell_1.CellStyle = sd["cell_green"];
                    cell_1.CellStyle.Alignment = HorizontalAlignment.Left;
                }
                // roll down the rest of the sheet and format
                for (int i = 9; i < 100; i++)
                {
                    IRow formatRow = sheet.CreateRow(i);
                    for (int j = 1; j < 29; j++)
                    {
                        ICell cell = formatRow.CreateCell(j);
                        if (j == 1 || j == 3 || (j >= 6 && j <= 28))
                        {
                            cell.CellStyle = sd["cell_green"];
                            cell.CellStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("General");
                        }
                    }
                }
                //set column widths, the width is measured in units of 1/256th of a character width
                sheet.SetColumnWidth(0, 256 * 16);
                sheet.SetColumnWidth(1, 256 * 25);
                sheet.SetColumnWidth(2, 256 * 60);
                sheet.SetColumnWidth(3, 256 * 9);
                sheet.SetColumnWidth(4, 256 * 30);
                sheet.SetColumnWidth(5, 256 * 10);
                sheet.SetZoom(5, 6);
                //------------------------------------------------------------------------
                // now add a second sheet called MasterControlSheet with person's state
                //------------------------------------------------------------------------
                ISheet sheet2 = workbook.GetSheetAt(1);
                IRow topRow = sheet2.CreateRow(0);
                ICell cell_f = topRow.CreateCell(0);
                cell_f.SetCellValue("State");
                ICell cell_g = topRow.CreateCell(1);
                cell_g.SetCellValue(nph.State);
                // -------- THE END -------------------
                FileStream sw = File.Create(f0);
                workbook.Write(sw);
                sw.Close();
            }
        }

        //--------------------------------------------------------------
        // create dictionary of styles to be used in xlsx creation
        //--------------------------------------------------------------
        private static Dictionary<string, ICellStyle> createStyles(IWorkbook wb)
        {
            // create dictionary
            Dictionary<string, ICellStyle> styleDictionary = new Dictionary<string, ICellStyle>();
            // fonts
            IFont dfont = wb.CreateFont();
            dfont.Boldweight = (short)FontBoldWeight.Bold;
            dfont.FontName = "Calibri";
            dfont.IsItalic = true;
            dfont.Color = HSSFColor.RoyalBlue.Index;
            dfont.FontHeightInPoints = 12;
            // red warning fonts
            IFont wfont = wb.CreateFont();
            wfont.Boldweight = (short)FontBoldWeight.Bold;
            wfont.IsItalic = true;
            wfont.FontName = "Calibri";
            wfont.Color = HSSFColor.Maroon.Index;
            wfont.FontHeightInPoints = 12;
            // define styles to apply to output data
            ICellStyle icstyle1 = wb.CreateCellStyle();
            //
            //---- Cell_Normal ------
            //
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.WrapText = true;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            styleDictionary.Add("cell_normal", icstyle1);
            //
            //---- Cell_Normal_Date ------
            //
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.WrapText = true;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");

            styleDictionary.Add("cell_normal_date", icstyle1);
            //
            //---- Cell_Green ------------
            icstyle1 = createBorderedStyle(wb);
            icstyle1.Alignment = HorizontalAlignment.Left;
            icstyle1.FillForegroundColor = HSSFColor.LightGreen.Index;
            icstyle1.FillPattern = FillPattern.SolidForeground;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            icstyle1.WrapText = true;
            styleDictionary.Add("cell_green", icstyle1);
            //
            //---- Header Style -----
            //
            IFont headerFont = wb.CreateFont();
            headerFont.Boldweight = (short)FontBoldWeight.Bold;
            headerFont.FontName = "Tahoma";
            icstyle1.Alignment = HorizontalAlignment.Center;
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            icstyle1.SetFont(headerFont);
            styleDictionary.Add("header", icstyle1);
            //
            //----- Header_Date Format -------
            //
            icstyle1.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            styleDictionary.Add("header_date", icstyle1);
            //
            //
            //---- Cell_OK_Date ------
            //
            ICellStyle okstyle = wb.CreateCellStyle();
            okstyle.Alignment = HorizontalAlignment.Left;
            okstyle.WrapText = true;
            okstyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            okstyle.SetFont(dfont);
            okstyle.FillForegroundColor = HSSFColor.LightYellow.Index;
            okstyle.FillPattern = FillPattern.SolidForeground;            // add to dictionary
            styleDictionary.Add("cell_ok_date", okstyle);
            //
            //---- Cell_Late_Date ------
            //
            //
            ICellStyle lateStyle = wb.CreateCellStyle();
            lateStyle.Alignment = HorizontalAlignment.Left;
            lateStyle.WrapText = true;
            lateStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("mmm-yy");
            lateStyle.SetFont(wfont);
            lateStyle.FillForegroundColor = HSSFColor.Grey25Percent.Index;
            lateStyle.FillPattern = FillPattern.SolidForeground;
            styleDictionary.Add("cell_late_date", lateStyle);
            //
            //
            //------ Late_Qty ------------
            //
            //
            ICellStyle lateFmt = wb.CreateCellStyle();
            lateFmt = createBorderedStyle(wb);
            lateFmt.SetFont(wfont);
            lateFmt.Alignment = HorizontalAlignment.Left;
            lateFmt.FillForegroundColor = HSSFColor.LightGreen.Index;
            lateFmt.FillPattern = FillPattern.SolidForeground;
            lateFmt.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            lateFmt.WrapText = true;
            //
            styleDictionary.Add("Late_Qty", lateFmt);
            //
            //--------OnTime_Qty -----------
            //
            //
            ICellStyle ontimeFmt = wb.CreateCellStyle();
            ontimeFmt = createBorderedStyle(wb);
            ontimeFmt.SetFont(dfont);
            ontimeFmt.Alignment = HorizontalAlignment.Left;
            ontimeFmt.FillForegroundColor = HSSFColor.LightGreen.Index;
            ontimeFmt.FillPattern = FillPattern.SolidForeground;
            ontimeFmt.DataFormat = HSSFDataFormat.GetBuiltinFormat("@");
            ontimeFmt.WrapText = true;
            //
            styleDictionary.Add("OK_Qty", ontimeFmt);
            //
            return styleDictionary;
        }
        //-----------------------------------------------------------------
        // create bordered style for cells
        //-----------------------------------------------------------------
        private static ICellStyle createBorderedStyle(IWorkbook wb)
        {
            ICellStyle style = wb.CreateCellStyle();
            style.BorderRight = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderBottom = BorderStyle.Thin;
            style.RightBorderColor = HSSFColor.Grey25Percent.Index;
            style.LeftBorderColor = HSSFColor.Grey25Percent.Index;
            style.TopBorderColor = HSSFColor.Grey25Percent.Index;
            style.BottomBorderColor = HSSFColor.Grey25Percent.Index;
            return style;
        }

        //------------------------------------------------------------------------
        // retrieve Item Warehouse Master Data from STAGING  
        //------------------------------------------------------------------------
        private static Dictionary<string, string> getItemWhseMaster(SqlConnection conn)
        {

            string mqry = "select distinct rtrim(ltrim(a.Item))+RTRIM(ltrim(a.Whse))+RTRIM(ltrim(a.State)) combo," +
                          "'Status: ' + RTRIM(ltrim(a.Whstatus)) + ' - Pareto: ' + RTRIM(ltrim(a.Pareto)) + " +
                          "' - Fcst: ' + RTRIM(ltrim(a.FcstMethod)) cmbval from dbo.DSX_ITEM_WAREHOUSE_MASTER a";

            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            objDataAdapter1.SelectCommand.CommandTimeout = 10000;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0].AsEnumerable().ToDictionary(r => r.Field<string>("combo").Trim(),
                                                                     v => v.Field<string>("cmbval").Trim());
        }
        //------------------------------------------------------------------------
        // retrieve Item Master from STAGING - explicit field order 
        //------------------------------------------------------------------------

        private static Dictionary<string, string> getItemMaster(SqlConnection conn)
        {

            string mqry = "select item,itemdesc from  dbo.[MVXItemMaster] where itemdesc is not null";

            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            objDataAdapter1.SelectCommand.CommandTimeout = 10000;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0].AsEnumerable().ToDictionary(r => r.Field<string>("Item").Trim(),
                                                                     v => v.Field<string>("ItemDesc").Trim());
        }

        //------------------------------------------------------------------------
        // retrieve Item Master from STAGING - explicit field order 
        //------------------------------------------------------------------------

        private static Dictionary<string, string> getBUMaster(SqlConnection conn)
        {

            string mqry = "select item,bushort from  dbo.[MVXItemMaster] where bushort is not null";

            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            objDataAdapter1.SelectCommand.CommandTimeout = 10000;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0].AsEnumerable().ToDictionary(r => r.Field<string>("item").Trim(),
                                                                     v => v.Field<string>("bushort").Trim());
        }
    }
}
