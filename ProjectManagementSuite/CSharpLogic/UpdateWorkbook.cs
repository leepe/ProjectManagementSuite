using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using ProjectManagementSuite.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;

namespace ProjectManagementSuite.CSharpLogic
{
    public class UpdateWorkbook
    {
        //-----------------------------------------------------------------
        // write out project lines to project workbook in standard format
        //-----------------------------------------------------------------
        public static void updateProjectWorkBookLines(newProject nph, DataTable dt,string f0)
        {
            using (FileStream fprojects = new FileStream(f0, FileMode.Open, FileAccess.Read))
            {
                if (dt.Rows.Count > 0)      // if there is no data don't try to write anything
                {
                    // set up NPOI objects
                    IWorkbook xWrite = new HSSFWorkbook(fprojects);
                    HSSFSheet xSheet1 = (HSSFSheet)xWrite.GetSheetAt(0);         // main excel project sheet
                    //------------------------------------------------------------
                    // load up dictionaries to map item / item warehouse / bu
                    //------------------------------------------------------------
                    Dictionary<string, string> dcn = ProjectManagementSuite.CSharpLogic.ManageData.getItemMaster();
                    Dictionary<string, string> diw = ProjectManagementSuite.CSharpLogic.ManageData.getItemWhseMaster();
                    Dictionary<string, string> dbu = ProjectManagementSuite.CSharpLogic.ManageData.getBUMaster();
                    //-----------------------------------------------------------
                    // iWarning Font - to use with dictionaries
                    //-----------------------------------------------------------
                    IFont iWarn = xWrite.CreateFont();
                    iWarn.Color = HSSFColor.Red.Index;
                    iWarn.Boldweight = (short)FontBoldWeight.Bold;
                    //-----------------------------------------------------------
                    // pull together cell style dictionary
                    //-----------------------------------------------------------
                    //
                    Dictionary<string, ICellStyle> sd = ProjectManagementSuite.CSharpLogic.GenerateWorkbook.createStyles(xWrite);
                    //
                    //-----------------------------------------------------------------------------------------------------------
                    // work out where the first date is in the datatable
                    //-----------------------------------------------------------------------------------------------------------
                    //
                    DataRow firstMnth = (from g in dt.AsEnumerable() select g).OrderBy(p => p.Field<int>("Month")).First();
                    DateTime d0 = DateTime.ParseExact(firstMnth.Field<int>("Month").ToString(), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                    DateTime startOfMonth = new DateTime(d0.Year, d0.Month, 1).AddMonths(-2);
                    DateTime dNow = DateTime.Now;
                    Int32 Datediff = ((dNow.Year - startOfMonth.Year) * 12) + dNow.Month - startOfMonth.Month > 0 ?
                                     ((dNow.Year - startOfMonth.Year) * 12) + dNow.Month - startOfMonth.Month : 0;
                    //
                    // loop through excel sheet pick up the dates on the sheet
                    string[] dte = new string[23];
                    for (int z = 0; z < 23; z++)
                    {
                        xSheet1.GetRow(8).GetCell(6 + z).SetCellValue(startOfMonth.AddMonths(z));
                        dte[z] = xSheet1.GetRow(8).GetCell(6 + z).DateCellValue.ToString("yyyyMMdd");
                        if (z < Datediff)
                        {
                            xSheet1.GetRow(8).GetCell(6 + z).CellStyle = sd["cell_late_date"];
                        }
                        else
                        {
                            xSheet1.GetRow(8).GetCell(6 + z).CellStyle = sd["cell_ok_date"];
                        }
                    }
                    //------------------------------------------------------------
                    // run out data onto spreadsheet
                    //------------------------------------------------------------
                    var numitem = (from h in dt.AsEnumerable()
                                   select new
                                   {
                                       Item = h.Field<string>("ProjItem").Trim(),
                                       Whse = h.Field<string>("Whse").Trim()
                                   }).Distinct();             // number of distinct items ie number of lines to write
                    int ncount = 0;
                    foreach (var j in numitem)
                    {
                        var subqry = (from w in dt.AsEnumerable() select w).Where(z => z.Field<string>("ProjItem").Trim().Equals(j.Item));
                        xSheet1.GetRow(ncount + 9).GetCell(1).SetCellValue(j.Item.Trim());
                        xSheet1.GetRow(ncount + 9).GetCell(3).SetCellValue(j.Whse.Trim());
                        //...............................................................................................................
                        // add in item description  - first dictionary D C N 
                        //...............................................................................................................
                        xSheet1.GetRow(ncount + 9).CreateCell(2);                                         // need to create cell first
                        string kval = string.Empty;
                        if (dcn.TryGetValue(j.Item.Trim(), out kval))                                      // protect dictionary
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(2).SetCellValue(kval);                       // if the item exists in dictionary
                        }
                        else
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(2).SetCellValue("ITEM DOES NOT EXIST");      // if the item doesn't exist
                            xSheet1.GetRow(ncount + 9).GetCell(2).RichStringCellValue.ApplyFont(iWarn);     // highlight in red
                        }
                        //...............................................................................................................
                        // add in item warehouse data  - second dictionary D I W 
                        //...............................................................................................................
                        xSheet1.GetRow(ncount + 9).CreateCell(4);                                           // create new cell first
                        string iwval = string.Empty;
                        string cmbStr = string.Concat(j.Item.Trim(), j.Whse.Trim(), nph.projectState);
                        if (diw.TryGetValue(cmbStr, out iwval))                // testing itme X whse X state combination
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(4).SetCellValue(iwval);
                        }
                        else
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(4).SetCellValue(string.Concat(j.Item.Trim(), "-", j.Whse.Trim(), "-", nph.projectState) + " DOES NOT EXIST");
                            xSheet1.GetRow(ncount + 9).GetCell(4).RichStringCellValue.ApplyFont(iWarn);
                        }
                        //...............................................................................................................
                        // add in business  - third dictionary D B U 
                        //...............................................................................................................
                        xSheet1.GetRow(ncount + 9).CreateCell(5);
                        string ibval = string.Empty;
                        if (dbu.TryGetValue(j.Item.Trim(), out ibval))
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(5).SetCellValue(ibval);
                        }
                        else
                        {
                            xSheet1.GetRow(ncount + 9).GetCell(5).SetCellValue("BU DOESNT EXIST");                  // if the BUShort doesn't exist
                            xSheet1.GetRow(ncount + 9).GetCell(5).RichStringCellValue.ApplyFont(iWarn);
                        }
                        //................................................................................................................
                        // now write out the monthly data
                        //................................................................................................................
                        foreach (var s in subqry)
                        {
                            if (Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()) > -1)  // if the Month is found in the date array
                            {
                                //------ write out the monthly data for editing -------------
                                xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim())).SetCellType(CellType.Numeric);
                                xSheet1.GetRow(ncount + 9).GetCell(6 + Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()))
                                                          .SetCellValue(Double.Parse(s.Field<decimal>("Qty").ToString("F0")));
                                // format the monthly data if it's late or not
                                if (Array.IndexOf(dte, s.Field<int>("Month").ToString().Trim()) < Datediff)
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
        }
        //------------------------------------------------------------------------------------------------------------------------------------------------------------
        // generate a product issues workbook for each state with problems
        //------------------------------------------------------------------------------------------------------------------------------------------------------------
        public static void generateProductIssuesLog(string f0)
        {
            // template workbook with tabs for each state
            string ftempl = HttpContext.Current.Server.MapPath("~/CSharpLogic") + "/ProblemItemsTemplate.xls";
            //
            //--------------------------------------------------------
            // get datatables of diagnostic product data
            //--------------------------------------------------------
            //
            DataTable dtFirst = ProjectManagementSuite.CSharpLogic.ManageData.getNonExistentItemsAcrossProjects();
            DataTable dtSecond = ProjectManagementSuite.CSharpLogic.ManageData.getNonExistentItemWhseStatesAcrossProjects();
            DataTable dtThird = ProjectManagementSuite.CSharpLogic.ManageData.getNonForecastItemWhseStatesAcrossProjects();
            //--------------------------------------------------------
            // define titles for each of the datatables
            //--------------------------------------------------------
            //
            string[] firstHeader = { "State", "ProjID", "ProjectName", "ProjectItem", "Whse", "Qty" };
            string[] secondHeader = { "State", "ProjID", "ProjectName", "ProjectItem", "ItemDesc", "Whse", "Qty", "Transactions" };
            string[] thirdHeader = { "BU", "State", "ProjID", "ProjectName", "ProjectItem", "ItemDesc", "Whse", "Qty", "Pareto", "FcstMethod" };
            //--------------------------------------------------------
            // create workbook and define sheets
            //--------------------------------------------------------
            using (FileStream ftemplate = new FileStream(ftempl, FileMode.Open, FileAccess.ReadWrite))
            {
                IWorkbook workbook = new HSSFWorkbook(ftemplate);
                //-----------------------------------------------------------------------
                // load up dictionary of sheet names for LINQ and workbook output
                //-----------------------------------------------------------------------
                //
                Dictionary<int, string> stateDic = new Dictionary<int, string>();
                for (int g = 0; g < workbook.NumberOfSheets; g++) stateDic.Add(g, workbook.GetSheetAt(g).SheetName);
                //
                //------------------------------------------------------------------------
                // big loop through the workbook downloading issues by sheet & state
                //------------------------------------------------------------------------
                //
                for (int g = 0; g < workbook.NumberOfSheets; g++)
                {
                    ISheet sheet = workbook.GetSheetAt(g);
                    IRow headerRow = sheet.CreateRow(0); // this is zero based row numbering
                    ICell cell = headerRow.CreateCell(0);
                    cell.SetCellValue("THESE ITEMS DO NOT EXIST:");
                    IRow title1Row = sheet.CreateRow(1);
                    for (int z = 0; z < firstHeader.Length; z++)
                    {
                        title1Row.CreateCell(z).SetCellValue(firstHeader[z]);
                    }
                    // set up master row counter to drive row creation
                    int rowcount = 2;
                    //--------------------------------------------------------------------------------------------------------
                    // FIRST QUERY : NON-EXISTENT ITEMS
                    //--------------------------------------------------------------------------------------------------------
                    //

                    foreach (var k in dtFirst.AsEnumerable().Where(p => p.Field<string>("State").Trim().Equals(stateDic[g])).OrderBy(p => p.Field<int>("ProjID")))
                    {
                        IRow nrow = sheet.CreateRow(rowcount);
                        ICell cell1 = nrow.CreateCell(0);
                        cell1.SetCellValue(k.Field<string>("State").Trim());
                        ICell cell2 = nrow.CreateCell(1);
                        cell2.SetCellValue(k.Field<int>("ProjID"));
                        ICell cell3 = nrow.CreateCell(2);
                        cell3.SetCellValue(k.Field<string>("ProjName").Trim());
                        ICell cell4 = nrow.CreateCell(3);
                        cell4.SetCellValue(k.Field<string>("ProjItem").Trim());
                        ICell cell5 = nrow.CreateCell(4);
                        cell5.SetCellValue(k.Field<string>("Whse").Trim());
                        ICell cell6 = nrow.CreateCell(5);
                        cell6.SetCellValue(k.Field<decimal>("Qty").ToString("F0"));
                        rowcount++;
                    }

                    rowcount += 2; // add a few blank lines
                    IRow headerRow2 = sheet.CreateRow(rowcount);
                    ICell cellNext = headerRow2.CreateCell(0);
                    cellNext.SetCellValue("THESE ITEM X WAREHOUSE X STATE COMBINATIONS DO NOT EXIST:");
                    rowcount++;   // add a new blank line
                    IRow title2Row = sheet.CreateRow(rowcount);
                    for (int z = 0; z < secondHeader.Length; z++)
                    {
                        title2Row.CreateCell(z).SetCellValue(secondHeader[z]);
                    }
                    rowcount++;    // add another blank line
                    //--------------------------------------------------------------------------------------------------------
                    // SECOND QUERY : NON-EXISTENT ITEM X WAREHOUSE X STATE COMBINATIONS
                    //--------------------------------------------------------------------------------------------------------
                    //
                    foreach (var i in dtSecond.AsEnumerable().Where(p => p.Field<string>("State").Trim().Equals(stateDic[g])).OrderBy(p => p.Field<int>("ProjID")))
                    {
                        IRow nrow = sheet.CreateRow(rowcount);
                        ICell cell1 = nrow.CreateCell(0);
                        cell1.SetCellValue(i.Field<string>("State").Trim());
                        ICell cell2 = nrow.CreateCell(1);
                        cell2.SetCellValue(i.Field<int>("ProjID"));
                        ICell cell3 = nrow.CreateCell(2);
                        cell3.SetCellValue(i.Field<string>("ProjName").Trim());
                        ICell cell4 = nrow.CreateCell(3);
                        cell4.SetCellValue(i.Field<string>("ProjItem").Trim());
                        ICell cell5 = nrow.CreateCell(4);
                        cell5.SetCellValue(i.Field<string>("ItemDesc").Trim());
                        ICell cell6 = nrow.CreateCell(5);
                        cell6.SetCellValue(i.Field<string>("Whse").Trim());
                        ICell cell7 = nrow.CreateCell(6);
                        cell7.SetCellValue(i.Field<decimal>("Qty").ToString("F0"));
                        ICell cell8 = nrow.CreateCell(7);
                        cell8.SetCellValue(i.Field<int>("Transactions").ToString("D0"));
                        rowcount++;

                    }
                    rowcount += 2; // add a few blank lines
                    IRow headerRow3 = sheet.CreateRow(rowcount);
                    ICell cellNext2 = headerRow3.CreateCell(0);
                    cellNext2.SetCellValue("THESE ITEM X WAREHOUSE X STATE COMBINATIONS ARE NON-FORECAST:");
                    rowcount++;   // add a new blank line
                    IRow title3Row = sheet.CreateRow(rowcount);
                    for (int z = 0; z < thirdHeader.Length; z++)
                    {
                        title3Row.CreateCell(z).SetCellValue(thirdHeader[z]);
                    }
                    rowcount++;    // add another blank line
                    //--------------------------------------------------------------------------------------------------------
                    // THIRD QUERY : NON-FORECAST ITEM X WAREHOUSE X STATE COMBINATIONS
                    //--------------------------------------------------------------------------------------------------------
                    //
                    foreach (var i in dtThird.AsEnumerable().Where(p => p.Field<string>("State").Trim().Equals(stateDic[g])).OrderBy(p => p.Field<int>("ProjID")))
                    {
                        IRow nrow = sheet.CreateRow(rowcount);
                        ICell cell1 = nrow.CreateCell(0);
                        cell1.SetCellValue(i.Field<string>("BUShort").Trim());
                        ICell cell2 = nrow.CreateCell(1);
                        cell2.SetCellValue(i.Field<string>("State"));
                        ICell cell3 = nrow.CreateCell(2);
                        cell3.SetCellValue(i.Field<int>("ProjID"));
                        ICell cell4 = nrow.CreateCell(3);
                        cell4.SetCellValue(i.Field<string>("ProjName").Trim());
                        ICell cell5 = nrow.CreateCell(4);
                        cell5.SetCellValue(i.Field<string>("ProjItem").Trim());
                        ICell cell6 = nrow.CreateCell(5);
                        cell6.SetCellValue(i.Field<string>("ItemDesc").Trim());
                        ICell cell7 = nrow.CreateCell(6);
                        cell7.SetCellValue(i.Field<string>("Whse").Trim());
                        ICell cell8 = nrow.CreateCell(7);
                        cell8.SetCellValue(i.Field<decimal>("Qty").ToString("F0"));
                        ICell cell9 = nrow.CreateCell(8);
                        cell9.SetCellValue(i.Field<string>("Pareto").Trim());
                        ICell cell10 = nrow.CreateCell(9);
                        cell10.SetCellValue(i.Field<string>("FcstMethod").Trim());
                        rowcount++;
                    }
                }
                //------------------------------------------------------------------------
                FileStream sw = File.Create(f0);
                workbook.Write(sw);
                sw.Close();
            }

        }
    }
}