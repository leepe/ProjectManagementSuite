using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace ProjectManagementSuite.CSharpLogic
{
    public class ReadProjectSheets
    {
        //..........................................................................
        //  file directory  get uploaded excel files (just pick *.xls !!! )
        //..........................................................................
        //
        public static List<FileInfo> getUploadedFiles()
        {
            // must use NET 3.5 at least
            var directory = new DirectoryInfo(HttpContext.Current.Server.MapPath("~/uploads"));
            List<FileInfo> xlswb = directory.GetFiles(@"*.xls", SearchOption.AllDirectories)
                                            .OrderByDescending(x => x.FullName)
                                            .ToList();
            return xlswb;

        }
        //..........................................................................
        // pass fileInfo to spreadsheet reading routine
        //..........................................................................
        //
        public static List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> readUploadedExcelFiles(List<FileInfo> fx)
        {
            List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> lph =
                  new List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList>();

            foreach (var g in fx)
            {
                // instantiate header + lines outside read loop
                ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList pcomb = new
                    ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList();
                //
                using (FileStream file = new FileStream(g.FullName, FileMode.Open, FileAccess.Read))
                {
                    //.................................................................
                    // do not use XSSF version - expects xlsx HSSF version expects xls
                    // will generate WEB API INTERNAL SERVER ERROR otherwise
                    //.................................................................
                    IWorkbook xfn = new HSSFWorkbook(file);
                    //.................................................................
                    // instantiate projectHeader class outside sheet loop
                    //.................................................................
                    ProjectManagementSuite.Models.ProjectSheets.projectHeader ph =
                         new ProjectManagementSuite.Models.ProjectSheets.projectHeader();
                    //.................................................................
                    foreach (HSSFSheet xf in xfn)
                    {
                        //-- read header information for project ---
                        if (xf.SheetName.Contains("Proj~"))
                        {
                            // need to manage ***New*** as potential new projectID
                            int intNum;
                            bool isItInt = int.TryParse(xf.GetRow(0).GetCell(2).ToString(), out intNum);
                            ph.projectID = intNum;
                            ph.projectNumber = xf.GetRow(1).GetCell(2).ToString();
                            ph.projectName = xf.GetRow(2).GetCell(2).ToString();
                            ph.projectDesc = xf.GetRow(3).GetCell(2).ToString();
                            ph.projectFcst = xf.GetRow(5).GetCell(2).ToString().Trim().Length == 0 ? "Yes" : "No";
                            // split excel cell based on newline and don't return any empty string order numbers
                            ph.mvxorders = xf.GetRow(5).GetCell(2).ToString()
                                             .Split(new string[] { "\n" }, StringSplitOptions.None)
                                             .Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
                            //-- read project line details for the project sheet --
                            List<ProjectManagementSuite.Models.ProjectSheets.projectDetails> pline =
                                loadLineData(xf, ph.projectID);
                            // add lines to projectHeadPlusList object
                            pcomb.pList = pline;
                        }
                        else if (xf.SheetName.Equals("MasterControlSheet"))
                        {
                            ph.projectState = xf.GetRow(0).GetCell(1).ToString();
                        }
                    }
                    // add header to projectHeadPlusList object
                    pcomb.pHead = ph;
                }
                // add lines plus header to list
                lph.Add(pcomb);
            }
            return lph;
        }
        //--------------------------------------------------------------------
        // load project workbook line data into projectDetails LIST
        //--------------------------------------------------------------------
        private static List<ProjectManagementSuite.Models.ProjectSheets.projectDetails> loadLineData(HSSFSheet xl, int projID)
        {
            List<ProjectManagementSuite.Models.ProjectSheets.projectDetails> preLimLines
                  = new List<ProjectManagementSuite.Models.ProjectSheets.projectDetails>();
            // now read project data off body of sheet
            for (int w = 9; w < xl.LastRowNum; w++)
            {
                for (int z = 6; z < 30; z++)
                {
                    // need to redeclare class so that new details are added without changing previous additions
                    ProjectManagementSuite.Models.ProjectSheets.projectDetails pd =
                        new ProjectManagementSuite.Models.ProjectSheets.projectDetails();
                    if (xl.GetRow(w).GetCell(z) != null)
                    {
                        if (!string.IsNullOrEmpty(xl.GetRow(w).GetCell(z).ToString()))                   // just read and add projectDetails with data - eliminate empty rows
                        {
                            pd.projectID = projID;
                            pd.projectItem = xl.GetRow(w).GetCell(1).ToString().Trim().ToUpper();        // item codes - trim and capitalise item codes
                            pd.projectWhse = xl.GetRow(w).GetCell(3).ToString().Trim();                  // item whse
                            pd.projectDate = Convert.ToInt32(xl.GetRow(8).GetCell(z).DateCellValue.ToString("yyyyMMdd"));
                            pd.projectQty = Convert.ToInt32(xl.GetRow(w).GetCell(z).ToString());
                            preLimLines.Add(pd);
                        }
                    }
                }

            }
            // need to apply checks that item and whse have been given some values
            // this is different to validating that MOVEX item-whse combinations exist.
            // filter out : BLANK ITEM - BLANK WHSE

            var qry1 = (from t in preLimLines
                        group t by new { t.projectID, t.projectItem, t.projectWhse, t.projectDate } into g
                        select new
                        {
                            projectID = g.Key.projectID,
                            Item = g.Key.projectItem,
                            Whse = g.Key.projectWhse,
                            Date = g.Key.projectDate,
                            tQty = g.Sum(a => a.projectQty) // Sum, not Max
                        }).Where(i => (i.Item.Length != 0 && i.Whse.Length != 0));


            List<ProjectManagementSuite.Models.ProjectSheets.projectDetails> finalLines =
                               new List<ProjectManagementSuite.Models.ProjectSheets.projectDetails>();
            // loop through grouping and sum across any groups  - eliminate duplicate lines
            foreach (var t in qry1)
            {
                ProjectManagementSuite.Models.ProjectSheets.projectDetails flines =
                            new ProjectManagementSuite.Models.ProjectSheets.projectDetails();
                flines.projectID = t.projectID;
                flines.projectItem = t.Item;
                flines.projectWhse = t.Whse;
                flines.projectDate = t.Date;
                flines.projectQty = t.tQty;
                finalLines.Add(flines);
            }
            //
            return finalLines;
        }

        //------------------------------------------------------------------------
        // retrieve project Number & Name from STAGING  use :
        // originalDictionary.ToDictionary(kp=>kp.value,kp=kp.key)
        //------------------------------------------------------------------------
        public static Dictionary<string, string> getExistingProjectNumbers(SqlConnection conn)
        {

            string mqry = "select distinct ltrim(rtrim(a.ProjNum))+ltrim(rtrim(a.ProjName)) Key1,ltrim(rtrim(a.ProjName)) Val1 from dbo.Projects a";

            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            objDataAdapter1.SelectCommand.CommandTimeout = 10000;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0].AsEnumerable().ToDictionary(r => r.Field<string>("Key1").Trim(),
                                                                     v => v.Field<string>("Val1").Trim());
        }

        //---------------------------------------------------------------------------
        // delete out duplicated manually forecasted projects in different workbooks
        //---------------------------------------------------------------------------
        public static List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList>
                      deleteOutDuplicateProjects(List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> phpl)
        {
            // returns a distinct list of projects
            List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> distList
                     = phpl.GroupBy(i => new { i.pHead.projectNumber, i.pHead.projectName, i.pHead.projectDesc, i.pHead.projectFcst, i.pHead.projectState })
                           .Select(group => group.First()).ToList();
            return distList;

        }
        //..........................................................................
        //  check out Number + Name combinations
        //..........................................................................
        //
        public static void probList(Dictionary<string, string> masterDic, List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> chkList)
        {
            // set already-exists flag to true if projectNumber + projectName already exists
            foreach (var pq in chkList.AsEnumerable())
            {   // if found in dictionary add to problem list
                if (masterDic.ContainsKey(pq.pHead.projectNumber.Trim() + pq.pHead.projectName.Trim())) pq.pHead.alreadyExists = true;
            }
        }
        //..........................................................................
        //  add records to DB - of new projects
        //..........................................................................
        //
        public static void addNewProjectsToDB(List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> fphl, SqlConnection cn)
        {
            cn.Open();
            // boolean alreadyExists is false
            foreach (ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList pq in fphl.AsEnumerable().Where(z => !z.pHead.alreadyExists))
            {
                // set up Object to retrieve AUTO-INCREMENT
                Object returnValue;
                //------ String builder for concatenating movex orders together - remove \n
                StringBuilder sb = new StringBuilder();
                sb.Append(string.Empty);
                foreach (var h in pq.pHead.mvxorders) sb.Append(h.Replace("\n", string.Empty));
                // need to pick up the newly generated AUTO-INCREMENT integer identifier to add to dbo.[ProjectItems]
                string headQry = string.Format("insert into dbo.[Projects] (ProjNum,ProjName,ProjDesc,State,ProjManFlag,MVXProjNum) " +
                                                  "OUTPUT INSERTED.ProjID VALUES ('{0}','{1}','{2}','{3}','{4}','{5}')",
                                                  pq.pHead.projectNumber, pq.pHead.projectName, pq.pHead.projectDesc,
                                                  pq.pHead.projectState, pq.pHead.projectFcst, sb.ToString());
                //
                SqlCommand uplH = new SqlCommand(headQry, cn);
                uplH.CommandType = CommandType.Text;
                returnValue = uplH.ExecuteScalar();
                uplH.Dispose();
                //
                foreach (ProjectManagementSuite.Models.ProjectSheets.projectDetails g in pq.pList)
                {
                    //----- Add in project lines with returned AUTO-INCREMENT project id
                    string lineQry = string.Format("insert into dbo.[ProjectItems] (ProjID,ProjItem,Whse,Month,Qty) " +
                                                      "VALUES ({0},'{1}','{2}', {3} , {4})",
                                                      Convert.ToInt32(returnValue), g.projectItem, g.projectWhse, g.projectDate, g.projectQty);
                    //
                    SqlCommand uplL = new SqlCommand(lineQry, cn);
                    uplL.CommandType = CommandType.Text;
                    uplL.ExecuteNonQuery();
                    uplL.Dispose();
                }

            }
            cn.Close();
        }

    }
}