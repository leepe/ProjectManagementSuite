using ProjectManagementSuite.Models;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ProjectManagementSuite.CSharpLogic
{
    //-------------------------------------------------------------------------------------------------
    // CRUD operations on Project Database
    //-------------------------------------------------------------------------------------------------
    public class ManageClientData
    {
        //---------------------------------------------------------------------------------------------
        // delete project out of database
        //---------------------------------------------------------------------------------------------
        //
        public static void DeleteProjectOutOfDatabase(int Id)
        {
            // connect to SQLEXPRESS
            SqlConnection scnxn = ProjectManagementSuite.CSharpLogic.ManageData.setUpMSSQLconn();
            scnxn.Open();
            // sqlcommand object
            SqlCommand Cmd1 = new SqlCommand("delete from dbo.[Projects] where ProjID=" + Id, scnxn);
            Cmd1.CommandType = CommandType.Text;
            Cmd1.ExecuteNonQuery();
            Cmd1.Dispose();
            SqlCommand Cmd2 = new SqlCommand("delete from dbo.[ProjectItems] where ProjID=" + Id, scnxn);
            Cmd2.CommandType = CommandType.Text;
            Cmd2.ExecuteNonQuery();
            Cmd2.Dispose();
            scnxn.Close();
        }
        //---------------------------------------------------------------------------------------------
        // get project from database
        //---------------------------------------------------------------------------------------------
        //
        public static IEnumerable<ProjectFcst> GetProjectDataFromDatabase(int Id)
        {

            // get datatable
            DataTable dt = GetProjectHeaderForProject(Id);
            // even though it's a single record use Ienumerable
            IEnumerable<ProjectFcst> indiv = from g in dt.AsEnumerable()
                                            select new ProjectFcst
                                            {
                                                ProjID = g.Field<int>("ProjId"),
                                                ProjNum = g.Field<string>("ProjNum").Trim(),
                                                ProjName = g.Field<string>("ProjName").Trim(),
                                                ProjDesc = string.IsNullOrWhiteSpace(g.Field<string>("ProjDesc")) ? string.Empty :
                                                           g.Field<string>("ProjDesc").Trim(),
                                                State = g.Field<string>("State").Trim(),
                                                ProjManFlag = g.Field<string>("ProjManFlag").Trim(),
                                                MVXProjNum2 = string.Join("\t", string.IsNullOrEmpty(g.Field<string>("MVXProjNum")) ? new List<string>() :
                                                             Enumerable.Range(0, g.Field<string>("MVXProjNum").Trim().Length / 10)
                                                             .Select(i => g.Field<string>("MVXProjNum").Trim().Substring(i * 10, 10)).ToList())
                                            };
            return indiv;
        }
        //------------------------------------------------------------------------
        // retrieve project header data from express for project lines controller
        //------------------------------------------------------------------------
        //
        public static newProject GetProjectHeaderForProjectLines(int Id)
        {
            // retrieve header record for project line request
            DataTable dt = ProjectManagementSuite.CSharpLogic.ManageClientData.GetProjectHeaderForProject(Id);
            // add returned datatable to newProject class object
            newProject oldProj = new newProject();
            foreach (var g in dt.AsEnumerable())
            {
                oldProj.projectNumber = g.Field<string>("ProjNum").Trim();
                oldProj.projectName = g.Field<string>("ProjName").Trim();
                oldProj.projectType = g.Field<string>("ProjDesc").Trim();
                oldProj.projectState = g.Field<string>("State").Trim();
                List<string> ords = Enumerable.Range(0, g.Field<string>("MVXProjNum").Trim().Length / 10)
                                           .Select(i => g.Field<string>("MVXProjNum").Trim().Substring(i * 10, 10)).ToList();
                // instantiate a new List of mvxorders
                oldProj.mvxorders = new List<mvxorders>();
                // add each string to new list
                foreach (var f in ords)
                {
                    mvxorders newb = new mvxorders();
                    newb.order = f.Trim();
                    oldProj.mvxorders.Add(newb);
                }
            }
            return oldProj;
        }

        //------------------------------------------------------------------------
        // retrieve project header data from express
        //------------------------------------------------------------------------
        //
        public static DataTable GetProjectHeaderForProject(int Id)
        {
            SqlConnection conn = ProjectManagementSuite.CSharpLogic.ManageData.setUpMSSQLconn();
            // get project header records
            string sql = "select ProjID,ProjNum,ProjName,ProjDesc,State,ProjManFlag," +
                         "isnull(MVXProjNum,'') MVXProjNum from dbo.[Projects] where ProjID=" + Id;
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
        // update project data from express
        //------------------------------------------------------------------------
        //
        public static void UpdateProjectDataInDB(ProjectFcst pf)
        {
            // connect to SQLEXPRESS
            SqlConnection scnxn = ProjectManagementSuite.CSharpLogic.ManageData.setUpMSSQLconn();
            SqlCommand cmd01 = new SqlCommand();
            cmd01.CommandType = CommandType.Text;
            cmd01.CommandText = "UPDATE dbo.Projects SET [ProjNum] = @ProjNum WHERE [ProjID] = @ProjID";
            cmd01.Parameters.AddWithValue("@ProjNum", pf.ProjNum);
            cmd01.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd01.Connection = scnxn;
            scnxn.Open();
            cmd01.ExecuteNonQuery();
            scnxn.Close();
            SqlCommand cmd02 = new SqlCommand();
            cmd02.CommandType = CommandType.Text;
            cmd02.CommandText = "UPDATE dbo.Projects SET [ProjName] = @ProjName WHERE [ProjID] = @ProjID";
            cmd02.Parameters.AddWithValue("@ProjName", pf.ProjName);
            cmd02.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd02.Connection = scnxn;
            scnxn.Open();
            cmd02.ExecuteNonQuery();
            scnxn.Close();
            SqlCommand cmd03 = new SqlCommand();
            cmd03.CommandType = CommandType.Text;
            cmd03.CommandText = "UPDATE dbo.Projects SET [ProjDesc] = @ProjDesc WHERE [ProjID] = @ProjID";
            cmd03.Parameters.AddWithValue("@ProjDesc", pf.ProjDesc);
            cmd03.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd03.Connection = scnxn;
            scnxn.Open();
            cmd03.ExecuteNonQuery();
            scnxn.Close();
            SqlCommand cmd04 = new SqlCommand();
            cmd04.CommandType = CommandType.Text;
            cmd04.CommandText = "UPDATE dbo.Projects SET [State] = @State WHERE [ProjID] = @ProjID";
            cmd04.Parameters.AddWithValue("@State", pf.State);
            cmd04.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd04.Connection = scnxn;
            scnxn.Open();
            cmd04.ExecuteNonQuery();
            scnxn.Close();
            SqlCommand cmd05 = new SqlCommand();
            cmd05.CommandType = CommandType.Text;
            cmd05.CommandText = "UPDATE dbo.Projects SET [ProjManFlag] = @ProjManFlag WHERE [ProjID] = @ProjID";
            cmd05.Parameters.AddWithValue("@ProjManFlag", pf.ProjManFlag);
            cmd05.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd05.Connection = scnxn;
            scnxn.Open();
            cmd05.ExecuteNonQuery();
            scnxn.Close();
            SqlCommand cmd06 = new SqlCommand();
            cmd06.CommandType = CommandType.Text;
            cmd06.CommandText = "UPDATE dbo.Projects SET [MVXProjNum] = @MVXProjNum2 WHERE [ProjID] = @ProjID";
            cmd06.Parameters.AddWithValue("@MVXProjNum2", pf.MVXProjNum2.Replace("\t","").Replace("\n","").Trim());
            cmd06.Parameters.AddWithValue("@ProjID", pf.ProjID);
            cmd06.Connection = scnxn;
            scnxn.Open();
            cmd06.ExecuteNonQuery();
            scnxn.Close();
        }
    }
}