using ProjectManagementSuite.Models;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace ProjectManagementSuite.CSharpLogic
{
    public class ManageData
    {

        public static IEnumerable<ProjectFcst> GetProjectHeaderData()
        {
            DataTable dtHead = GetProjectHeaders();
            // load datatable with project details 
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
                                                MVXProjNum2 = string.Join("\n", string.IsNullOrEmpty(g.Field<string>("MVXProjNum")) ? new List<string>() :
                                                             Enumerable.Range(0, g.Field<string>("MVXProjNum").Trim().Length / 10)
                                                             .Select(i => g.Field<string>("MVXProjNum").Trim().Substring(i * 10, 10)).ToList())
                                            };

            return pset;
        }

        //------------------------------------------------------------------------
        // retrieve project header data from express
        //------------------------------------------------------------------------

        private static DataTable GetProjectHeaders()
        {
            SqlConnection conn = setUpMSSQLconn();
            // get project header records
            string sql = "select ProjID,ProjNum,ProjName,ProjDesc,State,ProjManFlag," +
                         "MVXProjNum from dbo.[Projects]";
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
        // retrieve project line data for project ID
        //------------------------------------------------------------------------
        //
        public static DataTable GetProjectLinesForID(int projId)
        {
            SqlConnection cnxn = setUpMSSQLconn();
            // get project header records
            string sql = "select ProjItem,Whse,[MONTH],Qty from dbo.ProjectItems where [ProjID]=" + projId.ToString().Trim();
            cnxn.Open();
            SqlCommand SqlCmd = new SqlCommand(sql, cnxn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            cnxn.Close();
            return objDataset1.Tables[0];
        }
        //------------------------------------------------------------------------
        // retrieve movex order data for movex order numbers
        //------------------------------------------------------------------------
        //
        public static DataTable GetOpenMovexOrdersDetails(newProject np)
        {
            SqlConnection cnxn = setUpMSSQLconn();
           // now build query string of orders for MOVEX
            StringBuilder sb = new StringBuilder();
            sb.Append("(");
            foreach (var g in np.mvxorders)
            {   
                    sb.Append("'").Append(g.order).Append("',");
            }
            // remove last comma before adding closing parenthesis
            string qbit = sb.Remove(sb.Length - 1, 1).Append(")").ToString().Trim();
            // get project order records for ordernumbers
            //
            string sql = "select distinct ProjItem,Whse,[Month],sum(Qty) Qty from dbo.[ProjectOpenOrders] where OrderNo in " + qbit + 
                         " group by ProjItem,Whse,[Month]";
            cnxn.Open();
            SqlCommand SqlCmd = new SqlCommand(sql, cnxn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            cnxn.Close();
            return objDataset1.Tables[0];
        }
        //------------------------------------------------------------------------------------------
        // retrieve all projects across all states with non-forecast item X whse X states locations
        // --- number of result columns = 10 ----
        //------------------------------------------------------------------------------------------
        //
        public static DataTable getNonForecastItemWhseStatesAcrossProjects()
        {
            SqlConnection conn = setUpMSSQLconn();
            string mqry = "with currItems as " +
                          "( " +
                          "select distinct b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse,SUM(a.Qty) Qty from dbo.[ProjectItems] a " +
                          "inner join dbo.[Projects] b on (a.ProjID=b.ProjID) where " +
                          "a.[Month] >= cast(convert(varchar(10),DATEADD(dd, -DAY(GETDATE()) + 1, GETDATE()),112) as integer) " +
                          "and b.ProjManFlag='Yes' group by b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse " +
                          ")," +
                          "ItemWh as " +
                          "( " +
                          "select c.BUShort,currItems.State,currItems.ProjID,currItems.ProjName,currItems.ProjItem,c.ITEMDESC,currItems.Whse,currItems.Qty," +
                          "b.Pareto,b.FcstMethod from currItems inner join dbo.[DSX_ITEM_WAREHOUSE_MASTER] b on " +
                          "(currItems.ProjItem=b.Item and currItems.Whse=b.Whse and currItems.State=b.State) " +
                          "inner join dbo.[MVXItemMaster] c on (currItems.ProjItem=c.ITEM) " +
                          "where b.Pareto in ('F','G','J') and b.FcstMethod not in ('Y') " +
                          ") " +
                          "select ItemWh.BUShort,ItemWh.State,ItemWh.ProjID,ItemWh.ProjName,ItemWh.ProjItem,ItemWh.ITEMDESC," +
                          "ItemWh.Whse,ItemWh.Qty,ItemWh.Pareto,ItemWh.FcstMethod from ItemWh ";
            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }
        //---------------------------------------------------------------------------------
        // retrieve all projects across all states with non-existant item X whse X states
        // --- number of result columns = 8 ----
        //---------------------------------------------------------------------------------
        //
        public static DataTable getNonExistentItemWhseStatesAcrossProjects()
        {
            SqlConnection conn = setUpMSSQLconn();
            string mqry = "with currItems as " +
                          "( " +
                          "select distinct b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse,SUM(a.Qty) Qty from dbo.[ProjectItems] a " +
                          "inner join dbo.[Projects] b on (a.ProjID=b.ProjID) where " +
                          "a.[Month] >= cast(convert(varchar(10),DATEADD(dd, -DAY(GETDATE()) + 1, GETDATE()),112) as integer) " +
                          "and b.ProjManFlag='Yes' group by b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse " +
                          ")," +
                          "nonexItem as " +
                          "( " +
                          "select b.BUShort,currItems.[State],currItems.ProjID,currItems.ProjName,currItems.ProjItem,currItems.Whse,currItems.Qty from  " +
                          "currItems left join dbo.[MVXItemMaster] b on (currItems.ProjItem=b.ITEM) where b.BUShort is null " +
                          "), " +
                          "nonexItemWh as " +
                          "( " +
                          "select currItems.State,currItems.ProjID,currItems.ProjName,currItems.ProjItem,currItems.Whse,currItems.Qty,b.Pareto,b.FcstMethod from " +
                          "currItems left join dbo.[DSX_ITEM_WAREHOUSE_MASTER] b on (currItems.ProjItem=b.Item and currItems.Whse=b.Whse and currItems.State=b.State) " +
                          "where b.Item is null " +
                          "), " +
                          "exclnonexItem as " +
                          "( " +
                          "select nonexItemWh.State,nonexItemWh.ProjID,nonexItemWh.ProjName,nonexItemWh.ProjItem,nonexItemWh.Whse,nonexItemWh.Qty from " +
                          "nonexItemWh left join nonexItem on (nonexItemWh.ProjItem=nonexItem.ProjItem) where nonexItem.ProjItem is null " +
                          ") " +
                          "select exclnonexItem.State,exclnonexItem.ProjID,exclnonexItem.ProjName,exclnonexItem.ProjItem,b.ItemDesc,exclnonexItem.Whse," +
                          "exclnonexItem.Qty,isnull(c.NumLine,0) Transactions from exclnonexItem inner join dbo.[MVXItemMaster] b on (exclnonexItem.ProjItem=b.Item) " +
                          "left join dbo.[UNIQUE_BUSINESSKEYS] c on (exclnonexItem.State=c.SalesState and exclnonexItem.Whse=c.Whse and exclnonexItem.ProjItem=c.Item)";
            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }
        //------------------------------------------------------------------------
        // retrieve all projects across all states with non-existant items
        // --- number of result columns = 6 ----
        //------------------------------------------------------------------------
        //
        public static DataTable getNonExistentItemsAcrossProjects()
        {
            SqlConnection conn = setUpMSSQLconn();
            string mqry = "with currItems as " +
                          "( " +
	                      "select distinct b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse,SUM(a.Qty) Qty from dbo.[ProjectItems] a " +
                          "inner join dbo.[Projects] b on (a.ProjID=b.ProjID) where " +
	                      "a.[Month] >= cast(convert(varchar(10),DATEADD(dd, -DAY(GETDATE()) + 1, GETDATE()),112) as integer) " +
	                      "and b.ProjManFlag='Yes' group by b.[State],b.ProjID,b.ProjName,a.ProjItem,a.Whse " + 
                          ") " +
                          "select currItems.[State],currItems.ProjID,currItems.ProjName,currItems.ProjItem,currItems.Whse,currItems.Qty " +
                          "from currItems left join dbo.[MVXItemMaster] b on (currItems.ProjItem=b.ITEM) where b.BUShort is null";
            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(mqry, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }
        //------------------------------------------------------------------------
        // retrieve Item Warehouse Master Data from STAGING  
        //------------------------------------------------------------------------
        public static Dictionary<string, string> getItemWhseMaster()
        {
            SqlConnection conn = setUpMSSQLconn();
            //
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
        //
        public static Dictionary<string, string> getItemMaster()
        {
            SqlConnection conn = setUpMSSQLconn();
            //
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
        //
        public static Dictionary<string, string> getBUMaster()
        {
            SqlConnection conn = setUpMSSQLconn();
            //
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

        //-----------------------------------------------------------------------------
        // set up connection to SQL-EXPRESS - desktop at bella-vista / laptop at home
        //-----------------------------------------------------------------------------
        public static SqlConnection setUpMSSQLconn()
        {
            // create production connection object
            string strSQLsrvr = "SERVER=WETNT260;USER ID=#DSXdbadmin;PASSWORD=F0res!3R;" +
                    "DATABASE=STAGING;CONNECTION TIMEOUT=30;";
            // create test connection object
            //string strSQLsrvr = @"Data Source=9457F02\SQLEXPRESS;Initial Catalog=STAGING;Integrated Security=True";
            SqlConnection SqlConn = new SqlConnection(strSQLsrvr);
            return SqlConn;
        }

    }
}