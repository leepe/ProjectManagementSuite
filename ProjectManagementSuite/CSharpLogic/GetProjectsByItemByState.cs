using Newtonsoft.Json;
using ProjectManagementSuite.Models;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;

namespace ProjectManagementSuite.CSharpLogic
{
    public class GetProjectsByItemByState
    {
        //---------------------------------------------------------------------------
        // pull out json dataset for particular item
        //---------------------------------------------------------------------------
        public static string GetJsonStringForItemByState(string item_)
        {
            //-------------------------------------------------
            // now get data to add to newly created datatables
            //-------------------------------------------------
            SqlConnection cnxn = ProjectManagementSuite.CSharpLogic.ManageData.setUpMSSQLconn(); 
            DataTable dt = GetProjIDplusDetailsForItem(cnxn, item_);
            //-------------------------------------------------
            // get the distinct states from dataset
            //-------------------------------------------------
            var states = (from g in dt.AsEnumerable() select new { States = g.Field<string>("state").Trim() }).Distinct();
            // loop through and load C# poco
            List<State> stList = new List<State>();
            //
            foreach (var k in states)
            {
                State st = new State();
                st.state = k.States;
                List<Details> listDet = new List<Details>();
                foreach (var g in dt.AsEnumerable().Where(d => d.Field<string>("state").Trim().Equals(k.States)))
                {
                    Details dtl = new Details();
                    dtl.projID = g.Field<int>("ProjID");
                    dtl.projname = g.Field<string>("ProjName").Trim().ToUpper();
                    dtl.whse = g.Field<string>("Whse").Trim();
                    dtl.pastdue = g.Field<decimal>("PastDue");
                    dtl.current = g.Field<decimal>("NextHrz");
                    listDet.Add(dtl);
                    st.ProjectDetails = listDet;
                }
                stList.Add(st);
            }
            JsonConvert.DefaultSettings = () => new JsonSerializerSettings
            {
                Formatting = Newtonsoft.Json.Formatting.Indented
            };
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(stList);
            //
            return json;
        }
        //---------------------------------------------------------------------------
        // retrieve state + project id + details from express - for Manual Forecasts
        //---------------------------------------------------------------------------
        private static DataTable GetProjIDplusDetailsForItem(SqlConnection conn, string product)
        {
            // modified 25/02/205 to compare first day of current month
            //
            string sql = "select distinct b.[State],b.ProjId,b.ProjName,a.[Whse]," +
                         "sum(case when a.[Month] < cast(convert(varchar(10),DATEADD(dd, -DAY(GETDATE()) + 1, GETDATE()),112) as integer) then a.[Qty] else 0 end) PastDue," +
                         "sum(case when a.[Month] >= cast(convert(varchar(10),DATEADD(dd, -DAY(GETDATE()) + 1, GETDATE()),112) as integer) then a.[Qty] else 0 end) NextHrz " +
                         "from dbo.[ProjectItems] a inner join dbo.[Projects] b on (a.ProjID=b.ProjID) " +
                         "where a.ProjItem='" + product + "' and b.ProjManFlag='Yes' " +
                         "group by b.[State],b.ProjId,b.ProjName,a.[Whse] " +
                         "order by b.[State],b.ProjId,a.[Whse] asc";

            conn.Open();
            SqlCommand SqlCmd = new SqlCommand(sql, conn);
            SqlDataAdapter objDataAdapter1 = new SqlDataAdapter();
            objDataAdapter1.SelectCommand = SqlCmd;
            DataSet objDataset1 = new DataSet();
            objDataAdapter1.Fill(objDataset1);
            conn.Close();
            return objDataset1.Tables[0];
        }

    }
}