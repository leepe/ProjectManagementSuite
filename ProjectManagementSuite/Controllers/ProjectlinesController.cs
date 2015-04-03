using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ProjectManagementSuite.Models;
using System;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ProjectManagementSuite.Controllers
{
    public class ProjectlinesController : ApiController
    {
        // 
        // GET api/projectlines/5 - Will ALWAYS generate A SINGLE download project...
        [HttpGet]
        [Route("api/projectlines")]
        public HttpResponseMessage Get(int id)
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK);
            //
            var obj = new JObject();
            //
            // clear all previously generated templates from /GeneratedTemplates Folder
            string spath = HttpContext.Current.Server.MapPath("~/GeneratedTemplates");
            Array.ForEach(Directory.GetFiles(spath), File.Delete);
            // produce a newProject object with which to produce workbook
            newProject oph = ProjectManagementSuite.CSharpLogic.ManageClientData.GetProjectHeaderForProjectLines(id);
            // pass newProject object to generate workbook
            string tabName = oph.projectName.Replace(" ", string.Empty).Trim().Replace(",", string.Empty);       // compress name
            string fn0 = string.Format("{0}_{1}_{2}.xls", DateTime.Now.ToString("yyyyMMdd")                      // now date
                                                        , oph.projectNumber.Trim()                               // project number
                                                        , tabName.Substring(0, Math.Min(26, tabName.Length)));   // compressed name
            // start of loop
            ProjectManagementSuite.CSharpLogic.GenerateWorkbook.generateProjectWorkBook(oph, spath + "/" + fn0);
            // now write out lines from database into workbook
            DataTable dt = new DataTable();
            if (oph.mvxorders.Count > 0)
            {   // get the latest MOVEX order details instead of the stored lines
                dt = ProjectManagementSuite.CSharpLogic.ManageData.GetOpenMovexOrdersDetails(oph);
            }
            else
            {   // a manual forecast therefore just get the lines from the db
                dt = ProjectManagementSuite.CSharpLogic.ManageData.GetProjectLinesForID(id);
            }
            ProjectManagementSuite.CSharpLogic.UpdateWorkbook.updateProjectWorkBookLines(oph,dt,spath + "/" + fn0);
            // name key as file itself - value = path/file - avoid messy splitting of string on client side
            obj[fn0] = "GeneratedTemplates/" + fn0;
            // prepare json file response
            var jdata = JsonConvert.SerializeObject(obj);
            response.Content = new StringContent("");
            response.Content = new StringContent(jdata);
            response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/json");
            // set up server response
            return response;
        }

    }
}