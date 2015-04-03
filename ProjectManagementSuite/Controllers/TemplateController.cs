using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using ProjectManagementSuite.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ProjectManagementSuite.Controllers
{
    public class TemplateController : ApiController
    {
        // POST api/<controller>
        [HttpPost]
        [Route("api/template")]
        public HttpResponseMessage Post([FromBody]List<newProject> newp)
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK);
            //
            var obj = new JObject();
            //
            // clear all previously generated templates from /GeneratedTemplates Folder
            string spath = HttpContext.Current.Server.MapPath("~/GeneratedTemplates");
            Array.ForEach(Directory.GetFiles(spath), File.Delete);
            //
            for (int pp = 0; pp < newp.Count; pp++)
            {
                // 
                string tabName = newp[pp].projectName.Replace(" ", string.Empty).Trim().Replace(",", string.Empty);  // compress name
                string fn0 = string.Format("{0}_{1}_{2}.xls", DateTime.Now.ToString("yyyyMMdd")                      // now date
                                                            , newp[pp].projectNumber.Trim()                          // project number
                                                            , tabName.Substring(0, Math.Min(26, tabName.Length)));   // compressed name
                // start of loop
                ProjectManagementSuite.CSharpLogic.GenerateWorkbook.generateProjectWorkBook(newp[pp], spath + "/" + fn0);
                // if there are movex orders then map the new orders to workbook
                if (newp[pp].mvxorders.Count != 0)
                {  // get the latest MOVEX order details instead of the stored lines
                    DataTable  dt = ProjectManagementSuite.CSharpLogic.ManageData.GetOpenMovexOrdersDetails(newp[pp]);
                    ProjectManagementSuite.CSharpLogic.UpdateWorkbook.updateProjectWorkBookLines(newp[pp], dt, spath + "/" + fn0);
                }
                // name key as file itself - value = path/file - avoid messy splitting of string on client side
                obj[fn0] = "GeneratedTemplates/" + fn0;
            }
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