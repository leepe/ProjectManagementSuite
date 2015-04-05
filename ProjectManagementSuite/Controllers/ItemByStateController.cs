using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace ProjectManagementSuite.Controllers
{
    public class ItemByStateController : ApiController
    {
        // GET api/controller - to trigger generation of issues workbook generation
        [HttpGet]
        [Route("api/itembystate")]
        public HttpResponseMessage Get()
        {
            var response = new HttpResponseMessage(HttpStatusCode.OK);
            //
            var obj = new JObject();
            //
            // clear all previously generated templates from /GeneratedTemplates Folder
            string spath = HttpContext.Current.Server.MapPath("~/GeneratedTemplates");
            Array.ForEach(Directory.GetFiles(spath), File.Delete);
            //----- name of log file ----------------------------------------------------------------------
            string fn0 = string.Format("ProductIssues_{0}.xls", DateTime.Now.ToString("yyyyMMdd"));                      // now date
            //---------------------------------------------------------------------------------------------
            // generate all the project item issues log
            //---------------------------------------------------------------------------------------------
            ProjectManagementSuite.CSharpLogic.UpdateWorkbook.generateProductIssuesLog(spath + "/" + fn0);
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
        // GET api/<controller>/<id> - Query which projects contain a particular item
        [HttpGet]
        [Route("api/itembystate/{id}")]
        public string Get(string id)
        {
            // replace underscores in product name with periods before query
            //string json = ProjectManagementSuite.CSharpLogic.GetProjectsByItemByState.GetJsonStringForItemByState(id.Replace('_','.'));
            string json = ProjectManagementSuite.CSharpLogic.GetProjectsByItemByState.GetJsonStringForItemByState(id);
            return json;
        }
    }
}