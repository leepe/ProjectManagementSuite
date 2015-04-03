using ProjectManagementSuite.CSharpLogic;
using ProjectManagementSuite.Models;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ProjectManagementSuite.Controllers
{
    
    public class ProjectsController : ApiController
    {
        //
        // GET api/projects
        [HttpGet]
        [Route("api/projects")]
        public IEnumerable<ProjectFcst> Get()
        {
            // return a json string of project data
            return ProjectManagementSuite.CSharpLogic.ManageData.GetProjectHeaderData();
        }

        // GET api/projects/5
        [HttpGet]
        [Route("api/projects/{id}")]
        public IEnumerable<ProjectFcst> Get(int id)
        {
            IEnumerable<ProjectFcst> pfc_get = ProjectManagementSuite.CSharpLogic.ManageClientData.GetProjectDataFromDatabase(id);
            return pfc_get;
        }

        // POST api/<controller>
        public void Post([FromBody]string value)
        {
        }

        // PUT api/<controller>/5
        [HttpPut]
        [Route("api/projects/{id}")]
        public void Put(int id, [FromBody]ProjectFcst value)
        {
            // updating database record
            ProjectManagementSuite.CSharpLogic.ManageClientData.UpdateProjectDataInDB(value);
        }

        // DELETE api/<controller>/5
        [HttpDelete]
        [Route("api/projects/{id}")]
        public HttpResponseMessage Delete(int id)
        {
            ManageClientData.DeleteProjectOutOfDatabase(id);
            var response = new HttpResponseMessage(HttpStatusCode.Accepted);
            return response;
        }
    }
}