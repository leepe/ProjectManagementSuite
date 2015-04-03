using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;

namespace ProjectManagementSuite.Controllers
{
    public class FileUploadController : ApiController
    {
        //POST api/fileupload 
        [HttpPost]
        [Route("api/fileupload")]
        public async Task<List<string>> PostAsync()
        {
            if (Request.Content.IsMimeMultipartContent())
            {
                string uploadPath = HttpContext.Current.Server.MapPath("~/uploads");
                //------------------------------------------------------------------
                // clear all previously generated templates from /uploads Folder
                //------------------------------------------------------------------
                Array.ForEach(Directory.GetFiles(uploadPath), File.Delete);
                //
                MyStreamProvider streamProvider = new MyStreamProvider(uploadPath);
                //------------------------------------------------------------------
                // wait for files to be uploaded into UPLOAD folder
                //------------------------------------------------------------------
                await Request.Content.ReadAsMultipartAsync(streamProvider);
                //
                //------------------------------------------------------------------
                // get and read project data from uploaded files
                //------------------------------------------------------------------
                // sqlconnection to express
                SqlConnection sqlcnxn = ProjectManagementSuite.CSharpLogic.ManageData.setUpMSSQLconn();
                // dictionary of existing projectNumber + projectName combinations
                Dictionary<string, string> dicProjNum = ProjectManagementSuite.CSharpLogic.ReadProjectSheets.getExistingProjectNumbers(sqlcnxn);
                // get uploaded files
                List<FileInfo> uf = ProjectManagementSuite.CSharpLogic.ReadProjectSheets.getUploadedFiles();
                // read excel files
                List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> lph
                       = ProjectManagementSuite.CSharpLogic.ReadProjectSheets.readUploadedExcelFiles(uf);
                // check that there are no duplication of sheets
                List<ProjectManagementSuite.Models.ProjectSheets.projectHeadPlusList> newFphl =
                                        ProjectManagementSuite.CSharpLogic.ReadProjectSheets.deleteOutDuplicateProjects(lph);
                // check that projectNumber + projectName combination doesn't already exist
                ProjectManagementSuite.CSharpLogic.ReadProjectSheets.probList(dicProjNum, newFphl);
                // add new records to database - don't use Entity framework at this stage....
                ProjectManagementSuite.CSharpLogic.ReadProjectSheets.addNewProjectsToDB(newFphl, sqlcnxn);
                //
                //----------------------------------------------------------------------
                // return messages that summarize the outcome of the upload
                //----------------------------------------------------------------------
                //
                List<string> messages = new List<string>();
                foreach (var g in newFphl)
                {
                    string newMess = "";
                    if (g.pHead.alreadyExists) newMess = string.Format("{0}-{1} already exists on DB"
                                                                       , g.pHead.projectNumber, g.pHead.projectName);
                    else if (!g.pHead.alreadyExists) newMess = string.Format("{0}-{1} added to DB"
                                                                       , g.pHead.projectNumber, g.pHead.projectName);
                    messages.Add(newMess);
                }

                return messages;
            }
            else
            {
                HttpResponseMessage response = Request.CreateResponse(HttpStatusCode.BadRequest, "Invalid Request!");
                throw new HttpResponseException(response);
            }
        }
    }

    public class MyStreamProvider : MultipartFormDataStreamProvider
    {
        public MyStreamProvider(string uploadPath)
            : base(uploadPath)
        {

        }

        public override string GetLocalFileName(HttpContentHeaders headers)
        {
            string fileName = headers.ContentDisposition.FileName;
            if (string.IsNullOrWhiteSpace(fileName))
            {
                fileName = Guid.NewGuid().ToString() + ".data";
            }
            return fileName.Replace("\"", string.Empty);
        }
    }
}