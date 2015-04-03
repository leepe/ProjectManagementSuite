using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectManagementSuite.Models
{
    public class ProjectSheets
    {
        // class constructor for project header details
        //
        public class projectHeader
        {
            public int projectID { get; set; }
            public string projectNumber { get; set; }
            public string projectName { get; set; }
            public string projectDesc { get; set; }
            public string projectFcst { get; set; }
            public string projectState { get; set; }
            public bool alreadyExists { get; set; }
            public List<string> mvxorders { get; set; }
            public projectHeader()
            {
                mvxorders = new List<string>();
            }
        }
        // class constructor for project body details
        //
        public class projectDetails
        {
            public int projectID { get; set; }
            public string projectItem { get; set; }
            public string projectWhse { get; set; }
            public int projectDate { get; set; }
            public int projectQty { get; set; }
        }
        // class constructor for project header + body
        //
        public class projectHeadPlusList
        {
            public projectHeader pHead { get; set; }
            public List<projectDetails> pList { get; set; }
        }
    }
}