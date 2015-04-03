using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectManagementSuite.Models
{
    // must use the List<mvxorders>  rather than List<string> to map JSON properly to class
    //
    public class newProject
    {
        public string projectNumber { get; set; }
        public string projectName { get; set; }
        public string projectState { get; set; }
        public string projectType { get; set; }
        public List<mvxorders> mvxorders { get; set; }
    }

    public class mvxorders
    {
        public string order { get; set; }
    }

}