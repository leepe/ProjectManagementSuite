using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectManagementSuite.Models
{
    public class ProjectFcst
    {
        public int ProjID { get; set; }
        public string ProjNum { get; set; }
        public string ProjName { get; set; }
        public string ProjDesc { get; set; }
        public string State { get; set; }
        public string ProjManFlag { get; set; }
        public string MVXProjNum2 { get; set; }
    }
}