using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ProjectManagementSuite.Models
{
    public class State
    {
        public string state { get; set; }
        public List<Details> ProjectDetails { get; set; }
    }

    public class Details
    {
        public string state { get; set; }
        public int projID { get; set; }
        public string projname { get; set; }
        public string whse { get; set; }
        public decimal pastdue { get; set; }
        public decimal current { get; set; }
    }
}