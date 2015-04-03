using System;
using System.Web.Http;
using System.Web.Mvc;
using System.Web.Routing;
using ProjectManagementSuite.CSharpLogic;

namespace ProjectManagementSuite
{
    public class Global : System.Web.HttpApplication
    {

        protected void Application_Start(object sender, EventArgs e)
        {
            AreaRegistration.RegisterAllAreas();

            WebApiConfig.Register(GlobalConfiguration.Configuration);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
        }
 
    }
}