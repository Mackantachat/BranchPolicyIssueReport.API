using BranchPolicyIssueReport.BusinessLogic;
using BranchPolicyIssueReport.DataContract;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApp2;


namespace BranchPolicyIssueReport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class BaseController : ControllerBase
    {
        protected readonly IConfiguration Configuration;
        protected readonly ServiceAction serviceAction;
        public BaseController(IConfiguration Configuration)
        {
            this.Configuration = Configuration;
            var connecttionString = Configuration.GetConnectionString("BLA");
            this.serviceAction = new ServiceAction(connecttionString);
        }
    }
}
