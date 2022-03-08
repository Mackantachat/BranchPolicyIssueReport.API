using BranchPolicyIssueReport.DataContract;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Threading.Tasks;
using WebApp2;

namespace BranchPolicyIssueReport.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PolicyOfficeController : BaseController
    {
        public PolicyOfficeController(IConfiguration Configuration) : base(Configuration)
        {

        }

        [HttpPost]
        [Route("GetPolicyOffice")]
        public IActionResult GetPolicyOffice([FromBody]RequestPolicyModel sPolicy)
        {
            try
            {
                PolicyOfficeModel[] data = serviceAction.GetPolicyOffice(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }


        [HttpPost]
        [Route("GetPolicyOfficeOfMonth")]
        public IActionResult GetPolicyOfficeOfMonth([FromBody] RequestPolicyOfMonthModel sPolicy)
        {
            try
            {
                PolicyOfficeOfMonthModel[] data = serviceAction.GetPolicyOfficeOfMonth(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetPolicyOfficeOfRegion")]
        public IActionResult GetPolicyOfficeOfRegion([FromBody]RequestPolicyOfRegionModel sPolicy)
        {
            try
            {
                PolicyOfficeOfRegionModel[] data = serviceAction.GetPolicyOfficeOfRegion(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetPolicyOfficeOfDep")]
        public IActionResult GetPolicyOfficeOfDep([FromBody] RequestPolicyOfDepModel sPolicy)
        {
            try
            {
                PolicyOfficeOfDepModel[] data = serviceAction.GetPolicyOfficeOfDep(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetPolicyOfficeOfDate")]
        public IActionResult GetPolicyOfficeOfDate([FromBody] RequestPolicyOfDepModel sPolicy)
        {
            try
            {
                PolicyOfficeOfDateModel[] data = serviceAction.GetPolicyOfficeOfDate(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetPolicyList")]
        public IActionResult GetPolicyList([FromBody] RequestPolicyListModel sPolicy)
        {
            try
            {
                PolicyListModel[] data = serviceAction.GetPolicyList(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetPolicyDetail")]
        public IActionResult GetPolicyDetail([FromBody] RequestPolicyDetail sPolicy)
        {
            try
            {
                PolicyDetailModel data = serviceAction.GetPolicyDetail(sPolicy);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }
    }
}
