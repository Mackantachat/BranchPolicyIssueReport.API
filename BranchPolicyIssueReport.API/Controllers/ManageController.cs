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
    public class ManageController : BaseController
    {
        public ManageController(IConfiguration Configuration) : base(Configuration)
        {

        }

        [HttpGet]
        [Route("GetMonthShow")]
        public IActionResult GetMonthShow()
        {
            try
            {
                var dtmonth = manage.TableOfMonthShow();
                List<DropDownListModel> lstMonth = new List<DropDownListModel>();
                if (dtmonth.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (var row in dtmonth.Rows)
                    {
                        lstMonth.Add(new DropDownListModel
                        {
                            id = dtmonth.Rows[i].ItemArray[0].ToString(),
                            text = dtmonth.Rows[i].ItemArray[1].ToString()
                        });
                        i++;
                    }
                }

                return Ok(lstMonth);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }
        }

        [HttpGet]
        [Route("GetYearShow")]
        public IActionResult GetYearShow()
        {
            try
            {
                var dtYear = manage.TableOfYearShow();
                List<DropDownListModel> lstYear = new List<DropDownListModel>();
                if (dtYear.Rows.Count > 0)
                {
                    int i = 0;
                    foreach (var row in dtYear.Rows)
                    {
                        lstYear.Add(new DropDownListModel
                        {
                            id = dtYear.Rows[i].ItemArray[0].ToString(),
                            text = dtYear.Rows[i].ItemArray[0].ToString()
                        });
                        i++;
                    }
                }

                return Ok(lstYear);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }
        }

        [HttpPost]
        [Route("GetRegion")]
        public IActionResult GetRegion()
        {
            try
            {
                DropDownListModel[] data = serviceAction.GetRegion();

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetRegionSub")]
        public IActionResult GetRegionSub([FromBody]RequestPolicyOfMonthModel request)
        {
            try
            {
                DropDownListModel[] data = serviceAction.GetRegionSub(request.Section);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetBkkRegionDdl")]
        public IActionResult GetBkkRegionDdl([FromBody] RequestPolicyOfRegionModel sPolicy)
        {
            try
            {
                DropDownListModel[] data = serviceAction.GetBkkRegionDdl(sPolicy.Region);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

        [HttpPost]
        [Route("GetBkkRegion")]
        public IActionResult GetBkkRegion([FromBody] RequestPolicyOfDepModel sPolicy)
        {
            try
            {
                DropDownListModel[] data = serviceAction.GetBkkRegion(sPolicy.Office);

                return Ok(data);
            }
            catch (Exception ex)
            {
                return StatusCode(500, ex.Message.ToString());
            }

        }

    }
}
