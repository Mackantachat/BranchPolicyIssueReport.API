using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using RestSharp;
using BranchPolicyIssueReport.DataContract;
using Newtonsoft.Json;

namespace BranchPolicyIssueReport.API.Controllers
{
    public class HomeController : Controller
    {
        protected readonly IConfiguration Configuration;
        public HomeController(IConfiguration Configuration)
        {
            this.Configuration = Configuration;
        }
       
        public async ValueTask<IActionResult> Index()
        {
            string
            N_USERID = string.Empty;
            string fullname = string.Empty;
            try
            {
                if (Request.Host.Host == "localhost")
                {
                    N_USERID = "760138";
                }
                else
                {
                    if ((Request.HasFormContentType && Request.Form != null &&
                    Request.Form.Any() &&
                    Request.Form["userID"].Any()))
                    {
                        N_USERID = Request.Form["userID"];
                    }

                    if (string.IsNullOrWhiteSpace(N_USERID))
                    {
                        return Redirect($"{Configuration["Login:BlaLogin"]}");
                    }

                }
                var user = GetUerInfo(N_USERID);
                ViewBag.link = Configuration["page:branchpolicyissuereport"];
                ViewBag.fullname = $"{N_USERID} {user.NAME} {user.SURNAME}";
                ViewBag.nuserid = user.N_USERID;
                return View();
            }
            catch (Exception e)
            {
                ViewBag.Errormessage = e.Message;
                return View();
            }
        }

        private ZTB_USER GetUerInfo(string N_USERID)
        {
         
            var client = new RestClient($"{Configuration["Link:DocumnetStorageAPI"]}");
            var request = new RestRequest($"api/v1/RecycleDocs/GetUserInfo/{N_USERID}", Method.GET);
            string errMsg = string.Empty;

            var responseApi = client.Execute<BLAResponse<ZTB_USER>>(request);
            if (responseApi.IsSuccessful && responseApi.Data.IsSuccess)
            {
                return responseApi.Data.Data;
            }
            else
            {
                if (responseApi.Data != null)
                {
                    throw new Exception(responseApi.Data.ErrorMessage);
                }
                else
                {
                    throw new Exception(responseApi.StatusCode.ToString());
                }
            }
        }
    }
}
