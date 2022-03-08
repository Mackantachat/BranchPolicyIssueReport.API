using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class RequestPolicyModel
    {
        public string Month { get; set; }
        public string Year { get; set; }
    }

    public class RequestPolicyOfMonthModel : RequestPolicyModel
    {
        public string Section { get; set; }
    }

    public class RequestPolicyOfRegionModel : RequestPolicyOfMonthModel
    {
        public string Region { get; set; }

    }

    public class RequestPolicyOfDepModel : RequestPolicyOfRegionModel
    {
        public string Office { get; set; }
    }

    public class RequestPolicyListModel : RequestPolicyModel
    {
        public string StartInstallDT { get; set; }
        public string EndInstallDT { get; set; }
        public string Type { get; set; }
        public string Branch { get; set; }
        public string BancType { get; set; }
        public string PlanType { get; set; }
        public string Send { get; set; }
        public string Case { get; set; }
    }

    public class RequestPolicyDetail : RequestPolicyOfDepModel
    {
        public string Policy { get; set; }
    }
}
