using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class CR_PREMIUM_REC
    {
        public string Id { get; set; }
        public string PLAN_NAME { get; set; }
        public string SUMM { get; set; }
        public string PREMIUM { get; set; }
        public string XTR_PREMIUM { get; set; }
        public string OCP_CLASS { get; set; }
        public string MAT_DT { get; set; }
        public string LASTPAY_DT { get; set; }
        public string SumPremiumPlusXtr { get; set; }
    }
}
