using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class PolicyOfficeModel
    {
        public string Id { get; set; }
        public string POLICY_MONTH { get; set; }
        public string POLICY_MONTH_NAME { get; set; }
        public string POLICY_YEAR { get; set; }
        public string AMOUNT_CE { get; set; }
        public string SUMM_CE { get; set; }
        public string PREMIUM_CE { get; set; }
        public string AMOUNT_EA { get; set; }
        public string SUMM_EA { get; set; }
        public string PREMIUM_EA { get; set; }
        public string AMOUNT_NO { get; set; }
        public string SUMM_NO { get; set; }
        public string PREMIUM_NO { get; set; }
        public string AMOUNT_NE { get; set; }
        public string SUMM_NE { get; set; }
        public string PREMIUM_NE { get; set; }
        public string AmountTotal { get; set; }
        public string SummTotal { get; set; }
        public string PremiumTotal { get; set; }
    }
}
