using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class PolicyOfficeOfDepModel
    {
        public string Id { get; set; }
        public string SUM_TB1_SEQ { get; set; }
        public string OFFICE { get; set; }
        public string DESCRIPTION { get; set; }
        public string AMOUNT_POLICY { get; set; }
        public string SUMM_POLICY { get; set; }
        public string PREMIUM_POLICY { get; set; }
        public string AMOUNT_SEND_POLICY { get; set; }
        public string AMOUNT_NOTSEND_POLICY { get; set; }
        public string AMOUNT_DAY_CLEAN { get; set; }
        public string AMOUNT_CLEAN { get; set; }
        public string CLEAN_PERCENT { get; set; }
        public string AMOUNT_DAY_UNCLEAN { get; set; }
        public string AMOUNT_UNCLEAN { get; set; }
        public string UNCLEAN_PERCENT { get; set; }
        public string TOTAL_AMOUNT_DAY { get; set; }
        public string TOTAL_AMOUNT { get; set; }
        public string SUM_AMOUNT_DAY_CLEAN { get; set; }
        public string SUM_AMOUNT_DAY_UNCLEAN { get; set; }
        public string ISO { get; set; }

    }
}
