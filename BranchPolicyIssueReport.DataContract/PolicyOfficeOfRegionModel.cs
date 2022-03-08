using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class PolicyOfficeOfRegionModel
    {
        public string Id { get; set; }
        public string SECTION_CODE { get; set; }
        public string SECTION_NAME { get; set; }
        public string REGION_CODE { get; set; }
        public string REGION_NAME { get; set; }
        public string NAME { get; set; }
        public string SURNAME { get; set; }
        public string SEQ { get; set; }
        public string AMOUNT_POLICY { get; set; }
        public string SUMM_POLICY { get; set; }
        public string PREMIUM_POLICY { get; set; }
    }
}
