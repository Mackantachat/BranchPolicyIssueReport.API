using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class PolicyDetailModel
    {
        public CR_POLICY_REC CR_POLICY_REC { get; set; }
        public CR_PREMIUM_REC[] CR_PREMIUM_REC { get; set; }
        public Saleco Saleco { get; set; }
        public Agent Agent { get; set; }
    }
}
