using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public  class PolicyListModel
    {
        public string Id { get; set; }
        public string POLICY { get; set; }
        public string APP_NO { get; set; }
        public string CERT_NO { get; set; }
        public string FLG_TYPE { get; set; }
        public string STATUS { get; set; }
        public string POLNO { get; set; }
        public string IBBL_APP_DATE { get; set; }
        public string IBBL_TRN_DATE { get; set; }
        public string KPI_BANK { get; set; }
        public string CSP_PAYIN_DATE { get; set; }
        public string CSP_UPD_DATE_TIME { get; set; }
        public string KPI_CSP { get; set; }
        public string APPSYS_DATE { get; set; }
        public string KPI_APPSYS { get; set; }
        public string AP { get; set; }
        public string CO { get; set; }
        public string MO { get; set; }
        public string IMM2_PSTS_DATE { get; set; }
        public string ISB_STS_DATE { get; set; }
        public string INSTALL_DT { get; set; }
        public string KPI_INSTALL { get; set; }
        public string SEND_DATE { get; set; }
        public string KPI_SEND { get; set; }
        public string KPI_TOTAL { get; set; }
        public string AMOUNT_UNCLEAN { get; set; }
        public string AMOUNT_DAY_UNCLEAN { get; set; }
        public string AMOUNT_CLEAN { get; set; }
        public string AMOUNT_DAY_CLEAN { get; set; }
        public string AMOUNT_CLEAN_PERCENT { get; set; }
        public string CLEAN_CASE_AVG { get; set; }
        public string AMOUNT_UNCLEAN_PERCENT { get; set; }
        public string UNCLEAN_CASE_AVG { get; set; }
        public string SUM_AMOUNT_TOTAL { get; set; }
        public string SUM_PERCENT_TOTAL { get; set; }
        public string SUM_TOTAL_AVG { get; set; }


    }
}
