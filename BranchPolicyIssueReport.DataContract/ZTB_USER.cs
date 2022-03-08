using System;
using System.Collections.Generic;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{
    public class ZTB_USER
    {


        public string USERID { get; set; }


        public string USERPWD { get; set; }


        public string PRENAME { get; set; }


        public string NAME { get; set; }


        public string SURNAME { get; set; }


        public string NICKNAME { get; set; }


        public DateTime? BIRTH_DT { get; set; }


        public string IDCARD_NO { get; set; }


        public string DEPARTMENT { get; set; }


        public string POSITION { get; set; }


        public string CLASS { get; set; }


        public string TEL_EXT { get; set; }


        public string EMAIL { get; set; }


        public DateTime? ENTRY_DT { get; set; }


        public char? DISABLE { get; set; }


        public DateTime? UPD_DT { get; set; }


        public string UPD_ID { get; set; }


        public string O_USERID { get; set; }


        public DateTime? TERMINATE_DT { get; set; }


        public char? APPROVE_FLG { get; set; }


        public string DIRECTOR { get; set; }


        public string N_USERID { get; set; }


        public string ACC_BANK { get; set; }


        public string ACC_NO { get; set; }


        public string MB_PHONE { get; set; }


        public string O_USERPWD { get; set; }


        public char? RECEIVE_PWD { get; set; }


        public string SECTION { get; set; }


    }

    public class BLAResponse<T>
    {
        public bool IsSuccess { get; set; }
        public string ErrorMessage { get; set; }
        public int ErrorCode { get; set; }
        public T Data { get; set; }

    }
}
