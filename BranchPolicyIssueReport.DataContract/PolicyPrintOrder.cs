using ITUtility;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace BranchPolicyIssueReport.DataContract
{

    public class PolicyPrintOrder
    {
        private String _Policy;
        private DateTime _fSystemDt;
        private String _PolicyDocument;
        private String _SendOfc;
        private String _SendFlag;
        private String _SendID;
        private DateTime _ReceiveDt;
        private String _ReceiveID;
        private String _Process;
        private DateTime _ProcessDt;
        private String _ProcessID;
        private String _P_Name;
        private DateTime _BirthDate;
        private String _ID_CardNo;
        private String _OfficeName;
        private Int32 _PrintOrder;
        private String _Remarks;
        private DateTime _AgentReceiveDt;
        private String _SendRemarks;
        private String _ChequeDraft;


        public PolicyPrintOrder()
        {
            _Policy = null;
            _fSystemDt = new DateTime();
            _PolicyDocument = null;
            _SendOfc = null;
            _SendFlag = null;
            _SendID = null;
            _ReceiveDt = new DateTime();
            _ReceiveID = null;
            _Process = null;
            _ProcessDt = new DateTime();
            _ProcessID = null;
            _P_Name = null;
            _BirthDate = new DateTime();
            _ID_CardNo = null;
            _OfficeName = null;
            _PrintOrder = 0;
            _Remarks = null;
            _AgentReceiveDt = new DateTime();
            _SendRemarks = null;
            _ChequeDraft = null;
        }
        public PolicyPrintOrder(DataRow dr)
        {
            _Policy = Utility.GetDBStringValue(dr["POLICY"]);
            _fSystemDt = (DateTime)Utility.GetDBDateTimeValue(dr["FSYSTEMDATE"]);
            _PolicyDocument = Utility.GetDBStringValue(dr["POLICY_DOCUMENT"]);
            _SendOfc = Utility.GetDBStringValue(dr["SEND_OFC"]);
            _SendFlag = Utility.GetDBStringValue(dr["SEND_FLAG"]);
            _SendID = Utility.GetDBStringValue(dr["SEND_ID"]);
            _ReceiveDt = (DateTime)Utility.GetDBDateTimeValue(dr["RECEIVE_DT"]);
            _ReceiveID = Utility.GetDBStringValue(dr["RECEIVE_ID"]);
            _Process = Utility.GetDBStringValue(dr["PROCESS"]);
            _ProcessDt = (DateTime)Utility.GetDBDateTimeValue(dr["PROCESS_DT"]);
            _ProcessID = Utility.GetDBStringValue(dr["PROCESS_ID"]);
            _P_Name = Utility.GetDBStringValue(dr["P_NAME"]);
            _BirthDate = (DateTime)Utility.GetDBDateTimeValue(dr["BIRTH_DT"]);
            _ID_CardNo = Utility.GetDBStringValue(dr["ID_CARD_NO"]);
            _OfficeName = Utility.GetDBStringValue(dr["OFFICE_NAME"]);
            _PrintOrder = Convert.ToInt32(Utility.GetDBInt64Value(dr["PRINT_ORDER"]));
            _Remarks = Utility.GetDBStringValue(dr["REMARKS"]);
            _AgentReceiveDt = (DateTime)Utility.GetDBDateTimeValue(dr["AGENT_RECEIVE_DT"]);
            _SendRemarks = Utility.GetDBStringValue(dr["SEND_REMARKS"]);
            _ChequeDraft = Utility.GetDBStringValue(dr["CHEQUE_DRAFT"]);
        }

        public String Policy
        {
            get { return _Policy; }
            set { _Policy = value; }
        }
        public DateTime fSystemDt
        {
            get { return _fSystemDt; }
            set { _fSystemDt = value; }
        }
        public String PolicyDocument
        {
            get { return _PolicyDocument; }
            set { _PolicyDocument = value; }
        }
        public String SendOfc
        {
            get { return _SendOfc; }
            set { _SendOfc = value; }
        }
        public String SendFlag
        {
            get { return _SendFlag; }
            set { _SendFlag = value; }
        }
        public String SendID
        {
            get { return _SendID; }
            set { _SendID = value; }
        }
        public DateTime ReceiveDt
        {
            get { return _ReceiveDt; }
            set { _ReceiveDt = value; }
        }
        public String ReceiveID
        {
            get { return _ReceiveID; }
            set { _ReceiveID = value; }
        }
        public String Process
        {
            get { return _Process; }
            set { _Process = value; }
        }
        public DateTime ProcessDt
        {
            get { return _ProcessDt; }
            set { _ProcessDt = value; }
        }
        public String ProcessID
        {
            get { return _ProcessID; }
            set { _ProcessID = value; }
        }
        public String P_Name
        {
            get { return _P_Name; }
            set { _P_Name = value; }
        }
        public DateTime BirthDate
        {
            get { return _BirthDate; }
            set { _BirthDate = value; }
        }
        public String ID_CardNo
        {
            get { return _ID_CardNo; }
            set { _ID_CardNo = value; }
        }
        public String OfficeName
        {
            get { return _OfficeName; }
            set { _OfficeName = value; }
        }
        public Int32 PrintOrder
        {
            get { return _PrintOrder; }
            set { _PrintOrder = value; }
        }
        public String Remarks
        {
            get { return _Remarks; }
            set { _Remarks = value; }
        }
        public DateTime AgentReceiveDt
        {
            get { return _AgentReceiveDt; }
            set { _AgentReceiveDt = value; }
        }
        public String SendRemarks
        {
            get { return _SendRemarks; }
            set { _SendRemarks = value; }
        }
        public String ChequeDraft
        {
            get { return _ChequeDraft; }
            set { _ChequeDraft = value; }
        }
    }

}
