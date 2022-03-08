using System;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Data.OleDb;

/// <summary>
/// Summary description for Providers
/// </summary>
/// 
namespace WebApp2
{
    public class Providers
    {
        private string NL = "\n";

        private OleDbConnection _MyConn;
        private OleDbCommand oCommand;
        private OracleConnection oConn;
        private manage manage = new manage();
        public Providers()
        {
            _MyConn = null;
            oConn = null;
        }
        public Providers(OleDbConnection toConnection)
        {
            if (toConnection == null)
            {
                throw new Exception("Connection Error from ICProvider");
            }
            if (toConnection.State != ConnectionState.Open)
            {
                throw new Exception("Connection Error from ICProvider (Connection is not open.)");
            }
            _MyConn = toConnection;
        }
        public Providers(OracleConnection toConnection)
        {
            if (toConnection == null)
            {
                throw new Exception("Connection Error from ICProvider");
            }
            if (toConnection.State != ConnectionState.Open)
            {
                throw new Exception("Connection Error from ICProvider (Connection is not open.)");
            }
            oConn = toConnection;
        }
        public OleDbConnection Connection
        {
            get
            {
                return _MyConn;
            }
            set
            {
                _MyConn = value;
            }
        }
        public OracleConnection OracleConnection
        {
            get
            {
                return oConn;
            }
            set
            {
                oConn = value;
            }
        }
        private String _errMsg;
        public String errorMessage
        {
            get { return _errMsg; }
            set { _errMsg = value; }
        }
        public DataTable GetLogin(string vUserName, string vPassWord)
        {
            string sql = "SELECT U.USERID,U.N_USERID,U.NAME||' '||U.SURNAME AS NAME,U.DEPARTMENT " + NL +
                         "FROM ZTB_USER U " + NL +
                         "WHERE ( U.TERMINATE_DT IS NULL OR U.TERMINATE_DT > SYSDATE) " + NL +
                         "AND (U.N_USERID = '" + vUserName + "' OR U.USERID = '" + vUserName + "') " + NL;
                       //  "AND U.USERPWD = '" + vPassWord + "' ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn); 
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetOffice(string vUserName)
        {
            string sql = "SELECT O.OFFICE,DECODE(F.REGION,'BKK','A','B') AS FLG_TYPE " +
                         "FROM ZTB_USER_OFFICE O,ZTB_OFFICE F " +
                         "WHERE O.OFFICE = F.OFFICE " +
                         "AND O.USERID = '" + vUserName + "' " +
                         "AND O.EFF_DT = (SELECT MAX(O.EFF_DT) FROM ZTB_USER_OFFICE O WHERE O.USERID = '" + vUserName + "')";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppMain(string StrYear, string StrMonth, string ChkType)
        {
            manage manage = new manage();
            string myYear = StrYear.Substring(2, 2);
            //string CStartDt = manage.GetStrDate(StartDt);
            //string CEndDt = manage.GetStrDate(EndDt);
            string WhereType, WhereMonth;
            if (ChkType == "N")
            {
                WhereType = " AND M.FLG_DEP IN('C','D') ";
            }
            else
            {
                WhereType = " ";
            }
            if (StrMonth == "00")
            {
                WhereMonth = "";
            }
            else
            {
                WhereMonth = " AND SUBSTR(M.APP_DATE,3,2) = '" + StrMonth + "' ";
            }

            string sql = "SELECT COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " + NL +
                         ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " + NL +
                         ",M.FLG_DEP,DECODE(M.FLG_DEP,'A','ส่วนกรมธรรม์สามัญ 1','B','ส่วนกรมธรรม์สามัญ 2','C','สะสมเงินเดือน','D','Bancassurance','E','Tele Market') AS DEP " + NL +
                         "FROM WEB_APP_ALL_MAIN_NEW M " + NL +
                         "WHERE SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + NL +
                         WhereMonth + " " + WhereType + NL +
                         "GROUP BY M.FLG_DEP " + NL +
                         "ORDER BY M.FLG_DEP ASC ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppOffice(string StrYear, string StrMonth, string ChkType, string StartDt, string EndDt)
        {
            manage manage = new manage();
            string myYear = StrYear.Substring(2, 2);
            string CStartDt = "";
            string CEndDt = "";
            string WhereMonth = "";
            string WhereAppSysDate = "";
            if ((StartDt != "") && (EndDt != ""))
            {
                CStartDt = manage.GetStrDate(StartDt);
                CEndDt = manage.GetStrDate(EndDt);
                WhereAppSysDate = " AND M.APPSYS_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + NL;
            }
            else
            {
                WhereAppSysDate = "";
            }
            if (StrMonth == "00")
            {
                WhereMonth = "";
            }
            else
            {
                WhereMonth = " AND SUBSTR(M.APP_DATE,3,2) = '" + StrMonth + "' " + NL;
            }
            string sql = "SELECT DECODE(O.OFFICE,'สสง','สนญ',O.OFFICE) AS OFFICE,O.REGION,DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION) AS DESCRIPTION,M.FLG_DEP " + NL +
                         ",COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " + NL +
                         ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " + NL +
                         "FROM WEB_APP_ALL_MAIN_NEW M,ZTB_OFFICE O " + NL +
                         "WHERE M.OFFICE = O.OFFICE " + NL +
                         "AND M.FLG_DEP = '" + ChkType + "' " + NL +
                         "AND SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + WhereMonth + WhereAppSysDate + NL +
                         "GROUP BY DECODE(O.OFFICE,'สสง','สนญ',O.OFFICE),O.REGION,DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION),M.FLG_DEP " + NL +
                         "ORDER BY DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION) ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public SummaryOffice[] GetAppOfficeIsis(string StrYear, string StrMonth, string ChkType, string StartDt, string EndDt)
        {
            try
            {
                manage manage = new manage();
                string myYear = StrYear.Substring(2, 2);
                string CStartDt = "";
                string CEndDt = "";
                string WhereMonth = "";
                string WhereAppSysDate = "";
                if ((StartDt != "") && (EndDt != ""))
                {
                    CStartDt = manage.GetStrDate(StartDt);
                    CEndDt = manage.GetStrDate(EndDt);
                    WhereAppSysDate = " AND M.APPSYS_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + NL;
                }
                else
                {
                    WhereAppSysDate = "";
                }
                if (StrMonth == "00")
                {
                    WhereMonth = "";
                }
                else
                {
                    WhereMonth = " AND SUBSTR(M.APP_DATE,3,2) = '" + StrMonth + "' " + NL;
                }
                string sql = "SELECT DECODE(O.OFFICE,'สสง','สนญ',O.OFFICE) AS OFFICE,O.REGION,DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION) AS DESCRIPTION,M.FLG_DEP " + NL +
                             ",COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " + NL +
                             ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " + NL +
                             "FROM WEB_APP_ALL_MAIN_ISIS M,ZTB_OFFICE O " + NL +
                             "WHERE M.OFFICE = O.OFFICE " + NL +
                             "AND M.FLG_DEP = '" + ChkType + "' " + NL +
                             "AND SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + WhereMonth + WhereAppSysDate + NL +
                             "GROUP BY DECODE(O.OFFICE,'สสง','สนญ',O.OFFICE),O.REGION,DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION),M.FLG_DEP " + NL +
                             "ORDER BY DECODE(O.DESCRIPTION,'ส่วนสะสมเงินเดือน','สำนักงานใหญ่',O.DESCRIPTION) ASC";
                OracleCommand cmd = new OracleCommand(sql, oConn);
                OracleDataAdapter oDapt = new OracleDataAdapter(cmd);
                Array app;
                app = Utility.CreateArrayObject(oDapt, typeof(SummaryOffice));
                if (app != null)
                {
                    SummaryOffice[] returnValue = new SummaryOffice[app.Length];
                    Array.Copy(app, returnValue, app.Length);
                    return returnValue;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                _errMsg = ex.Message;
                return null;
            }

        }
        public DataTable GetAppForMonth(string StrYear, string StrMonth, string ChkType, string Office, string StartDt, string EndDt)
        {
            manage manage = new manage();
            string myYear = StrYear.Substring(2, 2);
            string CStartDt = "";
            string CEndDt = "";
            string sql = "";
            string whereType = "";
            string WhereMonth = "";
            string WhereAppSysDate = "";

            if ((StartDt != "") && (EndDt != ""))
            {
                CStartDt = manage.GetStrDate(StartDt);
                CEndDt = manage.GetStrDate(EndDt);
                WhereAppSysDate = " AND M.APPSYS_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " ";
            }
            else
            {
                WhereAppSysDate = "";
            }

            if (StrMonth == "00")
            {
                WhereMonth = "";
            }
            else
            {
                WhereMonth = " AND SUBSTR(M.APP_DATE,3,2) = '" + StrMonth + "' ";
            }
            if (Office == "สนญ")
            {
                whereType = "AND M.OFFICE IN('สนญ','สสง') ";
            }
            else
            {
                whereType = "AND M.OFFICE = '" + Office + "' ";
            }
            if ((ChkType == "C") || (ChkType == "D") || (ChkType == "E"))
            {
                sql = "SELECT SUBSTR(M.APP_DATE,3,2) AS MONTH_APP,M.FLG_DEP " +
                      ",'25'||SUBSTR(M.APP_DATE,0,2) AS YEAR_APP " +
                      ",COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " +
                      ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " +
                      "FROM WEB_APP_ALL_MAIN_NEW M " +
                      "WHERE SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + WhereMonth + WhereAppSysDate +
                      "AND M.FLG_DEP = '" + ChkType + "' " +
                      "GROUP BY SUBSTR(M.APP_DATE,3,2),M.FLG_DEP,SUBSTR(M.APP_DATE,0,2) " +
                      "ORDER BY SUBSTR(M.APP_DATE,0,2),SUBSTR(M.APP_DATE,3,2) ASC ";
            }
            else
            {
                if (Office == "")
                {
                    sql = "SELECT SUBSTR(M.APP_DATE,3,2) AS MONTH_APP " +
                          ",M.FLG_DEP " +
                          ",'25'||SUBSTR(M.APP_DATE,0,2) AS YEAR_APP " +
                          ",COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " +
                          ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " +
                          "FROM WEB_APP_ALL_MAIN_NEW M " +
                          "WHERE SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + WhereMonth + WhereAppSysDate +
                          "AND M.FLG_DEP = '" + ChkType + "' " +
                          "GROUP BY SUBSTR(M.APP_DATE,3,2) " +
                          ",M.FLG_DEP " +
                          ",SUBSTR(M.APP_DATE,0,2) " +
                          "ORDER BY SUBSTR(M.APP_DATE,0,2),SUBSTR(M.APP_DATE,3,2) ASC ";
                }
                else
                {
                    sql = "SELECT SUBSTR(M.APP_DATE,3,2) AS MONTH_APP,DECODE(M.OFFICE,'สสง','สนญ',M.OFFICE) AS OFFICE,M.FLG_DEP " +
                      ",'25'||SUBSTR(M.APP_DATE,0,2) AS YEAR_APP " +
                      ",COUNT(M.APP_NO) AS TOTAL_APP,SUM(M.SUMM) AS TOTAL_SUMM,SUM(M.PREMIUM) AS TOTAL_PREMIUM " +
                      ",SUM(M.APP_NO_M) AS TOTAL_APP_M,SUM(M.SUMM_M) AS TOTAL_SUMM_M,SUM(M.PREMIUM_M) AS TOTAL_PREMIUM_M " +
                      "FROM WEB_APP_ALL_MAIN_NEW M " +
                      "WHERE SUBSTR(M.APP_DATE,1,2) = '" + myYear + "' " + WhereMonth + WhereAppSysDate +
                      "AND M.FLG_DEP = '" + ChkType + "' " + whereType +
                      "GROUP BY SUBSTR(M.APP_DATE,3,2),DECODE(M.OFFICE,'สสง','สนญ',M.OFFICE),M.FLG_DEP,SUBSTR(M.APP_DATE,0,2) " +
                      "ORDER BY SUBSTR(M.APP_DATE,0,2),SUBSTR(M.APP_DATE,3,2) ASC ";
                }

            }

            //OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public string GetOfficeName(string myOffice)
        {
            string sql;
            sql = "SELECT O.REGION,O.DESCRIPTION " +
                  "FROM ZTB_OFFICE O " +
                  "WHERE O.OFFICE = '" + myOffice + "'";
            //OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            string Office = (string)dt.Rows[0]["DESCRIPTION"];
            return Office;
        }
        public DataTable GetAppForDaystring(string StartDt, string EndDt, string ChkType, string Office, string StartDtAppSys, string EndDtAppSys, string BancType, string PlanType, string UnderWrite)
        {
            manage manage = new manage();

            string CStartDt = manage.GetStrDate(StartDt);
            string CEndDt = manage.GetStrDate(EndDt);
            string CStartDtAppSys = "";
            string CEndDtAppSys = "";
            string sql = "";
            string whereType = "";
            string WhereAppSysDate = "";
            string WhereBancType = "";
            string WherePlanType = "";
            string WhereUnderWrite = "";

            if ((StartDtAppSys != "") && (EndDtAppSys != ""))
            {
                CStartDtAppSys = manage.GetStrDate(StartDtAppSys);
                CEndDtAppSys = manage.GetStrDate(EndDtAppSys);
                WhereAppSysDate = " AND A.APPSYS_DATE BETWEEN " + CStartDtAppSys + " AND " + CEndDtAppSys + " " + NL;
            }
            else
            {
                WhereAppSysDate = "";
            }

            if (Office == "สนญ")
            {
                whereType = "AND A.OFFICE IN('สนญ','สสง') " + NL;
            }
            else
            {
                whereType = "AND A.OFFICE = '" + Office + "' " + NL;
            }

            if (UnderWrite == "ALL")
            {
                WhereUnderWrite = "";
            }
            else
            {
                WhereUnderWrite = "AND A.UNDER_WRITE = '" + UnderWrite + "' ";
            }

            if ((ChkType == "C") || (ChkType == "E"))
            {
                sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "AS FLG_DEP " + NL +
                      ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                      ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                      "FROM WEB_APP_ALL_NEW A, IS_APP_REGION B " + NL +
                      "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                      "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                      "AND " + NL +
                      "( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      " ) = '" + ChkType + "' " + NL +
                      "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC ";

            }
            else if ((ChkType == "A") || (ChkType == "B"))
            {
                if (Office == "")
                {
                    sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN " + NL +
                          ",A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                          ",( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          "  ) " + NL +
                          "AS FLG_DEP " + NL +
                          ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                          ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                          ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                          "FROM WEB_APP_ALL_NEW A, IS_APP_REGION B " + NL +
                          "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                          "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                          "AND " + NL +
                          "( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          " ) = '" + ChkType + "' " + NL +
                          "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                          ",( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          "  ) " + NL +
                          "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC";
                }
                else
                {
                    sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",DECODE(A.OFFICE,'สสง','สนญ',A.OFFICE) AS OFFICE " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "AS FLG_DEP " + NL +
                      ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                      ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                      "FROM WEB_APP_ALL_NEW A, IS_APP_REGION B " + NL +
                      "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                      "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                      "AND " + NL +
                      "( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      " ) = '" + ChkType + "' " + whereType + " " + NL +
                      "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                      ",DECODE(A.OFFICE,'สสง','สนญ',A.OFFICE) " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC";
                }

            }
            else
            {
                if ((BancType == "1") && (PlanType == "G"))
                {
                    WhereBancType = " AND A.BANK_ASS = 'Y' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'G' " + NL;
                }
                else if ((BancType == "1") && (PlanType == "P"))
                {
                    WhereBancType = " AND A.BANK_ASS = 'Y' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'P' " + NL;
                }
                else if ((BancType == "1") && (PlanType == "E"))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'E' " + NL;
                }
                else if ((BancType == "1") && (PlanType == ""))
                {
                    WhereBancType = " AND (A.BANK_ASS = 'Y'  OR A.POLNO IN(SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.BANC_TYPE = 'E')) " + NL;
                    WherePlanType = " ";
                }
                else if ((BancType == "2") && (PlanType == "H"))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'H' " + NL;
                }
                else if ((BancType == "2") && (PlanType == "C"))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'C' " + NL;
                }
                else if ((BancType == "2") && (PlanType == ""))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' AND A.POLNO IN (SELECT  PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND  PLAN_BANC.BANC_TYPE IN('H','C')) " + NL;
                    WherePlanType = " ";
                }
                else if ((BancType == "") && (PlanType == ""))
                {
                    WhereBancType = " AND (A.BANK_ASS = 'Y' OR A.FLG_TYPE = 'C') " + NL;
                    WherePlanType = " ";
                }

                sql = "SELECT TB.KEY_IN,TB.APPSYS_DATE,TB.APP_DAY,TB.FLG_DEP " + NL +
                      ",COUNT(TB.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(TB.STATUS,'PO',1,0)) AS PO " + NL +
                      ",SUM(DECODE(TB.STATUS,'IF1',1,0)) AS IF1 " + NL +
                      ",SUM(DECODE(TB.STATUS,'MO',1,0)) AS MO " + NL +
                      ",SUM(DECODE(TB.STATUS,'AP',1,0)) AS AP " + NL +
                      ",SUM(DECODE(TB.STATUS,'CO',1,0)) AS CO " + NL +
                      ",SUM(DECODE(TB.STATUS,'NT',1,0)) AS NT " + NL +
                      ",SUM(DECODE(TB.STATUS,'EX',1,0)) AS EX " + NL +
                      ",SUM(DECODE(TB.STATUS,'CC',1,0)) AS CC " + NL +
                      ",SUM(DECODE(TB.STATUS,'PP',1,0)) AS PP " + NL +
                      ",SUM(DECODE(TB.STATUS,'DC',1,0)) AS DC " + NL +
                      ",SUM(DECODE(TB.STATUS,'IF2',1,0)) AS IF2 " + NL +
                      ",SUM(TB.AMOUNT_SEND) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(TB.STATUS,'IF2',1,0))- SUM(TB.AMOUNT_SEND)) AS NOT_SEND " + NL +
                      "FROM " + NL +
                      "( " + NL +
                      "SELECT 'A' AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",'D' AS FLG_DEP " + NL +
                      ",A.APP_NO " + NL +
                      ",A.POLNO " + NL +
                      ",A.FLG_TYPE,A.STATUS AS CHK_STATUS " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS IS NULL) THEN " + NL +
                      "        'PO' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS IS NULL) THEN " + NL +
                      "        'PO' " + NL +
                      "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF1' " + NL +
                      "      WHEN (A.FLG_TYPE = 'B') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF2' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF2' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'PC') THEN " + NL +
                      "        'IF1'" + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'AE') THEN " + NL +
                      "        'IF1'" + NL +
                      "      ELSE " + NL +
                      "        A.STATUS " + NL +
                      "    END " + NL +
                      " ) AS STATUS " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                      "        (SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO) " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                      "        (SELECT DECODE(COUNT(*),0,0,1)  FROM FSD_GM_MAST_SEND S,FSD_GM_APP F WHERE S.POLICY = F.POLICY AND S.CERT_NO = F.CERT_NO AND S.SEND_FAIL = 'N' AND TO_CHAR(F.APP_NO) = A.APP_NO AND F.POLICY = A.POLNO) " + NL +
                      "    END " + NL +
                      " ) AS AMOUNT_SEND " + NL +
                      "FROM WEB_APP_ALL_NEW A " + NL +
                      "WHERE A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + WhereBancType + WherePlanType + NL + WhereUnderWrite + NL +
                      ") TB " + NL +
                      "GROUP BY TB.KEY_IN,TB.APPSYS_DATE,TB.APP_DAY,TB.FLG_DEP " + NL +
                      "ORDER BY TB.APPSYS_DATE ASC";

            }
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppForDaystringIsis(string StartDt, string EndDt, string ChkType, string Office, string StartDtAppSys, string EndDtAppSys, string BancType, string PlanType, string UnderWrite)
        {
            manage manage = new manage();

            string CStartDt = manage.GetStrDate(StartDt);
            string CEndDt = manage.GetStrDate(EndDt);
            string CStartDtAppSys = "";
            string CEndDtAppSys = "";
            string sql = "";
            string whereType = "";
            string WhereAppSysDate = "";
            string WhereBancType = "";
            string WherePlanType = "";
            string WhereUnderWrite = "";

            if ((StartDtAppSys != "") && (EndDtAppSys != ""))
            {
                CStartDtAppSys = manage.GetStrDate(StartDtAppSys);
                CEndDtAppSys = manage.GetStrDate(EndDtAppSys);
                WhereAppSysDate = " AND A.APPSYS_DATE BETWEEN " + CStartDtAppSys + " AND " + CEndDtAppSys + " " + NL;
            }
            else
            {
                WhereAppSysDate = "";
            }

            if (Office == "สนญ")
            {
                whereType = "AND B.IAPR_ENT_OFC IN('สนญ','สสง') " + NL;
            }
            else
            {
                whereType = "AND B.IAPR_ENT_OFC = '" + Office + "' " + NL;
            }

            if (UnderWrite == "ALL")
            {
                WhereUnderWrite = "";
            }
            else
            {
                WhereUnderWrite = "AND A.UNDER_WRITE = '" + UnderWrite + "' ";
            }

            if ((ChkType == "C")||(ChkType == "E"))
            {
                sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "AS FLG_DEP " + NL +
                      ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                      ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                      "FROM WEB_APP_ALL_ISIS A, IS_APP_REGION B " + NL +
                      "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                      "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                      "AND " + NL +
                      "( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      " ) = '" + ChkType + "' " + NL +
                      "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC ";

            }
            else if ((ChkType == "A") || (ChkType == "B"))
            {
                if (Office == "")
                {
                    sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN " + NL +
                          ",A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                          ",( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          "  ) " + NL +
                          "AS FLG_DEP " + NL +
                          ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                          ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                          ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                          ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                          "FROM WEB_APP_ALL_ISIS A, IS_APP_REGION B " + NL +
                          "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                          "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                          "AND " + NL +
                          "( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          " ) = '" + ChkType + "' " + NL +
                          "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                          ",( " + NL +
                          "    CASE " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                          "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                          "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                          "    END " + NL +
                          "  ) " + NL +
                          "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC";
                }
                else
                {
                    sql = "SELECT DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",DECODE(B.IAPR_ENT_OFC,'สสง','สนญ',B.IAPR_ENT_OFC) AS OFFICE " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "AS FLG_DEP " + NL +
                      ",COUNT(A.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,NULL,1,0),0)) AS PO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'IF',1,0),0)) AS IF1 " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'MO',1,0),0)) AS MO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'AP',1,0),0)) AS AP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'A',DECODE(A.STATUS,'CO',1,0),0)) AS CO " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'NT',1,0),0)) AS NT " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'EX',1,0),0)) AS EX " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'CC',1,0),0)) AS CC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'PP',1,0),0)) AS PP " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'DC',1,0),0)) AS DC " + NL +
                      ",SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) AS IF2 " + NL +
                      ",SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO)) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(A.FLG_TYPE,'B',DECODE(A.STATUS,'IF',1,0),0)) - SUM((SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO))) AS NOT_SEND " + NL +
                      "FROM WEB_APP_ALL_ISIS A, IS_APP_REGION B " + NL +
                      "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                      "AND A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + NL + WhereUnderWrite + NL +
                      "AND " + NL +
                      "( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      " ) = '" + ChkType + "' " + whereType + " " + NL +
                      "GROUP BY DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) " + NL +
                      ",DECODE(B.IAPR_ENT_OFC,'สสง','สนญ',B.IAPR_ENT_OFC) " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                      "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000009') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000013') THEN 'D' " + NL +
                      "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO = '8000014') THEN 'D' " + NL +
                      "    END " + NL +
                      "  ) " + NL +
                      "ORDER BY A.APPSYS_DATE,DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) ASC";
                }

            }
            else
            {
                if ((BancType == "1") && (PlanType == "G"))
                {
                    WhereBancType = " AND A.BANK_ASS = 'Y' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'G' " + NL;
                }
                else if ((BancType == "1") && (PlanType == "P"))
                {
                    WhereBancType = " AND A.BANK_ASS = 'Y' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'P' " + NL;
                }
                else if ((BancType == "1") && (PlanType == ""))
                {
                    WhereBancType = " AND A.BANK_ASS = 'Y' " + NL;
                    WherePlanType = " ";
                }
                else if ((BancType == "2") && (PlanType == "H"))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'H' " + NL;
                }
                else if ((BancType == "2") && (PlanType == "C"))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " AND ( " + NL +
                                    "      CASE " + NL +
                                    "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                                    "      ELSE " + NL +
                                    "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                                    "      END " + NL +
                                    "      ) = 'C' " + NL;
                }
                else if ((BancType == "2") && (PlanType == ""))
                {
                    WhereBancType = " AND A.BANK_ASS IS NULL AND A.FLG_TYPE = 'C' " + NL;
                    WherePlanType = " ";
                }
                else if ((BancType == "") && (PlanType == ""))
                {
                    WhereBancType = " AND (A.BANK_ASS = 'Y' OR A.FLG_TYPE = 'C') " + NL;
                    WherePlanType = " ";
                }

                sql = "SELECT TB.KEY_IN,TB.APPSYS_DATE,TB.APP_DAY,TB.FLG_DEP " + NL +
                      ",COUNT(TB.APP_NO) AS APP_TOTAL " + NL +
                      ",SUM(DECODE(TB.STATUS,'PO',1,0)) AS PO " + NL +
                      ",SUM(DECODE(TB.STATUS,'IF1',1,0)) AS IF1 " + NL +
                      ",SUM(DECODE(TB.STATUS,'MO',1,0)) AS MO " + NL +
                      ",SUM(DECODE(TB.STATUS,'AP',1,0)) AS AP " + NL +
                      ",SUM(DECODE(TB.STATUS,'CO',1,0)) AS CO " + NL +
                      ",SUM(DECODE(TB.STATUS,'NT',1,0)) AS NT " + NL +
                      ",SUM(DECODE(TB.STATUS,'EX',1,0)) AS EX " + NL +
                      ",SUM(DECODE(TB.STATUS,'CC',1,0)) AS CC " + NL +
                      ",SUM(DECODE(TB.STATUS,'PP',1,0)) AS PP " + NL +
                      ",SUM(DECODE(TB.STATUS,'DC',1,0)) AS DC " + NL +
                      ",SUM(DECODE(TB.STATUS,'IF2',1,0)) AS IF2 " + NL +
                      ",SUM(TB.AMOUNT_SEND) AS AMOUNT_SEND " + NL +
                      ",(SUM(DECODE(TB.STATUS,'IF2',1,0))- SUM(TB.AMOUNT_SEND)) AS NOT_SEND " + NL +
                      "FROM " + NL +
                      "( " + NL +
                      "SELECT 'A' AS KEY_IN,A.APPSYS_DATE,SUBSTR(APPSYS_DATE,5,2) AS APP_DAY " + NL +
                      ",'D' AS FLG_DEP " + NL +
                      ",A.APP_NO " + NL +
                      ",A.POLNO " + NL +
                      ",A.FLG_TYPE,A.STATUS AS CHK_STATUS " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS IS NULL) THEN " + NL +
                      "        'PO' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS IS NULL) THEN " + NL +
                      "        'PO' " + NL +
                      "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF1' " + NL +
                      "      WHEN (A.FLG_TYPE = 'B') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF2' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'IF') THEN " + NL +
                      "        'IF2' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'PC') THEN " + NL +
                      "        'IF1' " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'AE') THEN " + NL +
                      "        'IF1' " + NL +
                      "      ELSE " + NL +
                      "        A.STATUS " + NL +
                      "    END " + NL +
                      " ) AS STATUS " + NL +
                      ",( " + NL +
                      "    CASE " + NL +
                      "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                      "        (SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO) " + NL +
                      "      WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                      "        (SELECT DECODE(COUNT(*),0,0,1)  FROM FSD_GM_MAST_SEND S,FSD_GM_APP F WHERE S.POLICY = F.POLICY AND S.CERT_NO = F.CERT_NO AND S.SEND_FAIL = 'N' AND TO_CHAR(F.APP_NO) = A.APP_NO AND F.POLICY = A.POLNO) " + NL +
                      "    END " + NL +
                      " ) AS AMOUNT_SEND " + NL +
                      "FROM WEB_APP_ALL_ISIS A " + NL +
                      "WHERE A.APP_DATE BETWEEN " + CStartDt + " AND " + CEndDt + " " + WhereAppSysDate + WhereBancType + WherePlanType + NL + WhereUnderWrite + NL +
                      ") TB " + NL +
                      "GROUP BY TB.KEY_IN,TB.APPSYS_DATE,TB.APP_DAY,TB.FLG_DEP " + NL +
                      "ORDER BY TB.APPSYS_DATE ASC";

            }
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppList(string myFlg, string myKey, string myStatus, string myType, string myDep, string myOffice, string StartAppSysDt, string EndAppSysDt, string StartAppDT, string EndAppDT, string Send, string Memo, string BancType, string PlanType, string UnderWrite)
        {
            manage manage = new manage();
            string sql = " ";
            string AppDateStart = " ";
            string AppDateEnd = " ";
            string AppSysDateStart = " ";
            string AppSysDateEnd = " ";
            string WhereAppDate = " ";
            string WhereAppSysDate = " ";
            string WhereStatus = " ";
            string WhereKey = " ";
            string WhereOffice = " ";
            string WhereSend = " ";
            string WhereBanc = " ";
            string WherePlan = " ";
            string WhereMemo = " ";
            string WhereUnderWrite = " ";

            if ((StartAppDT != "") && (EndAppDT != ""))
            {
                AppDateStart = manage.GetStrDate(StartAppDT);
                AppDateEnd = manage.GetStrDate(EndAppDT);
                WhereAppDate = " AND A.APP_DATE BETWEEN '" + AppDateStart + "' AND '" + AppDateEnd + "' ";
            }

            if ((StartAppSysDt != "") && (EndAppSysDt != ""))
            {
                AppSysDateStart = manage.GetStrDate(StartAppSysDt);
                AppSysDateEnd = manage.GetStrDate(EndAppSysDt);
                WhereAppSysDate = " AND A.APPSYS_DATE BETWEEN '" + AppSysDateStart + "' AND '" + AppSysDateEnd + "' ";
            }

            if (myStatus == "")
            {
                WhereStatus = " ";
            }
            else
            {
                WhereStatus = " AND" +
                              " ( " +
                              "    CASE " +
                              "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS IS NULL) THEN 'PO' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS IS NULL) THEN 'PO' " +
                              "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN 'IF1' " +
                              "      WHEN (A.FLG_TYPE = 'B') AND (A.STATUS = 'IF') THEN 'IF2' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'IF') THEN 'IF2' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'PC') THEN 'IF1' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'AE') THEN 'IF1' " +
                              "      ELSE A.STATUS " +
                              "    END " +
                              " ) = '" + myStatus + "' ";

            }

            if (myKey == "")
            {
                WhereKey = " ";
            }
            else
            {
                WhereKey = " AND DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) = '" + myKey + "' ";
            }

            if (myOffice == "")
            {
                WhereOffice = " ";
            }
            else if (myOffice == "สนญ")
            {
                WhereOffice = " AND A.OFFICE IN('สนญ','สสง') ";
            }
            else
            {
                WhereOffice = " AND A.OFFICE = '" + myOffice + "' ";
            }

            if (Send == "")
            {
                WhereSend = " ";
            }
            else
            {
                WhereSend = " AND " +
                            "( " +
                            "    CASE " +
                            "      WHEN (A.FLG_TYPE <> 'C') THEN " +
                            "        (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO) " +
                            "      ELSE " +
                            "        (SELECT DECODE(COUNT(*),0,'N','Y')  FROM FSD_GM_MAST_SEND S,FSD_GM_APP F WHERE S.POLICY = F.POLICY AND S.CERT_NO = F.CERT_NO AND S.SEND_FAIL = 'N' AND TO_CHAR(F.APP_NO) = A.APP_NO AND F.POLICY = A.POLNO) " +
                            "    END " +
                            " ) = '" + Send + "' ";
            }

            if (BancType == "")
            {
                WhereBanc = " ";
            }
            else if (BancType == "1")
            {
                WhereBanc = " AND (A.BANK_ASS = 'Y' OR A.POLNO IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.BANC_TYPE = 'E')) ";
            }
            else if (BancType == "2")
            {
                WhereBanc = " AND A.FLG_TYPE = 'C' AND A.POLNO IN(SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.BANC_TYPE IN ('C','H')) ";
            }

            if (PlanType == "")
            {
                WherePlan = " ";
            }
            else
            {
                WherePlan = " AND " +
                            " ( " +
                            "  CASE " +
                            "  WHEN (A.FLG_TYPE = 'C') THEN " +
                            "    (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " +
                            "  ELSE " +
                            "    (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " +
                            "  END " +
                            " ) = '" + PlanType + "' ";
            }

            if ((Memo == "") || (Memo == "ALL"))
            {
                WhereMemo = " ";
            }
            else
            {
                WhereMemo = " AND " +
                            "    ( " +
                            "     CASE " +
                            "     WHEN(A.FLG_TYPE <> 'C') THEN " +
                            "       WEB_PKG.GET_MEMOCODE(A.APP_NO) " +
                            "     ELSE " +
                            "       WEB_PKG.GET_MEMOBANCCODE(A.APP_NO,A.POLNO) " +
                            "     END " +
                            "    ) = '" + Memo + "' ";
            }

            if (UnderWrite == "ALL")
            {
                WhereUnderWrite = " ";
            }
            else
            {
                WhereUnderWrite = " AND A.UNDER_WRITE = '" + UnderWrite + "' ";
            }

            sql = "SELECT " + NL +
                  "/*( " + NL +
                  "    CASE " + NL +
                  "    WHEN ASCII(SUBSTR(A.APP_NO,0,1)) >= 65 THEN " + NL +
                  "        A.APP_NO " + NL +
                  "    ELSE " + NL +
                  "        LPAD(A.APP_NO,8,'0') " + NL +
                  "    END " + NL +
                  ") AS APP_NO */" + NL +
                  "A.APP_NO" + NL +
                  ",A.NAME,A.PLAN,A.SUMM,NVL(A.PREMIUM+A.PREM+A.RIDER,0) AS PREMIUM,A.CARD_NO " + NL +
                  ",A.APPSYS_DATE,A.APP_DATE " + NL +
                  "/*,( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') AND (A.STATUS = 'AP') AND (NVL(A.CALPREM,0) = 0) AND (A.BANK_ASS = 'Y') AND (A.CARD_NO IS NULL) THEN " + NL +
                  "        (SELECT IAB_DOC_OTHER FROM IS_APPL_BSC  WHERE IAB_APP_NO = A.APP_NO) " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        NVL(TRIM(A.CARD_NO||' '||(SELECT IAB_DOC_OTHER FROM IS_APPL_BSC  WHERE IAB_APP_NO = A.APP_NO)),(SELECT (SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 = C.CSP_OPTION) AS PAY_BY  FROM CS_SUSPENSE C  WHERE C.CSP_CANCEL IS NULL  AND  C.CSP_APP_NO =  A.APP_NO AND C.CSP_PAYIN_DATE = (SELECT MAX(C.CSP_PAYIN_DATE) FROM CS_SUSPENSE C  WHERE C.CSP_CANCEL IS NULL  AND  C.CSP_APP_NO = A.APP_NO) AND ROWNUM = 1)) " + NL +
                  "      ELSE " + NL +
                  "        NVL(A.CARD_NO,(SELECT (SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 =C.CFSP_OPTION) AS PAY_BY FROM CS_FSD_SP C WHERE C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = A.APP_NO AND C.CFSP_PAY_DT = (SELECT MAX(C.CFSP_PAY_DT) FROM  CS_FSD_SP C  WHERE C.CFSP_CANCEL IS NULL  AND C.CFSP_APP_NO = A.APP_NO)  AND ROWNUM = 1)) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_BY */" + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN  (A.FLG_TYPE <> 'C')THEN " + NL +
                  "            CASE " + NL +
                  "                WHEN TRIM((SELECT (SELECT DECODE(SS.CODE2,'08','EDC',SS.DESCRIPTION)  FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND ROWNUM = 1) AND ROWNUM = 1)) IS NOT NULL THEN " + NL +
                  "                     (SELECT (SELECT DECODE(SS.CODE2,'08','EDC','07','เงินสด','P7','เช็ค',SS.DESCRIPTION) FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 =S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' ))   AND ROWNUM = 1) AND ROWNUM = 1)" + NL +
                  "                ELSE" + NL +
                  "                    NVL(TRIM(A.CARD_NO),DECODE(A.BANK_ASS,'Y',NULL,'ยังไม่ได้ชำระ' ))" + NL +
                  "            END " + NL +
                  "        WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                  "            CASE " + NL +
                  "                WHEN TRIM((SELECT (SELECT DECODE(SS.CODE2,'08','EDC',SS.DESCRIPTION)  FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND ROWNUM = 1) AND ROWNUM = 1)) IS NOT NULL THEN " + NL +
                  "                     (SELECT (SELECT DECODE(SS.CODE2,'08','EDC','07','เงินสด','P7','เช็ค',SS.DESCRIPTION) FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 =S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO  and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' ))   AND ROWNUM = 1) AND ROWNUM = 1)" + NL +
                  "                ELSE" + NL +
                  "                    NVL(TRIM(A.CARD_NO),DECODE(A.BANK_ASS,'Y',NULL,'ยังไม่ได้ชำระ' ))" + NL +
                  "            END " + NL +
                  "    END" + NL +
                  ") " + NL +
                  "AS PAY_BY " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        A.STATUS_DT " + NL +
                  "      ELSE " + NL +
                  "        CASE " + NL +
                  "          WHEN (SELECT TO_NUMBER(TO_CHAR(C.STATUS_DT,'YY')) + 43||TO_CHAR(C.STATUS_DT,'MMDD') FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') IS NOT NULL THEN " + NL +
                  "            (SELECT TO_NUMBER(TO_CHAR(C.STATUS_DT,'YY')) + 43||TO_CHAR(C.STATUS_DT,'MMDD') FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') " + NL +
                  "          ELSE " + NL +
                  "            TO_NUMBER(TO_CHAR(NVL((SELECT MAX(M.LET_PRT) FROM FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLNO AND M.APP_NO = A.APP_NO AND M.LET_TYPE = A.STATUS),A.ENTRY_DT),'YY'))+43||TO_CHAR (NVL((SELECT MAX(M.LET_PRT) FROM FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLNO AND M.APP_NO = A.APP_NO AND M.LET_TYPE = A.STATUS),A.ENTRY_DT), 'MMDD')  " + NL +
                  "        END" + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS STATUS_DT " + NL +
                  "/*,( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE) FROM   CS_SUSPENSE CS WHERE  CS.CSP_CANCEL IS NULL AND CS.CSP_APP_NO = A.APP_NO) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT  MIN (SP.CFSP_PAY_DT) FROM   ZMIS.CS_FSD_SP SP WHERE SP.CFSP_CANCEL IS NULL AND SP.CFSP_APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_DT */" + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE)  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = A.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N'))) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE)  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = A.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N'))) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_DT " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(R.CSFY_WORK_DATE) FROM CS_SUSPENSE CS,CS_FY_RECEIVE R WHERE CS.CSP_SP_NO = R.CSFY_SP_NO AND CS.CSP_APP_NO = A.APP_NO AND CS.CSP_CANCEL IS NULL) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT MIN (SP.CFSP_WORK_DT)  FROM   ZMIS.CS_FSD_SP SP WHERE SP.CFSP_CANCEL IS NULL AND SP.CFSP_APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS WORK_DT " + NL +
                  ",(  " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN " + NL +
                  "      'ส่งเสนอผู้พิจารณา' " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      'กำลังพิจารณา '||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS  NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      'กำลังพิจารณา '||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    END " + NL +
                  " ) AS STATUS " + NL +
                  ",A.P_MODE,NVL(A.CALPREM,0) AS CALPREM,DECODE(B.IAPR_APP_NO,NULL,'A','B') AS KEY_IN " + NL +
                  ",A.FLG_TYPE,A.AGENT_CODE,A.UPLINE,A.OFFICE,A.STATUS AS STATUS_CODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_MEMODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_MEMOBANCDESC(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS STATUS_MEMO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_MEMOCODE(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_MEMOBANCCODE(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS MEMO_CODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_CODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        NVL(WEB_PKG.GET_COBANCDESC(A.APP_NO,A.POLNO),'ไม่ให้ความคุ้มครองเพิ่มเติมสำหรับการทุพพลภาพถาวรสิ้นเชิง') " + NL +
                  "    END " + NL +
                  " ) AS STATUS_CO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                  "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO IN(SELECT B.POLICY FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL)) THEN 'D' " + NL +
                  "    END " + NL +
                  " ) AS DEP " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.BANK_ASS = 'Y') AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT S.DESCRIPTION||' / Saleco:'||POLICY.WEB_PKG.GET_SALECONAME_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.BANK_ASS IS NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT S.DESCRIPTION||' / Saleco:'||POLICY.WEB_PKG.GET_SALECONAME_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.PL_BLOCK = 'S') THEN " + NL +
                  "      '' " + NL +
                  "    ELSE " + NL +
                  "      (SELECT O.DESCRIPTION||'('||O.OFFICE||')' AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = A.OFFICE) " + NL +
                  "    END " + NL +
                  " ) AS BRANCH " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.BANK_ASS = 'Y') AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT S.BBL_BRANCH  FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.BANK_ASS IS NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT S.BBL_BRANCH FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.PL_BLOCK = 'S') THEN " + NL +
                  "      '' " + NL +
                  "    ELSE " + NL +
                  "      '' " + NL +
                  "    END " + NL +
                  " ) AS BRANCH_CODE " + NL +
                  ",A.POLNO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SENDBY " + NL +
                  ",'วันที่นำส่งจากประกันชีวิต ' AS SEND_PLACE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DISTINCT S.SEND_DT FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT S.SEND_DT " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SEND_DATE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT P.STATUS FROM P_POLICY P WHERE P.POLICY = A.POLNO) " + NL +
                  "      ELSE " + NL +
                  "        (SELECT M.STATUS FROM FSD_GM_MAST M WHERE M.POLICY = A.POLNO  AND M.APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  " ) AS POL_STATUS " + NL +
                  "FROM WEB_APP_ALL_NEW A, IS_APP_REGION B " + NL +
                  "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                  "AND " + NL +
                  "(  " + NL +
                  "  CASE " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                  "  WHEN (A.FLG_TYPE = 'C') AND (A.POLNO IN(SELECT B.POLICY FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL)) THEN 'D' " + NL +
                  "  END " + NL +
                  ") = '" + myDep + "' " + NL +
                  "";

            sql = sql + WhereAppDate + WhereAppSysDate + WhereStatus + WhereKey + WhereOffice + WhereSend + WhereBanc + WherePlan + WhereMemo + WhereUnderWrite + " ORDER BY A.APPSYS_DATE ASC ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppListIsis(string myFlg, string myKey, string myStatus, string myType, string myDep, string myOffice, string StartAppSysDt, string EndAppSysDt, string StartAppDT, string EndAppDT, string Send, string Memo, string BancType, string PlanType, string UnderWrite)
        {
            manage manage = new manage();
            string sql = " ";
            string AppDateStart = " ";
            string AppDateEnd = " ";
            string AppSysDateStart = " ";
            string AppSysDateEnd = " ";
            string WhereAppDate = " ";
            string WhereAppSysDate = " ";
            string WhereStatus = " ";
            string WhereKey = " ";
            string WhereOffice = " ";
            string WhereSend = " ";
            string WhereBanc = " ";
            string WherePlan = " ";
            string WhereMemo = " ";
            string WhereUnderWrite = " ";

            if ((StartAppDT != "") && (EndAppDT != ""))
            {
                AppDateStart = manage.GetStrDate(StartAppDT);
                AppDateEnd = manage.GetStrDate(EndAppDT);
                WhereAppDate = " AND A.APP_DATE BETWEEN '" + AppDateStart + "' AND '" + AppDateEnd + "' ";
            }

            if ((StartAppSysDt != "") && (EndAppSysDt != ""))
            {
                AppSysDateStart = manage.GetStrDate(StartAppSysDt);
                AppSysDateEnd = manage.GetStrDate(EndAppSysDt);
                WhereAppSysDate = " AND A.APPSYS_DATE BETWEEN '" + AppSysDateStart + "' AND '" + AppSysDateEnd + "' ";
            }

            if (myStatus == "")
            {
                WhereStatus = " ";
            }
            else
            {
                WhereStatus = " AND" +
                              " ( " +
                              "    CASE " +
                              "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS IS NULL) THEN 'PO' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS IS NULL) THEN 'PO' " +
                              "      WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN 'IF1' " +
                              "      WHEN (A.FLG_TYPE = 'B') AND (A.STATUS = 'IF') THEN 'IF2' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'IF') THEN 'IF2' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'PC') THEN 'IF1' " +
                              "      WHEN (A.FLG_TYPE = 'C') AND (A.STATUS = 'AE') THEN 'IF1' " +
                              "      ELSE A.STATUS " +
                              "    END " +
                              " ) = '" + myStatus + "' ";

            }

            if (myKey == "")
            {
                WhereKey = " ";
            }
            else
            {
                WhereKey = " AND DECODE(B.IAPR_APP_NO,NULL,'A',DECODE(B.IAPR_ENT_OFC,'สนญ','A','สสง','A','B')) = '" + myKey + "' ";
            }

            if (myOffice == "")
            {
                WhereOffice = " ";
            }
            else if (myOffice == "สนญ")
            {
                WhereOffice = " AND B.IAPR_ENT_OFC IN('สนญ','สสง') ";
            }
            else
            {
                WhereOffice = " AND B.IAPR_ENT_OFC = '" + myOffice + "' ";
            }

            if (Send == "")
            {
                WhereSend = " ";
            }
            else
            {
                WhereSend = " AND " +
                            "( " +
                            "    CASE " +
                            "      WHEN (A.FLG_TYPE <> 'C') THEN " +
                            "        (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S WHERE S.SEND_FAIL IS NULL AND S.POLICY = A.POLNO) " +
                            "      ELSE " +
                            "        (SELECT DECODE(COUNT(*),0,'N','Y')  FROM FSD_GM_MAST_SEND S,FSD_GM_APP F WHERE S.POLICY = F.POLICY AND S.CERT_NO = F.CERT_NO AND S.SEND_FAIL = 'N' AND TO_CHAR(F.APP_NO) = A.APP_NO AND F.POLICY = A.POLNO) " +
                            "    END " +
                            " ) = '" + Send + "' ";
            }

            if (BancType == "")
            {
                WhereBanc = " ";
            }
            else if (BancType == "1")
            {
                WhereBanc = " AND A.BANK_ASS = 'Y' ";
            }
            else if (BancType == "2")
            {
                WhereBanc = " AND A.FLG_TYPE = 'C' ";
            }

            if (PlanType == "")
            {
                WherePlan = " ";
            }
            else
            {
                WherePlan = " AND " +
                            " ( " +
                            "  CASE " +
                            "  WHEN (A.FLG_TYPE = 'C') THEN " +
                            "    (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " +
                            "  ELSE " +
                            "    (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " +
                            "  END " +
                            " ) = '" + PlanType + "' ";
            }

            if ((Memo == "") || (Memo == "ALL"))
            {
                WhereMemo = " ";
            }
            else
            {
                WhereMemo = " AND " +
                            "    ( " +
                            "     CASE " +
                            "     WHEN(A.FLG_TYPE <> 'C') THEN " +
                            "       WEB_PKG.GET_MEMOCODE(A.APP_NO) " +
                            "     ELSE " +
                            "       WEB_PKG.GET_MEMOBANCCODE(A.APP_NO,A.POLNO) " +
                            "     END " +
                            "    ) = '" + Memo + "' ";
            }

            if (UnderWrite == "ALL")
            {
                WhereUnderWrite = " ";
            }
            else
            {
                WhereUnderWrite = " AND A.UNDER_WRITE = '" + UnderWrite + "' ";
            }

            sql = "SELECT " + NL +
                  "/*( " + NL +
                  "    CASE " + NL +
                  "    WHEN ASCII(SUBSTR(A.APP_NO,0,1)) >= 65 THEN " + NL +
                  "        A.APP_NO " + NL +
                  "    ELSE " + NL +
                  "        LPAD(A.APP_NO,8,'0') " + NL +
                  "    END " + NL +
                  ") AS APP_NO */" + NL +
                  "A.APP_NO" + NL +
                  ",A.NAME,A.PLAN,A.SUMM,NVL(A.PREMIUM+A.PREM+A.RIDER,0) AS PREMIUM,A.CARD_NO " + NL +
                  ",A.APPSYS_DATE,A.APP_DATE " + NL +
                  "/*,( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        NVL(TRIM(A.CARD_NO),(SELECT (SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 = C.CSP_OPTION) AS PAY_BY  FROM CS_SUSPENSE C  WHERE C.CSP_CANCEL IS NULL  AND  C.CSP_APP_NO =  A.APP_NO AND C.CSP_PAYIN_DATE = (SELECT MAX(C.CSP_PAYIN_DATE) FROM CS_SUSPENSE C  WHERE C.CSP_CANCEL IS NULL  AND  C.CSP_APP_NO = A.APP_NO) AND ROWNUM = 1)) " + NL +
                  "      ELSE " + NL +
                  "        NVL(TRIM(A.CARD_NO),(SELECT (SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 =C.CFSP_OPTION) AS PAY_BY FROM CS_FSD_SP C WHERE C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = A.APP_NO AND C.CFSP_PAY_DT = (SELECT MAX(C.CFSP_PAY_DT) FROM  CS_FSD_SP C  WHERE C.CFSP_CANCEL IS NULL  AND C.CFSP_APP_NO = A.APP_NO)  AND ROWNUM = 1)) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_BY */" + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN  (A.FLG_TYPE <> 'C')THEN " + NL +
                  "            CASE " + NL +
                  "                WHEN TRIM((SELECT (SELECT DECODE(SS.CODE2,'08','EDC',SS.DESCRIPTION)  FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND ROWNUM = 1) AND ROWNUM = 1)) IS NOT NULL THEN " + NL +
                  "                     (SELECT (SELECT DECODE(SS.CODE2,'08','EDC','07','เงินสด','P7','เช็ค',SS.DESCRIPTION) FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 =S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' ))   AND ROWNUM = 1) AND ROWNUM = 1)" + NL +
                  "                ELSE" + NL +
                  "                    NVL(TRIM(A.CARD_NO),DECODE(A.BANK_ASS,'Y',NULL,'ยังไม่ได้ชำระ' ))" + NL +
                  "            END " + NL +
                  "        WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                  "            CASE " + NL +
                  "                WHEN TRIM((SELECT (SELECT DECODE(SS.CODE2,'08','EDC',SS.DESCRIPTION)  FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND ROWNUM = 1) AND ROWNUM = 1)) IS NOT NULL THEN " + NL +
                  "                     (SELECT (SELECT DECODE(SS.CODE2,'08','EDC','07','เงินสด','P7','เช็ค',SS.DESCRIPTION) FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 =S.CSP_PAY_OPTION) AS PAY_BY  FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO    and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' )) AND S.CSP_PAYIN_DATE = (SELECT MAX(S.CSP_PAYIN_DATE) FROM CS_SUSPENSE_VIEW S  WHERE S.CSP_APP_NO = A.APP_NO   and ((S.csp_disabled = 'N') or (S.csp_disabled = 'Y' and S.csp_cancel_type= 'N' ))   AND ROWNUM = 1) AND ROWNUM = 1)" + NL +
                  "                ELSE" + NL +
                  "                    NVL(TRIM(A.CARD_NO),DECODE(A.BANK_ASS,'Y',NULL,'ยังไม่ได้ชำระ' ))" + NL +
                  "            END " + NL +
                  "    END" + NL +
                  ") " + NL +
                  "AS PAY_BY " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        A.STATUS_DT " + NL +
                  "      ELSE " + NL +
                  "        CASE " + NL +
                  "          WHEN (SELECT TO_NUMBER(TO_CHAR(C.STATUS_DT,'YY')) + 43||TO_CHAR(C.STATUS_DT,'MMDD') FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') IS NOT NULL THEN " + NL +
                  "            (SELECT TO_NUMBER(TO_CHAR(C.STATUS_DT,'YY')) + 43||TO_CHAR(C.STATUS_DT,'MMDD') FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') " + NL +
                  "          ELSE " + NL +
                  "            TO_NUMBER(TO_CHAR(NVL((SELECT MAX(M.LET_PRT) FROM FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLNO AND M.APP_NO = A.APP_NO AND M.LET_TYPE = A.STATUS),A.ENTRY_DT),'YY'))+43||TO_CHAR (NVL((SELECT MAX(M.LET_PRT) FROM FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLNO AND M.APP_NO = A.APP_NO AND M.LET_TYPE = A.STATUS),A.ENTRY_DT), 'MMDD')  " + NL +
                  "        END" + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS STATUS_DT " + NL +
                  "/*,( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE) FROM   CS_SUSPENSE CS WHERE  CS.CSP_CANCEL IS NULL AND CS.CSP_APP_NO = A.APP_NO) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT  MIN (SP.CFSP_PAY_DT) FROM   ZMIS.CS_FSD_SP SP WHERE SP.CFSP_CANCEL IS NULL AND SP.CFSP_APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_DT */" + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE)  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = A.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N'))) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT MIN(CS.CSP_PAYIN_DATE)  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = A.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N'))) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS PAY_DT " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "            (SELECT MIN(R.CSFY_WORK_DATE) FROM CS_SUSPENSE CS,CS_FY_RECEIVE R WHERE CS.CSP_SP_NO = R.CSFY_SP_NO AND CS.CSP_APP_NO = A.APP_NO AND CS.CSP_CANCEL IS NULL) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT MIN (SP.CFSP_WORK_DT)  FROM   ZMIS.CS_FSD_SP SP WHERE SP.CFSP_CANCEL IS NULL AND SP.CFSP_APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  ") " + NL +
                  "AS WORK_DT " + NL +
                  ",(  " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN " + NL +
                  "      'ส่งเสนอผู้พิจารณา' " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      'กำลังพิจารณา '||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS  NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      'กำลังพิจารณา '||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    END " + NL +
                  " ) AS STATUS " + NL +
                  ",A.P_MODE,NVL(A.CALPREM,0) AS CALPREM,DECODE(B.IAPR_APP_NO,NULL,'A','B') AS KEY_IN " + NL +
                  ",A.FLG_TYPE,A.AGENT_CODE,A.UPLINE,B.IAPR_ENT_OFC AS OFFICE,A.STATUS AS STATUS_CODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_MEMODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_MEMOBANCDESC(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS STATUS_MEMO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_MEMOCODE(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_MEMOBANCCODE(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS MEMO_CODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_CODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_COBANCDESC(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS STATUS_CO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                  "    WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                  "    WHEN (A.FLG_TYPE = 'C') AND (A.POLNO IN(SELECT B.POLICY FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL)) THEN 'D' " + NL +
                  "    END " + NL +
                  " ) AS DEP " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.BANK_ASS = 'Y') AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT S.DESCRIPTION||' / Saleco:'||POLICY.WEB_PKG.GET_SALECONAME_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.BANK_ASS IS NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT S.DESCRIPTION||' / Saleco:'||POLICY.WEB_PKG.GET_SALECONAME_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.PL_BLOCK = 'S') THEN " + NL +
                  "      '' " + NL +
                  "    ELSE " + NL +
                  "      (SELECT O.DESCRIPTION||'('||O.OFFICE||')' AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = B.IAPR_ENT_OFC) " + NL +
                  "    END " + NL +
                  " ) AS BRANCH " + NL +
                  ",A.POLNO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SENDBY " + NL +

                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.BANK_ASS = 'Y') AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT S.BBL_BRANCH  FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.BANK_ASS IS NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT S.BBL_BRANCH FROM GBBL_STRUCT S WHERE S.CLASS = 'B' AND S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.PL_BLOCK = 'S') THEN " + NL +
                  "      '' " + NL +
                  "    ELSE " + NL +
                  "      '' " + NL +
                  "    END " + NL +
                  " ) AS BRANCH_CODE " + NL +


                  ",'วันที่นำส่งจากประกันชีวิต ' AS SEND_PLACE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DISTINCT S.SEND_DT FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT S.SEND_DT " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SEND_DATE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT P.STATUS FROM P_POLICY P WHERE P.POLICY = A.POLNO) " + NL +
                  "      ELSE " + NL +
                  "        (SELECT M.STATUS FROM FSD_GM_MAST M WHERE M.POLICY = A.POLNO  AND M.APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  " ) AS POL_STATUS " + NL +
                  "FROM WEB_APP_ALL_ISIS A, IS_APP_REGION B " + NL +
                  "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL +
                  "AND " + NL +
                  "(  " + NL +
                  "  CASE " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                  "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                  "  WHEN (A.FLG_TYPE = 'C') AND (A.POLNO IN(SELECT B.POLICY FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL)) THEN 'D' " + NL +
                  "  END " + NL +
                  ") = '" + myDep + "' " + NL +
                  "";

            sql = sql + WhereAppDate + WhereAppSysDate + WhereStatus + WhereKey + WhereOffice + WhereSend + WhereBanc + WherePlan + WhereMemo + WhereUnderWrite + " ORDER BY A.APPSYS_DATE ASC ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetSaleco(string agentNo, string dep, string uplineNo)
        {
            string sql;
            if (dep == "D")
            {
                sql = "SELECT S.DESCRIPTION " +
                      ",WEB_PKG.GET_SALECO_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME " +
                      ",WEB_PKG.GET_MKG_BYAGENT(S.BBL_AGENTCODE) AS MKG_NAME " +
                      "FROM GBBL_STRUCT S " +
                      "WHERE S.BBL_AGENTCODE = '" + agentNo + "'";
            }
            else
            {
                sql = "SELECT NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = '" + agentNo + "'),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = '" + agentNo + "'))||' (สังกัด '||(SELECT O.DESCRIPTION AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = (NVL((SELECT OFFICE FROM GAG_DETAIL WHERE AGENTCODE = '" + agentNo + "'),(SELECT OFFICE FROM GAG_INACTIVE WHERE AGENTCODE = '" + agentNo + "' ))))||')' AS AGENT_NAME " +
                      ",NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = '" + uplineNo + "'),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = '" + uplineNo + "')) AS UPLINE_NAME " +
                      "FROM DUAL";
            }

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppDetail(string appNo)
        {


            OracleCommand com = new OracleCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_APP_WEB";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = oConn;
            com.Parameters.Clear();
            OracleParameter param = new OracleParameter("APP_REF", OracleDbType.RefCursor);
            param.Direction = ParameterDirection.Output;
            com.Parameters.Add(param);
            com.Parameters.Add("APPNO_IN", OracleDbType.Varchar2).Value = appNo;
            OracleDataAdapter da = new OracleDataAdapter(com);

            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppDescription(string type, string appNo)
        {
            OracleCommand com = new OracleCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_APPRIDER_WEB";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = oConn;
            com.Parameters.Clear();

            com.Parameters.Add("APPNO_IN", OracleDbType.Varchar2).Value = appNo;
            com.Parameters.Add("P_TYPE", OracleDbType.Varchar2).Value = type;
            OracleParameter param = new OracleParameter("APPRIDER_REF", OracleDbType.RefCursor);
            param.Direction = ParameterDirection.Output;
            com.Parameters.Add(param);
            OracleDataAdapter da = new OracleDataAdapter(com);

            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

            //OleDbDataAdapter da = new OleDbDataAdapter(com);
            //da.SelectCommand.Parameters.Add("APPNO_IN", OleDbType.VarChar).Value = appNo;
            //da.SelectCommand.Parameters.Add("P_TYPE", OleDbType.VarChar).Value = type;
            //DataTable dt = new DataTable();
            //da.Fill(dt);
            //return dt;
        }
        public string GetSendPolicy(string policy)
        {
            manage manage = new manage();
            String sendDT;
            string sql = "SELECT NVL(TO_CHAR(MAX(P_POLICY_SENDING.SEND_DT),'DD/MM/RRRR'),'-') AS SEND_DT " +
                         "FROM P_POLICY_SENDING  " +
                         "WHERE P_POLICY_SENDING.POLICY = LPAD('" + policy + "',8,'0')";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string myDate = (dt.Rows[0]["SEND_DT"]).ToString();
                if ((myDate == "-") || (myDate == null) || (myDate == ""))
                {
                    sendDT = "-";
                }
                else
                {
                    string[] buf = System.Text.RegularExpressions.Regex.Split(myDate, "/");
                    int yearEN = int.Parse(buf[2]) + 543;
                    sendDT = buf[0] + "/" + buf[1] + "/" + yearEN.ToString();
                }
            }
            else
            {
                sendDT = "-";
            }

            return sendDT;
        }
        public string GetSendPolicyPost(string policy)
        {
            manage manage = new manage();
            String sendDT;

            string sql = "SELECT DISTINCT MAX(TO_CHAR(D.DATEEND,'DD/MM/RRRR')) AS DATEEND " +
                         "FROM DMIMPORTDATA D,P_POLICY_SENDING S " +
                         "WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR " +
                         "AND S.POLICY = LPAD('" + policy + "',8,'0')";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                string myDate = (dt.Rows[0]["DATEEND"]).ToString();
                if ((myDate == "-") || (myDate == null) || (myDate == ""))
                {
                    sendDT = "-";
                }
                else
                {
                    string[] buf = System.Text.RegularExpressions.Regex.Split(myDate, "/");
                    int yearEN = int.Parse(buf[2]) + 543;
                    sendDT = buf[0] + "/" + buf[1] + "/" + yearEN.ToString();
                }
            }
            else
            {
                sendDT = "-";
            }

            return sendDT;
        }
        public string GetAge(string stDT, string endDT)
        {
            string sql = "SELECT POLICY.ORD_PRODUCT.FCL_AGE(TO_DATE('" + stDT + "','DD/MM/RRRR'),TO_DATE('" + endDT + "','DD/MM/RRRR'),TO_DATE('" + endDT + "','DD/MM/RRRR')) AS AGE " +
                         "FROM DUAL";
            string myAge = "";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                myAge = (dt.Rows[0]["AGE"]).ToString();
            }
            else
            {
                myAge = "";
            }
            return myAge;
        }
        public DataTable GetMain(string stYear)
        {
            string myYear = stYear.Substring(2, 2);
            String dateBegin = myYear + "0101";
            String dateEnd = myYear + "1231";
            string sql = "SELECT TB.APP_MONTH,TB.APP_YEAR " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'A',1,0)) AS AMOUNT_A " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'A',TB.SUMM,0)) AS SUMM_A " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'A',TB.PREMIUM,0)) AS PREMIUM_A " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'B',1,0)) AS AMOUNT_B " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'B',TB.SUMM,0)) AS SUMM_B " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'B',TB.PREMIUM,0)) AS PREMIUM_B " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'C',1,0)) AS AMOUNT_C " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'C',TB.SUMM,0)) AS SUMM_C " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'C',TB.PREMIUM,0)) AS PREMIUM_C " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'D',1,0)) AS AMOUNT_D " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'D',TB.SUMM,0)) AS SUMM_D " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'D',TB.PREMIUM,0)) AS PREMIUM_D " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'E',1,0)) AS AMOUNT_E " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'E',TB.SUMM,0)) AS SUMM_E " + NL +
                         ",SUM(DECODE(TB.FLG_DEP,'E',TB.PREMIUM,0)) AS PREMIUM_E " + NL +
                         "FROM " + NL +
                         "( " +
                         "SELECT  W.APP_DATE, " + NL +
                         "        W.OFFICE, " + NL +
                         "        W.APP_NO, " + NL +
                         "        W.SUMM, " +
                         "        (W.PREMIUM + W.PREM + W.RIDER) AS PREMIUM, " + NL +
                         "        ( " +
                         "          CASE " + NL +
                         "              WHEN (W.FLG_TYPE <> 'C') AND (W.BANK_ASS = 'Y') THEN " + NL +
                         "                  'D' " + NL +
                         "              WHEN (W.FLG_TYPE <> 'C') AND (W.PL_BLOCK = 'S') THEN " + NL +
                         "                  'C' " + NL +
                         "              WHEN (W.FLG_TYPE <> 'C') AND (W.MKG_TYPE = 'TEL') THEN " + NL +
                         "                  'E' " + NL +
                         "              WHEN (W.FLG_TYPE <> 'C')  AND (W.REGION = 'BKK') THEN " + NL +
                         "                  'A' " + NL +
                         "              WHEN (W.FLG_TYPE <> 'C') AND (W.REGION <> 'BKK') THEN " + NL +
                         "                  'B' " + NL +
                         "              WHEN (W.FLG_TYPE = 'C') THEN " + NL +
                         "                  'D' " + NL +
                         "          END " + NL +
                         "          ) AS FLG_DEP, " + NL +
                         "        SUBSTR(W.APP_DATE,1,2) AS APP_YEAR, " + NL +
                         "        SUBSTR(W.APP_DATE,3,2) AS APP_MONTH " + NL +
                         " FROM   WEB_APP_ALL_NEW W " + NL +
                         " WHERE  W.APP_DATE BETWEEN '" + dateBegin + "' AND '" + dateEnd + "' " + NL +
                         " /*WHERE  SUBSTR(W.APP_DATE,1,2) = '" + myYear + "' */" + NL +
                         " ) TB " + NL +
                         " GROUP BY TB.APP_MONTH,TB.APP_YEAR " + NL +
                         " ORDER BY TB.APP_MONTH ASC ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public string GetOfficeNameKey(string AppNo)
        {
            string StrOfficeName = "";
            string sql = "SELECT O.DESCRIPTION " +
                         "FROM IS_APP_REGION R,ZTB_OFFICE O " +
                         "WHERE R.IAPR_ENT_OFC = O.OFFICE " +
                         "AND R.IAPR_APP_NO = '" + AppNo + "'";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);

            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                StrOfficeName = (dt.Rows[0]["DESCRIPTION"]).ToString();
            }
            else
            {
                StrOfficeName = "สำนักงานใหญ่";
            }
            return StrOfficeName;
        }
        public DataTable GetMemoDetail(string AppNo)
        {
            string sql;
            //sql = "SELECT TB.APP_NO,TB.SEQ,TB.STATUS,TB.STS_DATE,TB.ORDERS,TB.UPD_ID " +
            //      ",TB.PEND_CODE,TB.PEND_DESC,TB.PSTS_DATE " +
            //      ",(SELECT R.ADMIN_DESC FROM ATB_PEND_REQM R WHERE R.PEND_CODE = TB.PEND_CODE) AS DESCRIPTION " +
            //      ",DECODE(TB.PEND_STATUS,'R','รับเอกสารแล้ว','C','ยกเลิก','F','กำลังพิจารณา','W','ยกเว้น','รอผลการตอบกลับ') AS PEND_STATUS " +
            //      "FROM " +
            //      "( " +
            //      "SELECT M.ILM2_APP_NO AS APP_NO " +
            //      ",M.ILM2_SEQ AS SEQ " +
            //      ",M.ILM2_STATUS AS STATUS " +
            //      ",M.ILM2_STS_DATE AS STS_DATE " +
            //      ",M.ILM2_ORDER AS ORDERS " +
            //      ",M.ILM2_UPD_ID AS UPD_ID " +
            //      ",M.ILM2_PEND_CODE AS PEND_CODE " +
            //      ",M.ILM2_PEND_DESC AS PEND_DESC " +
            //      ",M.ILM2_PEND_STATUS AS PEND_STATUS " +
            //      ",M.ILM2_PSTS_DATE AS PSTS_DATE " +
            //      "FROM IS_APPL_MO2 M " +
            //      "WHERE M.ILM2_APP_NO = '" + AppNo + "' " +
            //      "UNION ALL " +
            //      "SELECT M.IMM2_APP_NO AS APP_NO " +
            //      ",M.IMM2_SEQ AS SEQ " +
            //      ",M.IMM2_STATUS AS STATUS " +
            //      ",M.IMM2_STS_DATE AS STS_DATE " +
            //      ",M.IMM2_ORDER AS  ORDERS " +
            //      ",M.IMM2_UPD_ID AS UPD_ID " +
            //      ",M.IMM2_PEND_CODE AS PEND_CODE " +
            //      ",M.IMM2_PEND_DESC AS PEND_DESC " +
            //      ",M.IMM2_PEND_STATUS AS PEND_STATUS " +
            //      ",M.IMM2_PSTS_DATE AS PSTS_DATE " +
            //      "FROM IS_APPM_MO2 M " +
            //      "WHERE M.IMM2_APP_NO = '" + AppNo + "' " +
            //      ") TB " +
            //      "ORDER BY TB.SEQ ASC ";
            sql = "SELECT A.APP_NO,G.PRINT_SEQ AS SEQ,E.STATUS,H.STATUS_DT AS STS_DATE " + NL +
                  ",(SELECT LE.REFERENCE FROM U_LETTER_ID LE,W_SUMMARY_LETTER SL WHERE LE.ULETTER_ID = SL.ULETTER_ID AND LE.STATUS = H.STATUS AND LE.LETTER_DT = F.MEMO_TRN_DT AND LE.UAPP_ID = B.UAPP_ID ) AS ORDERS " + NL +
                  ",F.UPD_ID,G.PEND_CODE,NULL AS PEND_DESC,G.PEND_STATUS_DT AS PSTS_DATE ,(SELECT QM.DESCRIPTION FROM AUTB_PEND_REQM QM WHERE QM.PEND_CODE = G.PEND_CODE)||' '||I.PEND_DESCRIPTION AS DESCRIPTION " + NL +
                  ",(SELECT DD.DESCRIPTION FROM AUTB_DATADIC_DET DD WHERE DD.COL_NAME = 'PEND_STATUS' AND DD.CODE = G.PEND_STATUS) AS PEND_STATUS " + NL +
                  "FROM U_APPLICATION_ID A,U_APPLICATION B,W_UNDERWRITE_APPLICATION C,W_SUBUNDERWRITE_ID D,W_SUMMARY E,U_MEMO_ID F,U_MEMO_DETAIL G,U_STATUS_ID H,U_MEMO_DESCRIPTION I " + NL +
                  "WHERE A.APP_ID = B.APP_ID " + NL +
                  "AND A.APP_NO = '" + AppNo + "' " + NL +
                  "AND A.CHANNEL_TYPE = 'OD' " + NL +
                  "AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "AND C.UNDERWRITE_ID = D.UNDERWRITE_ID " + NL +
                  "AND D.CUSTOMER_TYPE = 'C' " + NL +
                  "AND D.SUBUNDERWRITE_ID = E.SUBUNDERWRITE_ID " + NL +
                  "AND E.TMN = 'N' " + NL +
                  "AND E.SUMMARY_ID = F.SUMMARY_ID " + NL +
                  "AND F.TMN = 'N' " + NL +
                  "AND F.UMEMO_ID = G.UMEMO_ID " + NL +
                  "AND B.UAPP_ID = H.UAPP_ID " + NL +
                  "AND H.TMN = 'N' " + NL +
                  "AND G.UMEMO_ID = I.UMEMO_ID(+) " + NL +
                  "AND G.PEND_CODE = I.PEND_CODE(+) " + NL +
                  "ORDER BY G.PRINT_SEQ ASC";

            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetCoDetail(string AppNo)
        {

            OracleCommand com = new OracleCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_CTOFE";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = oConn;
            com.Parameters.Clear();
            OracleParameter param = new OracleParameter("CTOFE_REF", OracleDbType.RefCursor);
            param.Direction = ParameterDirection.Output;
            com.Parameters.Add(param);
            com.Parameters.Add("APPNO_IN", OracleDbType.Varchar2).Value = AppNo;
            OracleDataAdapter da = new OracleDataAdapter(com);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;


        }
        public string GetBranchOfDep(string myDep, string myVal)
        {
            string sql;
            sql = "";
            if ((myDep == "A") || (myDep == "B"))
            {
                sql = "SELECT O.DESCRIPTION AS DESCRIPTION " +
                      "FROM ZTB_OFFICE O " +
                      "WHERE O.OFFICE = '" + myVal + "'";
            }
            else if (myDep == "D")
            {
                sql = "SELECT S.DESCRIPTION AS DESCRIPTION " +
                      "FROM GBBL_STRUCT S " +
                      "WHERE S.BBL_AGENTCODE = '" + myVal + "'";
            }
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            string Office = "";
            Office = (string)dt.Rows[0]["DESCRIPTION"];
            return Office;
        }
        public DataTable GetPolicySending(string myPolNo)
        {
            string sql = "SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') AS SENDBY " +
                         ",DECODE((SELECT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,'วันที่นำส่งจากประกันชีวิต ',(SELECT 'วันที่ส่งไปรษณีย์ ' FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) AS SEND_PLACE " +
                         ",DECODE((SELECT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,S.SEND_DT,(SELECT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) AS SEND_DATE " +
                         ",DECODE((SELECT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,'วันที่นำส่งจากประกันชีวิต '||TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),(SELECT 'วันที่ส่งไปรษณีย์ '||TO_CHAR(D.DATEEND,'DD/MM/RRRR') FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) AS SEND_DATE1 " +
                         "FROM P_POLICY_SENDING S " +
                         "WHERE S.POLICY = '" + myPolNo + "' " +
                         "AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = '" + myPolNo + "')";

            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public string GetMemoDesc(string AppNo)
        {
            string StrMemo = "";
            string sql = "SELECT C.DESCRIPTION " +
                         "FROM ZTB_CONSTANT3 C " +
                         "WHERE C.COL_NAME = 'MO_FLG' " +
                         "AND C.CODE2 = ( " +
                         "    SELECT A.MO_FLG " +
                         "    FROM ATB_PEND_REQM A " +
                         "    WHERE A.PEND_CODE = ( " +
                         "        SELECT TB1.PEND_CODE " +
                         "        FROM " +
                         "        ( " +
                         "            SELECT TB.PEND_CODE,TB.SEQ " +
                         "            FROM " +
                         "            ( " +
                         "                SELECT M.ILM2_PEND_CODE AS PEND_CODE,M.ILM2_SEQ AS SEQ " +
                         "                FROM IS_APPL_MO2 M " +
                         "                WHERE M.ILM2_APP_NO = SUBSTR('" + AppNo + "',2) " +
                         "                UNION " +
                         "                SELECT M.IMM2_PEND_CODE AS PEND_CODE,M.IMM2_SEQ AS SEQ " +
                         "                FROM IS_APPM_MO2 M " +
                         "                WHERE M.IMM2_APP_NO = SUBSTR('" + AppNo + "',2) " +
                         "            ) TB " +
                         "            ORDER BY TB.SEQ ASC " +
                         "        ) TB1 " +
                         "        WHERE ROWNUM = 1 " +
                         "    ) " +
                         ") ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            if (dt.Rows.Count > 0)
            {
                StrMemo = (dt.Rows[0]["DESCRIPTION"]).ToString();
            }
            else
            {
                StrMemo = "";
            }
            return StrMemo;
        }
        public DataTable GetPolicyMain(string startPolDT, string endPolDT, string monthVal)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string WhereMonth = "";
            if (monthVal == "00")
            {
                WhereMonth = " ";
            }
            else
            {
                WhereMonth = " AND TO_CHAR(P.POLICY_DT,'MM') = '" + monthVal + "' ";
            }
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";

            sql = "SELECT TO_CHAR(P.POLICY_DT,'YYYY') AS YEAR " +
                  ",TO_CHAR(P.POLICY_DT,'MM') AS MONTH " +
                  ",SUM(DECODE(P.POLICY_TYPE,'A',1,0)) AS AMOUNT_A " +
                  ",SUM(DECODE(P.POLICY_TYPE,'A',P.SUMM,0)) AS SUMM_A " +
                  ",SUM(DECODE(P.POLICY_TYPE,'A',(P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM),0)) AS PREMIUM_A " +
                  ",SUM(DECODE(P.POLICY_TYPE,'B',1,0)) AS AMOUNT_B " +
                  ",SUM(DECODE(P.POLICY_TYPE,'B',P.SUMM,0)) AS SUMM_B " +
                  ",SUM(DECODE(P.POLICY_TYPE,'B',(P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM),0)) AS PREMIUM_B " +
                  ",SUM(DECODE(P.POLICY_TYPE,'C',1,0)) AS AMOUNT_C " +
                  ",SUM(DECODE(P.POLICY_TYPE,'C',P.SUMM,0)) AS SUMM_C " +
                  ",SUM(DECODE(P.POLICY_TYPE,'C',(P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM),0)) AS PREMIUM_C " +
                  ",SUM(DECODE(P.POLICY_TYPE,'D',1,0)) AS AMOUNT_D " +
                  ",SUM(DECODE(P.POLICY_TYPE,'D',P.SUMM,0)) AS SUMM_D " +
                  ",SUM(DECODE(P.POLICY_TYPE,'D',(P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM),0)) AS PREMIUM_D " +
                  ",SUM(DECODE(P.POLICY_TYPE,'E',1,0)) AS AMOUNT_E " +
                  ",SUM(DECODE(P.POLICY_TYPE,'E',P.SUMM,0)) AS SUMM_E " +
                  ",SUM(DECODE(P.POLICY_TYPE,'E',(P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM),0)) AS PREMIUM_E " +
                  "FROM POLICY_ALL_MAIN_NEW P " +
                  "WHERE P.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " " + WhereMonth +
                  "GROUP BY TO_CHAR(P.POLICY_DT,'YYYY'),TO_CHAR(P.POLICY_DT,'MM') " +
                  "ORDER BY TO_CHAR(P.POLICY_DT,'MM') ASC";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public DataTable GetPolicyOffice(string startPolDT, string endPolDT, string monthVal)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string WhereMonth = "";
            if (monthVal == "00")
            {
                WhereMonth = " ";
            }
            else
            {
                WhereMonth = " AND TO_CHAR(B.POLICY_DT,'MM') = '" + monthVal + "' ";
            }
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";

            #region Comment by Korn
            //sql = "SELECT TO_CHAR(B.POLICY_DT,'MM') AS POLICY_MONTH,TO_CHAR(B.POLICY_DT,'YYYY') AS POLICY_YEAR \n" +
            //        ",SUM(DECODE(E.REGION_CODE,'CE',1,0)) AS AMOUNT_CE \n" +
            
            //        ",SUM(DECODE(E.REGION_CODE,'CE',A.SUMM,0)) AS SUMM_CE \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'CE',A.T_PREMIUM,0)) AS PREMIUM_CE \n" +
            //        ",SUM(DECODE(E.REGION_CODE,'EA',1,0)) AS AMOUNT_EA \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'EA',A.SUMM,0)) AS SUMM_EA \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'EA',A.T_PREMIUM,0)) AS PREMIUM_EA \n" +
            //        ",SUM(DECODE(E.REGION_CODE,'NO',1,0)) AS AMOUNT_NO \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'NO',A.SUMM,0)) AS SUMM_NO \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'NO',A.T_PREMIUM,0)) AS PREMIUM_NO \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'NE',1,0)) AS AMOUNT_NE \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'NE',A.SUMM,0)) AS SUMM_NE \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'NE',A.T_PREMIUM,0)) AS PREMIUM_NE \n" +
            //        ",SUM(DECODE(E.REGION_CODE,'SO',1,0)) AS AMOUNT_SO \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'SO',A.SUMM,0)) AS SUMM_SO \n" +

            //        ",SUM(DECODE(E.REGION_CODE,'SO',A.T_PREMIUM,0)) AS PREMIUM_SO \n" +
            //      " FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE D,ZTB_BLA_REGION E \n" +
            //      " WHERE A.CHANNEL_TYPE = 'OD' \n" +
            //      " AND A.PL_BLOCK = 'A' \n" +
            //      " AND C.MARKETING_TYPE <> 'TEL' \n" +
            //      " AND A.POLICY_ID = B.POLICY_ID \n" +
            //      " AND A.POLICY_ID = C.POLICY_ID \n" +
            //      " AND B.APP_OFC = D.OFFICE \n" +
            //      " AND D.REGION_CODE = E.REGION_CODE \n" +
            //      " AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " " + WhereMonth + "\n" +
            //      " GROUP BY TO_CHAR(B.POLICY_DT,'MM'),TO_CHAR(B.POLICY_DT,'YYYY') \n" +
            //      " ORDER BY TO_CHAR(B.POLICY_DT,'MM') ASC \n";
            #endregion

            // add by Korn
            sql = "SELECT TO_CHAR(B.POLICY_DT,'MM') AS POLICY_MONTH,TO_CHAR(B.POLICY_DT,'YYYY') AS POLICY_YEAR \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SM',1,0))                AS AMOUNT_CE \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SM',A.SUMM,0))           AS SUMM_CE \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SM',A.T_PREMIUM,0))      AS PREMIUM_CE \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SS',1,0))                AS AMOUNT_EA \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SS',A.SUMM,0))           AS SUMM_EA \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SS',A.T_PREMIUM,0))      AS PREMIUM_EA \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SN',1,0))                AS AMOUNT_NO \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SN',A.SUMM,0))           AS SUMM_NO \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SN',A.T_PREMIUM,0))      AS PREMIUM_NO \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SE',1,0))                AS AMOUNT_NE \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SE',A.SUMM,0))           AS SUMM_NE \n" +
                    ",SUM(DECODE(S.SECTION_CODE,'SE',A.T_PREMIUM,0))      AS PREMIUM_NE \n" +

                  " FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE2 D,ZTB_BLA_REGION2 E ,ZTB_BLA_SECTION S \n" +
                  " WHERE A.CHANNEL_TYPE = 'OD' \n" +
                  " AND A.PL_BLOCK = 'A' \n" +
                  " AND C.MARKETING_TYPE <> 'TEL' \n" +
                  " AND A.POLICY_ID = B.POLICY_ID \n" +
                  " AND A.POLICY_ID = C.POLICY_ID \n" +
                  " AND B.APP_OFC = D.OFFICE \n" +
                  " AND D.REGION_CODE = E.REGION_CODE \n" +
                  " AND E.SECTION_CODE = S.SECTION_CODE \n" +
                  " AND D.REGION_CODE = E.REGION_CODE \n" +
                  //" AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " " + WhereMonth + "\n" +
                  " AND B.POLICY_DT  BETWEEN TO_DATE('" + StrStartPolDT + " 00:00:00','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE('" + StrEndPolDT + " 23:59:59','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''')" + " " + WhereMonth + "\n" + 
                  " GROUP BY TO_CHAR(B.POLICY_DT,'MM'),TO_CHAR(B.POLICY_DT,'YYYY') \n" +
                  " ORDER BY TO_CHAR(B.POLICY_DT,'MM') ASC \n";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyOfficeOfMonth(string txtregion, string startPolDT, string endPolDT)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string condition = "";
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);

            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";

            string test = txtregion.ToString();

            if (txtregion != "")
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                //condition = condition + "D.REGION_CODE = " + Utility.SQLValueString(txtregion) + "\n"; //comment by Korn 
                condition = condition + "S.SECTION_CODE = " + Utility.SQLValueString(txtregion) + "\n"; //add by Korn
            }

            #region comment by Korn
            //sql = "SELECT E.REGION_CODE,E.REGION_NAME,G.NAME,G.SURNAME,E.SEQ \n " +
            //" ,COUNT(A.POLICY_ID) AS AMOUNT_POLICY \n " +
            //" ,SUM(A.SUMM) AS SUMM_POLICY \n " +
            //" ,SUM(A.T_PREMIUM) AS PREMIUM_POLICY \n " +
            //" FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE D \n " +
            //" ,ZTB_BLA_REGION E,ZTB_BLA_REGION_USER F,ZTB_USER G \n " +
            //" WHERE A.CHANNEL_TYPE = 'OD' \n " +
            //" AND A.PL_BLOCK = 'A' \n " +
            //" AND A.POLICY_ID = B.POLICY_ID \n " +
            //" AND A.POLICY_ID = C.POLICY_ID \n " +
            //" AND B.APP_OFC = D.OFFICE \n " +
            //" AND D.REGION_CODE = E.REGION_CODE \n " +
            //" AND E.REGION_CODE = F.REGION_CODE \n " +
            //" AND F.N_USERID = G.N_USERID \n " +
            //" AND C.MARKETING_TYPE <> 'TEL' \n " +
            //" AND B.POLICY_DT BETWEEN " + StrStartPolDTO + "\n" +
            //" AND " + StrEndPolDTO + " " + condition + "\n" +
            //" GROUP BY E.REGION_CODE,E.REGION_NAME,G.NAME,G.SURNAME,E.SEQ \n" +
            //" Order by E.SEQ ASC ";
            #endregion

            sql = "SELECT S.SECTION_CODE \n" +
                    " ,S.SECTION_NAME \n" +
                    " ,G.NAME,G.SURNAME \n" +
                    " ,S.SEQ \n " +
                    " ,COUNT(A.POLICY_ID) AS AMOUNT_POLICY \n " +
                    " ,SUM(A.SUMM) AS SUMM_POLICY \n " +
                    " ,SUM(A.T_PREMIUM) AS PREMIUM_POLICY \n " +
                 " FROM  PW_LIFE_SUMM A \n" +
                      " ,P_INSTALL B \n" +
                      " ,P_AGENT C \n" +
                      " ,ZTB_BLA_REGION_OFFICE2 D \n" +
                      " ,ZTB_BLA_REGION2 E \n" +
                      " ,ZTB_BLA_REGION_ID F \n" +
                      " ,ZTB_USER G  \n" +
                      " ,ZTB_BLA_SECTION S \n" +
                 " WHERE A.CHANNEL_TYPE = 'OD' \n " +
                    " AND A.PL_BLOCK = 'A' \n " +
                    " AND A.POLICY_ID = B.POLICY_ID \n " +
                    " AND A.POLICY_ID = C.POLICY_ID \n " +
                    " AND B.APP_OFC = D.OFFICE \n " +
                    " AND D.REGION_CODE = E.REGION_CODE \n" +
                    " AND E.SECTION_CODE = S.SECTION_CODE \n" +
                    " AND F.SECTION_CODE = S.SECTION_CODE \n" +
                    " AND F.REGION_CODE is null  \n" +
                    " AND F.N_USERID <> '000199' \n" +
                    " AND F.N_USERID <> '000198' \n" +
                    " AND F.N_USERID <> '001541' \n" +
                    " AND F.N_USERID = G.N_USERID \n " +
                    " AND C.MARKETING_TYPE <> 'TEL' \n " +
                    //" AND B.POLICY_DT BETWEEN " + StrStartPolDTO + "\n" +
                    //" AND " + StrEndPolDTO + " " + "condition + "\n" +
                    " AND B.POLICY_DT  BETWEEN TO_DATE('" + StrStartPolDT + " 00:00:00','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE('" + StrEndPolDT + " 23:59:59','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''')" + " " + condition + "\n" + 
                    " GROUP BY S.SECTION_CODE,S.SECTION_NAME,G.NAME,G.SURNAME,S.SEQ \n" +
                    " Order by S.SEQ ASC ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyOfficeOfRegion(string txtsection, string txtregion, string startPolDT, string endPolDT) //add by Korn
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string condition = "";
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);

            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";

            string test = txtregion.ToString();

            if (txtsection != "")
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                condition = condition + "S.SECTION_CODE = " + Utility.SQLValueString(txtsection) + "\n";
            }


            if (txtregion != "")
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                condition = condition + "E.REGION_CODE = " + Utility.SQLValueString(txtregion) + "\n"; 
            }

            sql = "SELECT S.SECTION_CODE \n" +
                    " ,S.SECTION_NAME \n" +
                    " ,E.REGION_CODE \n"+
                    " ,E.REGION_NAME \n "+
                    " ,G.NAME,G.SURNAME \n" +
                    " ,E.SEQ \n " +
                    " ,COUNT(A.POLICY_ID) AS AMOUNT_POLICY \n " +
                    " ,SUM(A.SUMM) AS SUMM_POLICY \n " +
                    " ,SUM(A.T_PREMIUM) AS PREMIUM_POLICY \n " +
                 " FROM  PW_LIFE_SUMM A \n" +
                      " ,P_INSTALL B \n" +
                      " ,P_AGENT C \n" +
                      " ,ZTB_BLA_REGION_OFFICE2 D \n" +
                      " ,ZTB_BLA_REGION2 E \n" +
                      " ,ZTB_BLA_REGION_ID F \n" +
                      " ,ZTB_USER G  \n" +
                      " ,ZTB_BLA_SECTION S \n" +
                 " WHERE A.CHANNEL_TYPE = 'OD' \n " +
                    " AND A.PL_BLOCK = 'A' \n " +
                    " AND A.POLICY_ID = B.POLICY_ID \n " +
                    " AND A.POLICY_ID = C.POLICY_ID \n " +
                    " AND B.APP_OFC = D.OFFICE \n " +
                    " AND D.REGION_CODE = E.REGION_CODE \n" +
                    " AND E.SECTION_CODE = S.SECTION_CODE \n" +
                    " AND F.SECTION_CODE = S.SECTION_CODE \n" +
                    " AND F.REGION_CODE is null  \n" +
                    " AND F.N_USERID <> '000199' \n" +
                    " AND F.N_USERID <> '000198' \n" +
                    " AND F.N_USERID <> '001541' \n" +
                    " AND F.N_USERID = G.N_USERID \n " +
                    " AND C.MARKETING_TYPE <> 'TEL' \n " +
                    //" AND B.POLICY_DT BETWEEN " + StrStartPolDTO + "\n" +
                    //" AND " + StrEndPolDTO + " " + condition + "\n" +
                    " AND B.POLICY_DT  BETWEEN TO_DATE('" + StrStartPolDT + " 00:00:00','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE('" + StrEndPolDT + " 23:59:59','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''')" + " " + condition + "\n" + 
                    " GROUP BY S.SECTION_CODE,S.SECTION_NAME,E.REGION_CODE,E.REGION_NAME,G.NAME,G.SURNAME,E.SEQ \n" +
                    " Order by E.SEQ ASC ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetBkkRegionDdl(string Dep,string Ofc)
        {
            string sql = "";
            string condition = "";
            string conofc = "";
            if (Dep != "")
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + " AND \n    ";

                condition = condition + " ztbofc.REGION_CODE = " + Utility.SQLValueString(Dep) + "\n";
            }
            if (Ofc != "")
            {
                if (conofc != "")
                    conofc = conofc + "   AND ";
                else
                    conofc = conofc + " AND \n    ";

                conofc = conofc + " ztbofc.OFFICE = " + Utility.SQLValueString(Ofc) + "\n";
            }
            sql = " select *from ZTB_BLA_REGION_OFFICE2 ztbofc,ztb_office ofc \n" +
                      " where ztbofc.OFFICE=ofc.OFFICE  " + condition + conofc +" \n "+
                      " and (ztbofc.OFFICE not like 'สสง') \n" +
                      " ORDER BY ofc.DESCRIPTION ASC ";           
               
           
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetBkkRegionDdl(string Sec, string Dep, string Ofc)
        {
            string sql = "";
            string condition = "";
            string conofc = "";
            if (Dep != "")
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + " AND \n    ";

                condition = condition + " ztbofc.REGION_CODE = " + Utility.SQLValueString(Dep) + "\n";
            }
            if (Ofc != "")
            {
                if (conofc != "")
                    conofc = conofc + "   AND ";
                else
                    conofc = conofc + " AND \n    ";

                conofc = conofc + " ztbofc.OFFICE = " + Utility.SQLValueString(Ofc) + "\n";
            }
            sql =     " SELECT * \n" +
                      " FROM ZTB_BLA_REGION_OFFICE ztbofc, ztb_office ofc \n" +
                      " WHERE ztbofc.OFFICE = ofc.OFFICE  " + condition + conofc + " \n " +
                      " AND (ztbofc.OFFICE not like 'สสง') \n" +
                      " ORDER BY ofc.DESCRIPTION ASC ";


            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }   //add by Korn 
        public DataTable GetPolicyTypeOfMonth(string startPolDT, string endPolDT)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);

            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";

            sql = "SELECT P.POLICY_TYPE " +
                  ",COUNT(P.POLICY) AS AMOUNT " +
                  ",SUM(P.SUMM) AS SUMM " +
                  ",SUM((P.PREMIUM1 + P.PREMIUM2 + P.R_PREMIUM + P.R_XTR_PREMIUM)) AS PREMIUM " +
                  "FROM POLICY_ALL_MAIN_NEW P " +
                  "WHERE P.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " " +
                  "GROUP BY P.POLICY_TYPE " +
                  "ORDER BY P.POLICY_TYPE ASC";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public DataTable GetUserOffice(string UserId) //add by Korn ใช้หาว่า user อยุ่ ส่วน ภาค สาขา ไหน
        {
            string sql = "";

            sql =   "SELECT     Us.Name \n" +
                    "           ,Us.Surname \n" +
                    "           ,Ofc.Office \n" +
                    "           ,Rf.Region_Code \n" +
                    "           ,Sec.Section_Code \n" +
                    "FROM   Ztb_User_Office OFC \n" +
                    "       INNER JOIN ZTB_BLA_REGION_OFFICE2 RF ON Rf.Office = OFC.Office \n" +
                    "       INNER JOIN Ztb_Bla_Region2 RE ON Re.Region_Code = Rf.Region_Code \n" +
                    "       INNER JOIN Ztb_Bla_Section SEC ON Sec.Section_Code = Re.Section_Code \n" +
                    "       INNER JOIN ZTB_USER US ON Us.N_Userid = Ofc.N_Userid \n" +
                    "WHERE Ofc.N_Userid = '" + UserId + "' and OFC.TMN_DT IS NULL";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }


        public DataTable GetPolicyOfDep(string startPolDT, string endPolDT, string typeDep, string branch, string bancType, string planType)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string WhereBranch = "";
            string whereBanc = " ";
            string wherePlan = " ";
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            if (branch == "00")
            {
                WhereBranch = "";
            }
            else
            {
                if (typeDep == "D")
                {
                    WhereBranch = " AND P.SALE_AGENT = '" + branch + "' " + NL;
                }
                else
                {
                    WhereBranch = " AND P.APP_OFC = '" + branch + "' " + NL;
                }
            }

            if (bancType == "")
            {
                whereBanc = " ";
            }
            else if (bancType == "1")
            {
                whereBanc = " AND (P.FLG_TYPE = 'A' OR (P.FLG_TYPE = 'B' AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E'))) " + NL;
            }
            else if (bancType == "2")
            {
                whereBanc = " AND P.FLG_TYPE = 'B' AND   P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE IN ('C','H')) " + NL;
            }

            if (planType == "")
            {
                wherePlan = " ";
            }
            else if ((planType == "G") || (planType == "P"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE  B.PL_CODE = P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) = '" + planType + "' " + NL;
            }
            else if ((planType == "H") || (planType == "C") || (planType == "E"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) = '" + planType + "' " + NL;
            }

            sql = "SELECT P.INSTALL_DT,P.POLICY_TYPE " + NL +
                  ",NVL(COUNT(P.POLICY),0) AS AMOUNT " + NL +
                  ",NVL(SUM(P.SUMM),0) AS SUMM " + NL +
                  ",NVL(SUM((P.PREMIUM1 + P.PREMIUM2)),0) AS PREMIUM " + NL +
                  ",SUM( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "           (SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N')) " + NL +
                  "        WHEN INSTALL_DT >= TO_DATE('06/09/2016','DD/MM/RRRR','nls_calendar=''gregorian''') THEN  " + NL +
                  "           (SELECT DECODE(COUNT(*),0,0,1) FROM P_POLICY_SENDING S,P_APPL_ID APL WHERE S.POLICY_ID = APL.POLICY_ID AND APL.APP_NO=P.APP_NO AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N') ) " + NL +
                  "        ELSE " + NL +
                  "           (SELECT DECODE(COUNT(*),0,0,1) FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO) " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_SEND " + NL +
                  " ,SUM( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "           (SELECT DECODE(COUNT(*),0,1,0) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N')) " + NL +
                   "        WHEN INSTALL_DT >= TO_DATE('06/09/2016','DD/MM/RRRR','nls_calendar=''gregorian''') THEN  " + NL +
                  "           (SELECT DECODE(COUNT(*),0,1,0) FROM P_POLICY_SENDING S,P_APPL_ID APL WHERE S.POLICY_ID = APL.POLICY_ID AND APL.APP_NO=P.APP_NO AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N') ) " + NL +
                  "        ELSE " + NL +
                  "           (SELECT DECODE(COUNT(*),0,1,0) FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO) " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_NOTSEND " + NL +
                  "FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "WHERE  P.POLICY_TYPE = '" + typeDep + "' " + NL +
                  "AND P.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + "  " + NL + WhereBranch + NL + whereBanc + NL + wherePlan + NL +
                  "GROUP BY  P.INSTALL_DT,P.POLICY_TYPE " + NL +
                  "ORDER BY P.INSTALL_DT ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable getRegion() {
            string sql = "";

            //sql = "select * from ztb_bla_region order by region_code asc"; //comment by Korn
            sql = "select * from ZTB_BLA_SECTION order by seq asc"; //add by Korn
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        
        
        
        }
        public DataTable getRegion_Sub(string Sec_Code) //add by Korn
        {
            string sql = "";

            sql = "select * from ztb_bla_region2 where SECTION_CODE = '"+ Sec_Code +"' order by seq asc";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }
        public DataTable GetPolicyOfficeOfDep(string section, string region, string office, string startPolDT, string endPolDT)  //add by Korn
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string txtsection = "";
            string txtsection2 = "";
            string txtregion = "";
            string txtoffice = "";
            string txtregion2 = "";
            string txtoffice2 = "";
            string txtoffice3 = "";
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";


            if (section != "") //add new 
            {
                if (txtsection != "")
                    txtsection = txtsection + "   AND ";
                else
                    txtsection = txtsection + "AND \n    ";

                txtsection = txtsection + "S.SECTION_CODE = " + Utility.SQLValueString(section) + "\n";
            }
            if (section != "") //add new 
            {
                if (txtsection2 != "")
                    txtsection2 = txtsection2 + "   AND ";
                else
                    txtsection2 = txtsection2 + "AND \n    ";

                txtsection2 = txtsection2 + "SEC.SECTION_CODE = " + Utility.SQLValueString(section) + "\n";
            }

            if (region != "")
            {
                if (txtregion != "")
                    txtregion = txtregion + "   AND ";
                else
                    txtregion = txtregion + "AND \n    ";

                txtregion = txtregion + "D.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }

            if (region != "")
            {
                if (txtregion2 != "")
                    txtregion2 = txtregion2 + "   AND ";
                else
                    txtregion2 = txtregion2 + "AND \n    ";

                txtregion2 = txtregion2 + "TB1.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }


            if (office != "")
            {
                if (txtoffice != "")
                    txtoffice = txtoffice + "   AND ";
                else
                    txtoffice = txtoffice + "AND \n    ";

                txtoffice = txtoffice + "D.OFFICE = " + Utility.SQLValueString(office) + "\n";
            }
            if (office == "")
            {
                if (txtoffice2 != "")
                    txtoffice2 = txtoffice2 + "    ";
                else
                    txtoffice2 = txtoffice2 + "\n    ";

                txtoffice2 = txtoffice2 +
                         " SELECT SUM(TB1.SEQ),decode(TB3.OFFICE,'สสง','สนญ',TB3.OFFICE) AS OFFICE,DECODE(TB3.DESCRIPTION,'สำนักงานใหญ่(สสง)','สำนักงานใหญ่',TB3.DESCRIPTION) AS DESCRIPTION \n" +
                         " ,SUM(NVL(TB2.AMOUNT_POLICY,0)) AS AMOUNT_POLICY \n" +
                         " ,SUM(NVL(TB2.SUMM_POLICY,0)) AS SUMM_POLICY \n" +
                         " ,SUM(NVL(TB2.PREMIUM_POLICY,0)) AS PREMIUM_POLICY  \n" +
                         " ,SUM(NVL(TB2.AMOUNT_SEND_POLICY,0)) AS AMOUNT_SEND_POLICY \n" +
                         " ,SUM(NVL(TB2.AMOUNT_NOTSEND_POLICY,0)) AS AMOUNT_NOTSEND_POLICY  \n" +
                         " ,SUM(NVL(TB2.AMOUNT_DAY_CLEAN,0)) AS AMOUNT_DAY_CLEAN \n" +
                         " ,SUM(NVL(TB2.AMOUNT_CLEAN,0)) AS AMOUNT_CLEAN \n" +
                         " ,SUM(NVL(TB2.CLEAN_PERCENT,0)) AS CLEAN_PERCENT  \n" +
                         " ,SUM(NVL(TB2.AMOUNT_DAY_UNCLEAN,0)) AS AMOUNT_DAY_UNCLEAN \n" +
                         " ,SUM(NVL(TB2.AMOUNT_UNCLEAN,0)) AS AMOUNT_UNCLEAN \n" +
                         " ,SUM(NVL(TB2.UNCLEAN_PERCENT,0)) AS UNCLEAN_PERCENT  \n" +
                         " ,SUM(NVL(TB2.TOTAL_AMOUNT_DAY,0)) AS TOTAL_AMOUNT_DAY \n" +
                         " ,SUM(NVL(TB2.TOTAL_AMOUNT,0)) AS TOTAL_AMOUNT   \n" +
                         " ,SUM(NVL(TB2.SUM_AMOUNT_DAY_CLEAN,0)) AS SUM_AMOUNT_DAY_CLEAN \n" +
                         " ,SUM(NVL(TB2.SUM_AMOUNT_DAY_UNCLEAN,0)) AS SUM_AMOUNT_DAY_UNCLEAN  \n" +
                         " ,DECODE(TB3.ISO_FLG,'Y','(ISO)',null) As ISO  \n" +
                         " FROM ZTB_BLA_REGION_OFFICE2 TB1,ZTB_OFFICE TB3, ZTB_BLA_REGION2 REG, ZTB_BLA_SECTION SEC,  \n" +
                        " ( \n";
            }
            if (office == "")
            {
                if (txtoffice3 != "")
                    txtoffice3 = txtoffice3 + "    ";
                else
                    txtoffice3 = txtoffice3 + "\n    ";

                txtoffice3 = txtoffice3 + " ) TB2 \n" +
                    " WHERE TB1.OFFICE = TB2.OFFICE(+) " + txtregion2 + txtsection2 + " \n" +
                    " AND TB1.OFFICE = TB3.OFFICE \n" +
                    " AND TB1.REGION_CODE = REG.REGION_CODE \n" +
                    " AND REG.SECTION_CODE = SEC.SECTION_CODE \n" +
                    " GROUP BY decode(TB3.OFFICE,'สสง','สนญ',TB3.OFFICE) \n" +
                    " ,DECODE(TB3.DESCRIPTION,'สำนักงานใหญ่(สสง)','สำนักงานใหญ่',TB3.DESCRIPTION) \n" +
                    " ,DECODE(TB3.ISO_FLG,'Y','(ISO)',null)  \n" +
                    "  ORDER BY 1,3 ASC  \n"
                     ;
            }

            sql =
                    txtoffice2 + " \n" +
                    " SELECT TB.OFFICE,TB.DESCRIPTION \n" +
                    "   ,COUNT(TB.POLICY_ID) AS AMOUNT_POLICY\n" +
                    "   ,SUM(TB.SUMM) AS SUMM_POLICY \n" +
                    "   ,SUM(TB.T_PREMIUM) AS PREMIUM_POLICY \n" +
                    "   ,SUM(TB.AMOUNT_SEND_POLICY) AS AMOUNT_SEND_POLICY \n" +
                    "   ,SUM(TB.AMOUNT_NOTSEND_POLICY) AS AMOUNT_NOTSEND_POLICY \n" +
                    "   ,DECODE(TB.ISO_FLG,'Y','(ISO)',null) As ISO   \n" +
                    "   ,DECODE(SUM(TB.AMOUNT_CLEAN),0,0,ROUND((SUM(TB.DAY_CLEAN)/SUM(TB.AMOUNT_CLEAN)))) AS AMOUNT_DAY_CLEAN \n" +
                    "   ,SUM(TB.AMOUNT_CLEAN) AS AMOUNT_CLEAN \n" +
                    "   ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_CLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS CLEAN_PERCENT \n" +
                    "   ,DECODE(SUM(TB.AMOUNT_UNCLEAN),0,0,ROUND((SUM(TB.DAY_UNCLEAN)/SUM(TB.AMOUNT_UNCLEAN)))) AS AMOUNT_DAY_UNCLEAN \n" +
                    "   ,SUM(TB.AMOUNT_UNCLEAN) AS AMOUNT_UNCLEAN \n" +
                    "   ,DECODE(SUM(TB.DAY_CLEAN),0,0,SUM(TB.DAY_CLEAN)) AS SUM_AMOUNT_DAY_CLEAN \n" +
                    "   ,DECODE(SUM(TB.DAY_UNCLEAN),0,0,SUM(TB.DAY_UNCLEAN)) AS SUM_AMOUNT_DAY_UNCLEAN \n" +
                    "   ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_UNCLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS UNCLEAN_PERCENT \n" +
                    "   ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.DAY_CLEAN) + SUM(TB.DAY_UNCLEAN))/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))) AS TOTAL_AMOUNT_DAY \n" +
                    "   ,(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)) AS TOTAL_AMOUNT \n" +
                    "   FROM \n" +
                    "   ( \n" +
                    "     SELECT D.OFFICE,H.DESCRIPTION,H.ISO_FLG,A.POLICY_ID,A.POLICY,NVL(A.SUMM,0) AS SUMM,NVL(A.T_PREMIUM,0) AS T_PREMIUM,J.IMB_APP_NO,TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR') AS SEND_DT \n" +
                    "     ,DECODE(I.SEND_DT,NULL,0,1) AS AMOUNT_SEND_POLICY,DECODE(I.SEND_DT,NULL,1,0) AS AMOUNT_NOTSEND_POLICY \n" +
                    "     ,TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR') AS IMB_APPSYS_DATE \n" +
                    "     ,( \n" +
                    "         CASE \n" +
                    "             WHEN I.SEND_DT IS NULL THEN \n" +
                    "                 0 \n" +
                    "             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
                    "                 1 \n" +
                    "             ELSE \n" +
                    "                 0 \n" +
                    "         END \n" +
                    "      ) AS AMOUNT_UNCLEAN \n" +
                    "          ,( \n" +
                    "         CASE \n" +
                    "             WHEN I.SEND_DT IS NULL THEN \n" +
                    "                 0 \n" +
                    "             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
                    "                  POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
                    "             ELSE \n" +
                    "                 0 \n" +
                    "         END \n" +
                    "      ) AS DAY_UNCLEAN \n" +
                    "     ,( \n" +
                    "         CASE \n" +
                    "             WHEN I.SEND_DT IS NULL THEN \n" +
                    "                 0 \n" +
                    "             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
                    "                 1 \n" +
                    "             ELSE \n" +
                    "                 0 \n" +
                    "         END \n" +
                    "      ) AS AMOUNT_CLEAN  \n" +
                    "      ,( \n" +
                    "         CASE \n" +
                    "             WHEN I.SEND_DT IS NULL THEN \n" +
                    "                 0 \n" +
                    "             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
                    "                  POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
                    "             ELSE \n" +
                    "                 0 \n" +
                    "         END \n" +
                    "      ) AS DAY_CLEAN \n" +
                    "     FROM  PW_LIFE_SUMM A \n"+
                    "           ,P_INSTALL B \n"+
                    "           ,P_AGENT C \n"+
                    "           ,ZTB_BLA_REGION_OFFICE2 D \n"+
                    "           ,ZTB_BLA_REGION2 E \n"+
                    "           ,ZTB_BLA_REGION_ID F \n"+
                    "           ,ZTB_USER G,ZTB_OFFICE H \n"+
                    "           ,P_POLICY_SENDING I \n"+
                    "           ,IS_APPM_BSC J \n"+
                    "           ,ZTB_BLA_SECTION S \n"+
                    "     WHERE A.CHANNEL_TYPE = 'OD' \n" +
                    "           AND A.PL_BLOCK = 'A' \n" +
                    "           AND A.POLICY_ID = B.POLICY_ID \n" +
                    "           AND A.POLICY_ID = C.POLICY_ID \n" +
                    "           AND B.APP_OFC = D.OFFICE \n" +
                    "           AND D.REGION_CODE = E.REGION_CODE \n" +
                    "           AND E.SECTION_CODE = S.SECTION_CODE --add \n"+
                    "           AND S.SECTION_CODE = F.SECTION_CODE --e \n"+
                    "           AND F.REGION_CODE is null  --a \n"+
                    "           AND F.N_USERID <> '000199' --a \n"+
                    "           AND F.N_USERID <> '000198' --a \n"+
                    "           AND F.N_USERID <> '001541' \n"+
                    "           AND F.N_USERID = G.N_USERID \n" +
                    "           AND D.OFFICE = H.OFFICE " + txtsection + txtregion + txtoffice + " \n " +
                    "           AND A.POLICY = I.POLICY(+) \n" +
                    "           AND I.SEQ(+) = 1 \n" +
                    "           AND A.POLICY = LPAD(J.IMB_POL_NO,8,'0') \n" +

                    "           AND C.MARKETING_TYPE <> 'TEL' \n" +
                    //"         AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " \n" +
                    "           AND B.POLICY_DT  BETWEEN TO_DATE('" + StrStartPolDT + " 00:00:00','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE('" + StrEndPolDT + " 23:59:59','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') \n" + 
                    "   ) TB \n" +
                    "   GROUP BY TB.OFFICE,TB.DESCRIPTION,TB.ISO_FLG \n" + txtoffice3;




            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyOfficeOfDep(string region, string office, string startPolDT, string endPolDT)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string txtregion = "";
            string txtoffice = "";
            string txtregion2 = "";
            string txtoffice2 = "";
            string txtoffice3 = "";
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";
            if (region != "")
            {
                if (txtregion != "")
                    txtregion = txtregion + "   AND ";
                else
                    txtregion = txtregion + "AND \n    ";

                txtregion = txtregion + "D.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }
            if (region != "")
            {
                if (txtregion2 != "")
                    txtregion2 = txtregion2 + "   AND ";
                else
                    txtregion2 = txtregion2 + "AND \n    ";

                txtregion2 = txtregion2 + "TB1.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }


            if (office != "")
            {
                if (txtoffice != "")
                    txtoffice = txtoffice + "   AND ";
                else
                    txtoffice = txtoffice + "AND \n    ";

                txtoffice = txtoffice + "D.OFFICE = " + Utility.SQLValueString(office) + "\n";
            }
            if (office == "")
            {
                if (txtoffice2 != "")
                    txtoffice2 = txtoffice2 + "    ";
                else
                    txtoffice2 = txtoffice2 + "\n    ";

                txtoffice2 = txtoffice2 +
                     " SELECT SUM(TB1.SEQ),decode(TB3.OFFICE,'สสง','สนญ',TB3.OFFICE) AS OFFICE,DECODE(TB3.DESCRIPTION,'สำนักงานใหญ่(สสง)','สำนักงานใหญ่',TB3.DESCRIPTION) AS DESCRIPTION \n" +
                     " ,SUM(NVL(TB2.AMOUNT_POLICY,0)) AS AMOUNT_POLICY \n" +
                     " ,SUM(NVL(TB2.SUMM_POLICY,0)) AS SUMM_POLICY \n" +
                     " ,SUM(NVL(TB2.PREMIUM_POLICY,0)) AS PREMIUM_POLICY  \n" +
                     " ,SUM(NVL(TB2.AMOUNT_SEND_POLICY,0)) AS AMOUNT_SEND_POLICY \n" +
                     " ,SUM(NVL(TB2.AMOUNT_NOTSEND_POLICY,0)) AS AMOUNT_NOTSEND_POLICY  \n" +
                     " ,SUM(NVL(TB2.AMOUNT_DAY_CLEAN,0)) AS AMOUNT_DAY_CLEAN \n" +
                     " ,SUM(NVL(TB2.AMOUNT_CLEAN,0)) AS AMOUNT_CLEAN \n" +
                     " ,SUM(NVL(TB2.CLEAN_PERCENT,0)) AS CLEAN_PERCENT  \n" +
                     " ,SUM(NVL(TB2.AMOUNT_DAY_UNCLEAN,0)) AS AMOUNT_DAY_UNCLEAN \n" +
                     " ,SUM(NVL(TB2.AMOUNT_UNCLEAN,0)) AS AMOUNT_UNCLEAN \n" +
                     " ,SUM(NVL(TB2.UNCLEAN_PERCENT,0)) AS UNCLEAN_PERCENT  \n" +
                     " ,SUM(NVL(TB2.TOTAL_AMOUNT_DAY,0)) AS TOTAL_AMOUNT_DAY \n" +
                     " ,SUM(NVL(TB2.TOTAL_AMOUNT,0)) AS TOTAL_AMOUNT   \n" +
                     " ,SUM(NVL(TB2.SUM_AMOUNT_DAY_CLEAN,0)) AS SUM_AMOUNT_DAY_CLEAN \n" +
                     " ,SUM(NVL(TB2.SUM_AMOUNT_DAY_UNCLEAN,0)) AS SUM_AMOUNT_DAY_UNCLEAN  \n" +
                     " ,DECODE(TB3.ISO_FLG,'Y','(ISO)',null) As ISO  \n" +
                     " FROM ZTB_BLA_REGION_OFFICE TB1,ZTB_OFFICE TB3,  \n" +
                    " ( \n";
            }

            if (office == "")
            {
                if (txtoffice3 != "")
                    txtoffice3 = txtoffice3 + "    ";
                else
                    txtoffice3 = txtoffice3 + "\n    ";

                txtoffice3 = txtoffice3 + " ) TB2 \n" +
                " WHERE TB1.OFFICE = TB2.OFFICE(+) " + txtregion2 + " \n" +
                " AND TB1.OFFICE = TB3.OFFICE \n" +
                " GROUP BY decode(TB3.OFFICE,'สสง','สนญ',TB3.OFFICE) \n" +
                " ,DECODE(TB3.DESCRIPTION,'สำนักงานใหญ่(สสง)','สำนักงานใหญ่',TB3.DESCRIPTION) \n" +
                " ,DECODE(TB3.ISO_FLG,'Y','(ISO)',null)  \n" +
                "  ORDER BY 1,3 ASC  \n"
                 ;
            }


            sql =
txtoffice2 + " \n" +
" SELECT TB.OFFICE,TB.DESCRIPTION \n" +
"  ,COUNT(TB.POLICY_ID) AS AMOUNT_POLICY\n" +
"   ,SUM(TB.SUMM) AS SUMM_POLICY \n" +
"   ,SUM(TB.T_PREMIUM) AS PREMIUM_POLICY \n" +
"   ,SUM(TB.AMOUNT_SEND_POLICY) AS AMOUNT_SEND_POLICY \n" +
"   ,SUM(TB.AMOUNT_NOTSEND_POLICY) AS AMOUNT_NOTSEND_POLICY \n" +
"   ,DECODE(TB.ISO_FLG,'Y','(ISO)',null) As ISO   \n" +
"   ,DECODE(SUM(TB.AMOUNT_CLEAN),0,0,ROUND((SUM(TB.DAY_CLEAN)/SUM(TB.AMOUNT_CLEAN)))) AS AMOUNT_DAY_CLEAN \n" +
"   ,SUM(TB.AMOUNT_CLEAN) AS AMOUNT_CLEAN \n" +
"    ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_CLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS CLEAN_PERCENT \n" +
"   ,DECODE(SUM(TB.AMOUNT_UNCLEAN),0,0,ROUND((SUM(TB.DAY_UNCLEAN)/SUM(TB.AMOUNT_UNCLEAN)))) AS AMOUNT_DAY_UNCLEAN \n" +
"   ,SUM(TB.AMOUNT_UNCLEAN) AS AMOUNT_UNCLEAN \n" +
" ,DECODE(SUM(TB.DAY_CLEAN),0,0,SUM(TB.DAY_CLEAN)) AS SUM_AMOUNT_DAY_CLEAN \n" +
" ,DECODE(SUM(TB.DAY_UNCLEAN),0,0,SUM(TB.DAY_UNCLEAN)) AS SUM_AMOUNT_DAY_UNCLEAN \n" +
"   ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_UNCLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS UNCLEAN_PERCENT \n" +
"   ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.DAY_CLEAN) + SUM(TB.DAY_UNCLEAN))/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))) AS TOTAL_AMOUNT_DAY \n" +
"   ,(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)) AS TOTAL_AMOUNT \n" +
"   FROM \n" +
"   ( \n" +
"     SELECT D.OFFICE,H.DESCRIPTION,H.ISO_FLG,A.POLICY_ID,A.POLICY,NVL(A.SUMM,0) AS SUMM,NVL(A.T_PREMIUM,0) AS T_PREMIUM,J.IMB_APP_NO,TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR') AS SEND_DT \n" +
"     ,DECODE(I.SEND_DT,NULL,0,1) AS AMOUNT_SEND_POLICY,DECODE(I.SEND_DT,NULL,1,0) AS AMOUNT_NOTSEND_POLICY \n" +
"     ,TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR') AS IMB_APPSYS_DATE \n" +
"     ,( \n" +
"         CASE \n" +
"             WHEN I.SEND_DT IS NULL THEN \n" +
"                 0 \n" +
"             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
"                 1 \n" +
"             ELSE \n" +
"                 0 \n" +
"         END \n" +
"      ) AS AMOUNT_UNCLEAN \n" +
"          ,( \n" +
"         CASE \n" +
"             WHEN I.SEND_DT IS NULL THEN \n" +
"                 0 \n" +
"             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
"                  POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
"             ELSE \n" +
"                 0 \n" +
"         END \n" +
"      ) AS DAY_UNCLEAN \n" +
"     ,( \n" +
"         CASE \n" +
"             WHEN I.SEND_DT IS NULL THEN \n" +
"                 0 \n" +
"             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
"                 1 \n" +
"             ELSE \n" +
"                 0 \n" +
"         END \n" +
"      ) AS AMOUNT_CLEAN  \n" +
"      ,( \n" +
"         CASE \n" +
"             WHEN I.SEND_DT IS NULL THEN \n" +
"                 0 \n" +
"             WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
"                  POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
"             ELSE \n" +
"                 0 \n" +
"         END \n" +
"      ) AS DAY_CLEAN \n" +
"     FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE D,ZTB_BLA_REGION E,ZTB_BLA_REGION_USER F,ZTB_USER G,ZTB_OFFICE H,P_POLICY_SENDING I,IS_APPM_BSC J \n" +
"     WHERE A.CHANNEL_TYPE = 'OD' \n" +
"     AND A.PL_BLOCK = 'A' \n" +
"     AND A.POLICY_ID = B.POLICY_ID \n" +
"     AND A.POLICY_ID = C.POLICY_ID \n" +
"     AND B.APP_OFC = D.OFFICE \n" +
"     AND D.REGION_CODE = E.REGION_CODE \n" +
"     AND E.REGION_CODE = F.REGION_CODE \n" +
"     AND F.N_USERID = G.N_USERID \n" +
"     AND D.OFFICE = H.OFFICE " + txtregion + txtoffice + " \n " +
"     AND A.POLICY = I.POLICY(+) \n" +
"     AND I.SEQ(+) = 1 \n" +
"     AND A.POLICY = LPAD(J.IMB_POL_NO,8,'0') \n" +

"     AND C.MARKETING_TYPE <> 'TEL' \n" +
"     AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " \n" +
"   ) TB \n" +
"   GROUP BY TB.OFFICE,TB.DESCRIPTION,TB.ISO_FLG \n" + txtoffice3;




            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyOfficeOfDate(string region, string office, string startPolDT, string endPolDT)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string txtregion = "";

            string txtoffice = "";
            string txtoffice2 = "";
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";
            if (region != "")
            {
                if (txtregion != "")
                    txtregion = txtregion + "   AND ";
                else
                    txtregion = txtregion + "AND \n    ";

                txtregion = txtregion + "D.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }
            if (office != "")
            {
                if (office == "สนญ")
                {
                    if (txtoffice != "")
                        txtoffice = txtoffice + "   AND ";
                    else
                        txtoffice = txtoffice + "AND \n    ";

                    txtoffice = txtoffice + "(D.OFFICE = " + Utility.SQLValueString(office) + "OR  D.OFFICE ='สสง') \n";
                }
                else
                {
                    if (txtoffice != "")
                        txtoffice = txtoffice + "   AND ";
                    else
                        txtoffice = txtoffice + "AND \n    ";

                    txtoffice = txtoffice + "D.OFFICE = " + Utility.SQLValueString(office) + "\n";
                }
            }

            sql =
                    "  SELECT TB.INSTALL_DT,TB.OFFICE,TB.POLICY_TYPE \n" +
                    " ,COUNT(TB.POLICY_ID) AS AMOUNT_POLICY \n" +
                    " ,SUM(TB.SUMM) AS SUMM_POLICY \n" +
                    " ,SUM(TB.T_PREMIUM) AS PREMIUM_POLICY \n" +
                    " ,SUM(TB.AMOUNT_SEND_POLICY) AS AMOUNT_SEND_POLICY \n" +
                    " ,SUM(TB.AMOUNT_NOTSEND_POLICY) AS AMOUNT_NOTSEND_POLICY \n" +
                    " ,DECODE(SUM(TB.DAY_CLEAN),0,0,SUM(TB.DAY_CLEAN)) AS SUM_AMOUNT_DAY_CLEAN \n" +
                    " ,DECODE(SUM(TB.DAY_UNCLEAN),0,0,SUM(TB.DAY_UNCLEAN)) AS SUM_AMOUNT_DAY_UNCLEAN  \n" +
                    " ,DECODE(SUM(TB.AMOUNT_CLEAN),0,0,ROUND((SUM(TB.DAY_CLEAN)/SUM(TB.AMOUNT_CLEAN)))) AS AMOUNT_DAY_CLEAN \n" +
                    " ,SUM(TB.AMOUNT_CLEAN) AS AMOUNT_CLEAN \n" +
                    " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_CLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS CLEAN_PERCENT  \n" +
                    " ,DECODE(SUM(TB.AMOUNT_UNCLEAN),0,0,ROUND((SUM(TB.DAY_UNCLEAN)/SUM(TB.AMOUNT_UNCLEAN)))) AS AMOUNT_DAY_UNCLEAN \n" +
                    " ,SUM(TB.AMOUNT_UNCLEAN) AS AMOUNT_UNCLEAN \n" +
                    " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_UNCLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS UNCLEAN_PERCENT \n" +
                    " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.DAY_CLEAN) + SUM(TB.DAY_UNCLEAN))/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))) AS TOTAL_AMOUNT_DAY \n" +
                    " ,(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)) AS TOTAL_AMOUNT \n" +
                    " FROM \n" +
                    " ( \n" +
                    "   SELECT B.INSTALL_DT,D.OFFICE,H.DESCRIPTION,A.POLICY_ID,A.POLICY,NVL(A.SUMM,0) AS SUMM,NVL(A.T_PREMIUM,0) AS T_PREMIUM,J.IMB_APP_NO,TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR') AS SEND_DT \n" +
                    "   ,DECODE(I.SEND_DT,NULL,0,1) AS AMOUNT_SEND_POLICY,DECODE(I.SEND_DT,NULL,1,0) AS AMOUNT_NOTSEND_POLICY ,DECODE(H.APP_REGION,'BKK','A','B') AS POLICY_TYPE\n" +
                    "   ,TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR') AS IMB_APPSYS_DATE \n" +
                    "   ,( \n" +
                    "       CASE \n" +
                    "           WHEN I.SEND_DT IS NULL THEN \n" +
                    "                0 \n" +
                    "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
                    "               1 \n" +
                    "           ELSE \n" +
                    "               0 \n" +
                    "       END \n" +
                    "    ) AS AMOUNT_UNCLEAN \n" +

                    "    ,( \n" +
                    "       CASE \n" +
                    "           WHEN I.SEND_DT IS NULL THEN \n" +
                    "               0 \n" +
                    "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
                    "                POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
                    "           ELSE \n" +
                    "               0 \n" +
                    "       END \n" +
                    "    ) AS DAY_UNCLEAN \n" +
                    "   ,( \n" +
                    "       CASE \n" +
                    "           WHEN I.SEND_DT IS NULL THEN \n" +
                    "               0 \n" +
                    "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
                    "               1 \n" +
                    "            ELSE \n" +
                    "               0 \n" +
                    "       END \n" +
                    "    ) AS AMOUNT_CLEAN \n" +

                    "    ,( \n" +
                    "       CASE \n" +
                    "           WHEN I.SEND_DT IS NULL THEN \n" +
                    "               0 \n" +
                    "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
                    "                POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
                    "           ELSE \n" +
                    "               0 \n" +
                    "       END \n" +
                    "    ) AS DAY_CLEAN \n" +
                    "   FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE2 D,ZTB_BLA_REGION2 E,ZTB_BLA_REGION_ID F,ZTB_USER G,ZTB_OFFICE H,P_POLICY_SENDING I,IS_APPM_BSC J \n" +
                    "   WHERE A.CHANNEL_TYPE = 'OD' \n" +
                    "   AND A.PL_BLOCK = 'A' \n" +
                    "   AND A.POLICY_ID = B.POLICY_ID \n" +
                    "   AND A.POLICY_ID = C.POLICY_ID \n" +
                    "   AND B.APP_OFC = D.OFFICE \n" +
                    "   AND D.REGION_CODE = E.REGION_CODE \n" +
                    "   AND E.REGION_CODE = F.REGION_CODE \n" +
                    "   AND F.N_USERID = G.N_USERID (+) \n" +
                    "   AND D.OFFICE = H.OFFICE " + txtregion + txtoffice + " \n " +
                    "   AND A.POLICY = I.POLICY(+) \n" +
                    "   AND I.SEQ(+) = 1 \n" +
                    "   AND A.POLICY = LPAD(J.IMB_POL_NO,8,'0') \n" +

                    "   AND C.MARKETING_TYPE <> 'TEL' \n" +

                    //"   AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " \n" +
                    "     AND B.POLICY_DT BETWEEN TO_DATE('" + StrStartPolDT + " 00:00:00','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE('" + StrEndPolDT + " 23:59:59','DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') \n" + 
                    " ) TB \n" +
                    " GROUP BY TB.INSTALL_DT,TB.OFFICE,TB.POLICY_TYPE \n" +
                    " ORDER BY TB.INSTALL_DT ASC \n"
                    ;


            #region comment 
            //sql =
            //        "  SELECT TB.INSTALL_DT,TB.OFFICE,TB.POLICY_TYPE \n" +
            //        " ,COUNT(TB.POLICY_ID) AS AMOUNT_POLICY \n" +
            //        " ,SUM(TB.SUMM) AS SUMM_POLICY \n" +
            //        " ,SUM(TB.T_PREMIUM) AS PREMIUM_POLICY \n" +
            //        " ,SUM(TB.AMOUNT_SEND_POLICY) AS AMOUNT_SEND_POLICY \n" +
            //        " ,SUM(TB.AMOUNT_NOTSEND_POLICY) AS AMOUNT_NOTSEND_POLICY \n" +
            //        " ,DECODE(SUM(TB.DAY_CLEAN),0,0,SUM(TB.DAY_CLEAN)) AS SUM_AMOUNT_DAY_CLEAN \n" +
            //        " ,DECODE(SUM(TB.DAY_UNCLEAN),0,0,SUM(TB.DAY_UNCLEAN)) AS SUM_AMOUNT_DAY_UNCLEAN  \n" +
            //        " ,DECODE(SUM(TB.AMOUNT_CLEAN),0,0,ROUND((SUM(TB.DAY_CLEAN)/SUM(TB.AMOUNT_CLEAN)))) AS AMOUNT_DAY_CLEAN \n" +
            //        " ,SUM(TB.AMOUNT_CLEAN) AS AMOUNT_CLEAN \n" +
            //        " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_CLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS CLEAN_PERCENT  \n" +
            //        " ,DECODE(SUM(TB.AMOUNT_UNCLEAN),0,0,ROUND((SUM(TB.DAY_UNCLEAN)/SUM(TB.AMOUNT_UNCLEAN)))) AS AMOUNT_DAY_UNCLEAN \n" +
            //        " ,SUM(TB.AMOUNT_UNCLEAN) AS AMOUNT_UNCLEAN \n" +
            //        " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.AMOUNT_UNCLEAN)/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))*100,2)) AS UNCLEAN_PERCENT \n" +
            //        " ,DECODE((SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)),0,0,ROUND((SUM(TB.DAY_CLEAN) + SUM(TB.DAY_UNCLEAN))/(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)))) AS TOTAL_AMOUNT_DAY \n" +
            //        " ,(SUM(TB.AMOUNT_CLEAN) + SUM(TB.AMOUNT_UNCLEAN)) AS TOTAL_AMOUNT \n" +
            //        " FROM \n" +
            //        " ( \n" +
            //        "   SELECT B.INSTALL_DT,D.OFFICE,H.DESCRIPTION,A.POLICY_ID,A.POLICY,NVL(A.SUMM,0) AS SUMM,NVL(A.T_PREMIUM,0) AS T_PREMIUM,J.IMB_APP_NO,TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR') AS SEND_DT \n" +
            //        "   ,DECODE(I.SEND_DT,NULL,0,1) AS AMOUNT_SEND_POLICY,DECODE(I.SEND_DT,NULL,1,0) AS AMOUNT_NOTSEND_POLICY ,DECODE(H.APP_REGION,'BKK','A','B') AS POLICY_TYPE\n" +
            //        "   ,TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR') AS IMB_APPSYS_DATE \n" +
            //        "   ,( \n" +
            //        "       CASE \n" +
            //        "           WHEN I.SEND_DT IS NULL THEN \n" +
            //        "                0 \n" +
            //        "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
            //        "               1 \n" +
            //        "           ELSE \n" +
            //        "               0 \n" +
            //        "       END \n" +
            //        "    ) AS AMOUNT_UNCLEAN \n" +
  
            //        "    ,( \n" +
            //        "       CASE \n" +
            //        "           WHEN I.SEND_DT IS NULL THEN \n" +
            //        "               0 \n" +
            //        "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) > 0 THEN \n" +
            //        "                POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
            //        "           ELSE \n" +
            //        "               0 \n" +
            //        "       END \n" +
            //        "    ) AS DAY_UNCLEAN \n" +
            //        "   ,( \n" +
            //        "       CASE \n" +
            //        "           WHEN I.SEND_DT IS NULL THEN \n" +
            //        "               0 \n" +
            //        "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
            //        "               1 \n" +
            //        "            ELSE \n" +
            //        "               0 \n" +
            //        "       END \n" +
            //        "    ) AS AMOUNT_CLEAN \n" +
  
            //        "    ,( \n" +
            //        "       CASE \n" +
            //        "           WHEN I.SEND_DT IS NULL THEN \n" +
            //        "               0 \n" +
            //        "           WHEN  I.SEND_DT IS NOT NULL AND (SELECT COUNT(AB.ISB_APP_NO) FROM IS_APPS_BSC AB WHERE  AB.ISB_APP_NO = J.IMB_APP_NO  AND AB.ISB_STS IN('AP','MO','MD','CO')) = 0 THEN \n" +
            //        "                POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(SUBSTR(J.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(J.IMB_APPSYS_DATE,3,2)||'/'||TO_CHAR(TO_NUMBER('25'||SUBSTR(J.IMB_APPSYS_DATE,1,2)) - 543),'DD/MM/RRRR'),TO_DATE(TO_CHAR(I.SEND_DT,'DD/MM/RRRR'),'DD/MM/RRRR')) \n" +
            //        "           ELSE \n" +
            //        "               0 \n" +
            //        "       END \n" +
            //        "    ) AS DAY_CLEAN \n" +
            //        "   FROM PW_LIFE_SUMM A,P_INSTALL B,P_AGENT C,ZTB_BLA_REGION_OFFICE D,ZTB_BLA_REGION E,ZTB_BLA_REGION_USER F,ZTB_USER G,ZTB_OFFICE H,P_POLICY_SENDING I,IS_APPM_BSC J \n" +
            //        "   WHERE A.CHANNEL_TYPE = 'OD' \n" +
            //        "   AND A.PL_BLOCK = 'A' \n" +
            //        "   AND A.POLICY_ID = B.POLICY_ID \n" +
            //        "   AND A.POLICY_ID = C.POLICY_ID \n" +
            //        "   AND B.APP_OFC = D.OFFICE \n" +
            //        "   AND D.REGION_CODE = E.REGION_CODE \n" +
            //        "   AND E.REGION_CODE = F.REGION_CODE \n" +
            //        "   AND F.N_USERID = G.N_USERID \n" +
            //        "   AND D.OFFICE = H.OFFICE " + txtregion + txtoffice + " \n " +
            //        "   AND A.POLICY = I.POLICY(+) \n" +
            //        "   AND I.SEQ(+) = 1 \n" +
            //        "   AND A.POLICY = LPAD(J.IMB_POL_NO,8,'0') \n" +

            //        "   AND C.MARKETING_TYPE <> 'TEL' \n" +

            //        "   AND B.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " \n" +
            //        " ) TB \n" +
            //        " GROUP BY TB.INSTALL_DT,TB.OFFICE,TB.POLICY_TYPE \n" +
            //        " ORDER BY TB.INSTALL_DT ASC \n"
            //        ;
            #endregion



            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public DataTable GetPolicyList(string startPolDT, string endPolDT, string startInstallDT, string endInstallDT, string typeDep, string branch, string bancType, string planType, string sendType, string caseType)
        {
            string sql = "";
            string cPolicyDateSt = "";
            string cPolicyDateEnd = "";
            string cInstallDateSt = "";
            string cInstallDateEnd = "";
            string wherePolicyDate = "";
            string whereInstallDate = "";
            string WhereBranch = "";
            string whereBanc = " ";
            string wherePlan = " ";
            string whereSend = " ";
            string whereCase = " ";

            if ((startPolDT != "") && (endPolDT != ""))
            {
                cPolicyDateSt = manage.GetDateFomatEN(startPolDT);
                cPolicyDateEnd = manage.GetDateFomatEN(endPolDT);
                wherePolicyDate = " AND P.POLICY_DT BETWEEN TO_DATE('" + cPolicyDateSt + "','DD/MM/RRRR','nls_calendar=''gregorian''') AND TO_DATE('" + cPolicyDateEnd + "','DD/MM/RRRR','nls_calendar=''gregorian''') " + NL;
            }
            else
            {
                wherePolicyDate = " ";
            }

            if ((startInstallDT != "") && (endInstallDT != ""))
            {
                cInstallDateSt = manage.GetDateFomatEN(startInstallDT);
                cInstallDateEnd = manage.GetDateFomatEN(endInstallDT);
                whereInstallDate = " AND P.INSTALL_DT BETWEEN TO_DATE('" + cInstallDateSt + "','DD/MM/RRRR','nls_calendar=''gregorian''')  AND TO_DATE('" + cInstallDateEnd + "','DD/MM/RRRR','nls_calendar=''gregorian''')  " + NL;
            }
            else
            {
                whereInstallDate = " ";
            }

            if (branch == "00")
            {
                WhereBranch = "";
            }
            else
            {
                if (typeDep == "D")
                {
                    WhereBranch = " AND P.SALE_AGENT = '" + branch + "' " + NL;
                }
                else
                {
                    WhereBranch = " AND P.APP_OFC = '" + branch + "' " + NL;
                }
            }

            if (bancType == "")
            {
                whereBanc = " ";
            }
            else if (bancType == "1")
            {
                whereBanc = " AND (P.FLG_TYPE = 'A' OR (P.FLG_TYPE = 'B' AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E')))  " + NL;
            }
            else if (bancType == "2")
            {
                whereBanc = " AND P.FLG_TYPE = 'B' AND   P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE IN ('C','H')) " + NL;
            }

            if (planType == "")
            {
                wherePlan = " ";
            }
            else if ((planType == "G") || (planType == "P"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE  B.PL_CODE = P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) = '" + planType + "' " + NL;
            }
            else if ((planType == "H") || (planType == "C") || (planType == "E"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) = '" + planType + "' " + NL;
            }

            if (sendType == "")
            {
                whereSend = " ";
            }
            else
            {
                whereSend = " AND " + NL +
                            "  ( " + NL +
                            "      CASE " + NL +
                            "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND (S.SEND_FAIL IS NULL or S.SEND_FAIL = 'N')) " + NL +
                            "          WHEN INSTALL_DT >= TO_DATE('06/09/2016','DD/MM/RRRR','nls_calendar=''gregorian''') THEN  " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S,P_APPL_ID APL WHERE S.POLICY_ID = APL.POLICY_ID AND APL.APP_NO=P.APP_NO AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N') ) " + NL +
                            "          ELSE " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO) " + NL +
                            "      END " + NL +
                            "  ) = '" + sendType + "' " + NL;
            }

            if (caseType == "C")
            {
                whereCase = "WHERE " + NL +
                            "( " + NL +
                            "    CASE " + NL +
                            "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                            "        1 " + NL +
                            "      ELSE " + NL +
                            "        0 " + NL +
                            "    END " + NL +
                            ") = 1 " + NL;
            }
            else if (caseType == "U")
            {
                whereCase = "WHERE " + NL +
                            "( " + NL +
                            "    CASE " + NL +
                            "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                            "        1 " + NL +
                            "      ELSE " + NL +
                            "        0 " + NL +
                            "    END " + NL +
                            ") = 1 " + NL;
            }
            else
            {
                whereCase = " ";
            }

            sql = "SELECT TB.POLICY,TB.APP_NO,TB.CERT_NO,TB.FLG_TYPE,TB.STATUS,DECODE(TB.FLG_TYPE,'A',TB.POLICY,'B',TB.POLICY||'/'||TB.CERT_NO) AS POLNO " + NL +
                  ",TB.IBBL_APP_DATE,TB.IBBL_TRN_DATE " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.IBBL_APP_DATE IS NOT NULL) AND (TB.IBBL_TRN_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.IBBL_APP_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_APP_DATE,1,2)||'/'||SUBSTR(TB.IBBL_APP_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_APP_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.IBBL_TRN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_TRN_DATE,1,2)||'/'||SUBSTR(TB.IBBL_TRN_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_TRN_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_BANK " + NL +
                  ",TB.CSP_PAYIN_DATE,TB.CSP_UPD_DATE_TIME " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.CSP_PAYIN_DATE IS NOT NULL) AND (TB.CSP_UPD_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.CSP_PAYIN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.CSP_PAYIN_DATE,1,2)||'/'||SUBSTR(TB.CSP_PAYIN_DATE,4,2)||'/'||(SUBSTR(TB.CSP_PAYIN_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.CSP_UPD_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.CSP_UPD_DATE,1,2)||'/'||SUBSTR(TB.CSP_UPD_DATE,4,2)||'/'||(SUBSTR(TB.CSP_UPD_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_CSP " + NL +
                  ",TB.APPSYS_DATE " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.IBBL_TRN_DATE IS NOT NULL) AND (TB.APPSYS_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.IBBL_TRN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_TRN_DATE,1,2)||'/'||SUBSTR(TB.IBBL_TRN_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_TRN_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_APPSYS " + NL +
                  ",TB.AP,TB.CO,TB.MO,TB.IMM2_PSTS_DATE,TB.ISB_STS_DATE,TB.INSTALL_DT " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.ISB_STS_DATE IS NOT NULL) AND (TB.INSTALL_DT IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.ISB_STS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.ISB_STS_DATE,1,2)||'/'||SUBSTR(TB.ISB_STS_DATE,4,2)||'/'||(SUBSTR(TB.ISB_STS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.INSTALL_DT,NULL,NULL,TO_DATE(SUBSTR(TB.INSTALL_DT,1,2)||'/'||SUBSTR(TB.INSTALL_DT,4,2)||'/'||(SUBSTR(TB.INSTALL_DT,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_INSTALL" + NL +
                  ",TB.SEND_DATE " + NL +
                  ",( " + NL +
                  "  CASE" + NL +
                  "    WHEN (TB.INSTALL_DT IS NOT NULL) AND (TB.SEND_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.INSTALL_DT,NULL,NULL,TO_DATE(SUBSTR(TB.INSTALL_DT,1,2)||'/'||SUBSTR(TB.INSTALL_DT,4,2)||'/'||(SUBSTR(TB.INSTALL_DT,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_SEND  " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.APPSYS_DATE IS NOT NULL) AND (TB.SEND_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_TOTAL " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN" + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_UNCLEAN" + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS AMOUNT_DAY_UNCLEAN " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_CLEAN " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS AMOUNT_DAY_CLEAN " + NL +
                  "FROM" + NL +
                  "( " + NL +
                  "  SELECT P.POLICY,P.APP_NO,P.CERT_NO,P.STATUS" + NL +
                  "  ,SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),7)+543) AS INSTALL_DT" + NL +
                  "  ,P.FLG_TYPE" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_APP_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          WHEN P.FLG_TYPE = 'B'  AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E')THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_APP_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT MAX(DECODE(H.IBRH_APP_DATE,NULL,NULL,SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_BBL_REF_H H WHERE H.IBRH_APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS IBBL_APP_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE" + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_PAY_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),7)+543)))   FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_PAYIN_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE " + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_UPD_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),7)+543))) FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_UPD_DATE" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||SUBSTR(C.CSP_UPD_TIME,1,2)||':'||SUBSTR(C.CSP_UPD_TIME,3,2)||':'||SUBSTR(C.CSP_UPD_TIME,5,2))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "               (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||TO_CHAR(CS.CSP_UPD_DATE,'HH24:MI:SS')))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE " + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_UPD_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),7)+543)||' '||SUBSTR(C.CFSP_UPD_TIME,1,2)||':'||SUBSTR(C.CFSP_UPD_TIME,3,2)||':'||SUBSTR(C.CFSP_UPD_TIME,5,2)))   FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "               (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||TO_CHAR(CS.CSP_UPD_DATE,'HH24:MI:SS')))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_UPD_DATE_TIME" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_TRN_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          WHEN P.FLG_TYPE = 'B'  AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E')THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_TRN_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT MAX(DECODE(H.IBRH_TRN_DATE,NULL,NULL,SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_BBL_REF_H H WHERE H.IBRH_APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS IBBL_TRN_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT DECODE(B.IMB_APPSYS_DATE,NULL,NULL,SUBSTR(B.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(B.IMB_APPSYS_DATE,3,2)||'/25'||SUBSTR(B.IMB_APPSYS_DATE,1,2)) FROM IS_APPM_BSC B WHERE B.IMB_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT DECODE(A.APP_DT,NULL,NULL,SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),7)+543))  FROM FSD_GM_APP A WHERE A.POLICY =  P.POLICY AND A.APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS APPSYS_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'AP')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'AP') " + NL +
                  "      END " + NL +
                  "    ) AS AP " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'CO')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'CO') " + NL +
                  "      END " + NL +
                  "    ) AS CO " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'MO')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'MO') " + NL +
                  "      END " + NL +
                  "    ) AS MO " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  "          ELSE " + NL +
                  "             (SELECT DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  "       END " + NL +
                  "   ) AS SEND_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DISTINCT  MAX(DECODE(M.IMM2_PSTS_DATE,NULL,NULL,SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_APPM_MO2 M WHERE M.IMM2_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "             (SELECT MAX(DECODE(M.RECEIVE_DT,NULL,NULL,SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),7)+543))) FROM FSD_GM_MOCODE M WHERE M.APP_NO = P.APP_NO AND M.POLICY = P.POLICY) " + NL +
                  "       END " + NL +
                  "   ) AS IMM2_PSTS_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DISTINCT MAX(DECODE(B.ISB_STS_DATE,NULL,NULL,SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_APPS_BSC B WHERE B.ISB_STS = 'IF' AND B.ISB_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "             NULL " + NL +
                  "       END " + NL +
                  "   ) AS ISB_STS_DATE " + NL +
                  "  FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "  WHERE P.POLICY_TYPE = '" + typeDep + "' " + wherePolicyDate + whereInstallDate + WhereBranch + whereBanc + wherePlan + whereSend +
                  ") TB " + NL + whereCase + NL +
                  "ORDER BY TB.INSTALL_DT,TB.POLICY ASC ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyList2(string startPolDT, string endPolDT, string startInstallDT, string endInstallDT, string typeDep, string branch, string bancType, string planType, string sendType, string caseType)
        {
            string sql = "";
            string cPolicyDateSt = "";
            string cPolicyDateEnd = "";
            string cInstallDateSt = "";
            string cInstallDateEnd = "";
            string wherePolicyDate = "";
            string whereInstallDate = "";
            string WhereBranch = "";
            string whereBanc = " ";
            string wherePlan = " ";
            string whereSend = " ";
            string whereCase = " ";
            string condition = "";

            if ((startPolDT != "") && (endPolDT != ""))
            {
                cPolicyDateSt = manage.GetDateFomatEN(startPolDT);
                cPolicyDateEnd = manage.GetDateFomatEN(endPolDT);
                wherePolicyDate = " AND P.POLICY_DT BETWEEN TO_DATE('" + cPolicyDateSt + "','DD/MM/RRRR','nls_calendar=''gregorian''')  AND TO_DATE('" + cPolicyDateEnd + "','DD/MM/RRRR','nls_calendar=''gregorian''')  " + NL;
            }
            else
            {
                wherePolicyDate = " ";
            }

            if ((startInstallDT != "") && (endInstallDT != ""))
            {
                cInstallDateSt = manage.GetDateFomatEN(startInstallDT);
                cInstallDateEnd = manage.GetDateFomatEN(endInstallDT);
                whereInstallDate = " AND P.INSTALL_DT BETWEEN TO_DATE('" + cInstallDateSt + "','DD/MM/RRRR','nls_calendar=''gregorian''')  AND TO_DATE('" + cInstallDateEnd + "','DD/MM/RRRR','nls_calendar=''gregorian''')  " + NL;
            }
            else
            {
                whereInstallDate = " ";
            }

            if (branch == "00")
            {
                WhereBranch = "";
            }
            else
            {
                if (typeDep == "D")
                {
                    WhereBranch = " AND P.SALE_AGENT = '" + branch + "' " + NL;
                }
                else
                {
                    if (branch == "สนญ")
                    {
                        WhereBranch = " AND (P.APP_OFC = 'สนญ' OR  P.APP_OFC ='สสง')" + NL;
                    }
                    else
                    {
                        WhereBranch = " AND P.APP_OFC = '" + branch + "' " + NL;
                    }
                }
            }

            if (bancType == "")
            {
                whereBanc = " ";
            }
            else if (bancType == "1")
            {
                whereBanc = " AND (P.FLG_TYPE = 'A' OR (P.FLG_TYPE = 'B' AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E'))) " + NL;
            }
            else if (bancType == "2")
            {
                whereBanc = " AND P.FLG_TYPE = 'B' AND   P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE IN ('C','H')) " + NL;
            }

            if (planType == "")
            {
                wherePlan = " ";
            }
            else if ((planType == "G") || (planType == "P"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE  B.PL_CODE = P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) = '" + planType + "' " + NL;
            }
            else if ((planType == "H") || (planType == "C") || (planType == "E"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) = '" + planType + "' " + NL;
            }

            if (sendType == "")
            {
                whereSend = " ";
            }
            else
            {
                whereSend = " AND " + NL +
                            "  ( " + NL +
                            "      CASE " + NL +
                            "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND (S.SEND_FAIL IS NULL || S.SEND_FAIL = 'N' )) " + NL +
                            "          WHEN INSTALL_DT >= TO_DATE('06/09/2016','DD/MM/RRRR','nls_calendar=''gregorian''') THEN  " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM P_POLICY_SENDING S,P_APPL_ID APL WHERE S.POLICY_ID = APL.POLICY_ID AND APL.APP_NO=P.APP_NO AND (S.SEND_FAIL IS NULL OR S.SEND_FAIL = 'N') ) " + NL +
                            "          ELSE " + NL +
                            "             (SELECT DECODE(COUNT(*),0,'N','Y') FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO) " + NL +
                            "      END " + NL +
                            "  ) = '" + sendType + "' " + NL;
            }

            if (caseType == "C")
            {
                whereCase = "WHERE " + NL +
                            "( " + NL +
                            "    CASE " + NL +
                            "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                            "        1 " + NL +
                            "      ELSE " + NL +
                            "        0 " + NL +
                            "    END " + NL +
                            ") = 1 " + NL;
            }
            else if (caseType == "U")
            {
                whereCase = "WHERE " + NL +
                            "( " + NL +
                            "    CASE " + NL +
                            "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                            "        1 " + NL +
                            "      ELSE " + NL +
                            "        0 " + NL +
                            "    END " + NL +
                            ") = 1 " + NL;
            }
            else
            {
                whereCase = " ";
            }
            if (typeDep != "" && typeDep != null && typeDep != " ")
            {
                if (condition == "")
                    condition = condition + " AND \n    ";
                else
                    condition = condition + "   AND ";
                condition = condition + "P.POLICY_TYPE = '" + typeDep + "'\n";
            }
            sql = "SELECT TB.POLICY,TB.APP_NO,TB.CERT_NO,TB.FLG_TYPE,TB.STATUS,DECODE(TB.FLG_TYPE,'A',TB.POLICY,'B',TB.CERT_NO) AS POLNO " + NL +
                  ",TB.IBBL_APP_DATE,TB.IBBL_TRN_DATE " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.IBBL_APP_DATE IS NOT NULL) AND (TB.IBBL_TRN_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.IBBL_APP_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_APP_DATE,1,2)||'/'||SUBSTR(TB.IBBL_APP_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_APP_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.IBBL_TRN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_TRN_DATE,1,2)||'/'||SUBSTR(TB.IBBL_TRN_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_TRN_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_BANK " + NL +
                  ",TB.CSP_PAYIN_DATE,TB.CSP_UPD_DATE_TIME " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.CSP_PAYIN_DATE IS NOT NULL) AND (TB.CSP_UPD_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.CSP_PAYIN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.CSP_PAYIN_DATE,1,2)||'/'||SUBSTR(TB.CSP_PAYIN_DATE,4,2)||'/'||(SUBSTR(TB.CSP_PAYIN_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.CSP_UPD_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.CSP_UPD_DATE,1,2)||'/'||SUBSTR(TB.CSP_UPD_DATE,4,2)||'/'||(SUBSTR(TB.CSP_UPD_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_CSP " + NL +
                  ",TB.APPSYS_DATE " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.IBBL_TRN_DATE IS NOT NULL) AND (TB.APPSYS_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.IBBL_TRN_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.IBBL_TRN_DATE,1,2)||'/'||SUBSTR(TB.IBBL_TRN_DATE,4,2)||'/'||(SUBSTR(TB.IBBL_TRN_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_APPSYS " + NL +
                  ",TB.AP,TB.CO,TB.MO,TB.IMM2_PSTS_DATE,TB.ISB_STS_DATE,TB.INSTALL_DT " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.ISB_STS_DATE IS NOT NULL) AND (TB.INSTALL_DT IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.ISB_STS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.ISB_STS_DATE,1,2)||'/'||SUBSTR(TB.ISB_STS_DATE,4,2)||'/'||(SUBSTR(TB.ISB_STS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.INSTALL_DT,NULL,NULL,TO_DATE(SUBSTR(TB.INSTALL_DT,1,2)||'/'||SUBSTR(TB.INSTALL_DT,4,2)||'/'||(SUBSTR(TB.INSTALL_DT,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_INSTALL" + NL +
                  ",TB.SEND_DATE " + NL +
                  ",( " + NL +
                  "  CASE" + NL +
                  "    WHEN (TB.INSTALL_DT IS NOT NULL) AND (TB.SEND_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.INSTALL_DT,NULL,NULL,TO_DATE(SUBSTR(TB.INSTALL_DT,1,2)||'/'||SUBSTR(TB.INSTALL_DT,4,2)||'/'||(SUBSTR(TB.INSTALL_DT,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_SEND  " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN (TB.APPSYS_DATE IS NOT NULL) AND (TB.SEND_DATE IS NOT NULL) THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS KPI_TOTAL " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN" + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_UNCLEAN" + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS AMOUNT_DAY_UNCLEAN " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  " ) AS AMOUNT_CLEAN " + NL +
                  ",( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),DECODE(TB.SEND_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.SEND_DATE,1,2)||'/'||SUBSTR(TB.SEND_DATE,4,2)||'/'||(SUBSTR(TB.SEND_DATE,7)-543),'DD/MM/RRRR')))  " + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  " ) AS AMOUNT_DAY_CLEAN " + NL +
                  "FROM" + NL +
                  "( " + NL +
                  "  SELECT P.POLICY,P.APP_NO,P.CERT_NO,P.STATUS" + NL +
                  "  ,SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(P.INSTALL_DT,'DD/MM/RRRR'),7)+543) AS INSTALL_DT" + NL +
                  "  ,P.FLG_TYPE" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_APP_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_APP_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT MAX(DECODE(H.IBRH_APP_DATE,NULL,NULL,SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(H.IBRH_APP_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_BBL_REF_H H WHERE H.IBRH_APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS IBBL_APP_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE" + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_PAY_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_PAY_DT,'DD/MM/RRRR'),7)+543)))   FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_PAYIN_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_PAYIN_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_PAYIN_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE " + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_UPD_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),7)+543))) FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "              (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_UPD_DATE" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              /*(SELECT MAX(DECODE(C.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||SUBSTR(C.CSP_UPD_TIME,1,2)||':'||SUBSTR(C.CSP_UPD_TIME,3,2)||':'||SUBSTR(C.CSP_UPD_TIME,5,2))) FROM CS_SUSPENSE C WHERE C.CSP_CANCEL IS NULL AND C.CSP_APP_NO = P.APP_NO) */" + NL +
                  "               (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||TO_CHAR(CS.CSP_UPD_DATE,'HH24:MI:SS')))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "          ELSE " + NL +
                  "              /*(SELECT MAX(DECODE(C.CFSP_UPD_DT,NULL,NULL,SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(C.CFSP_UPD_DT,'DD/MM/RRRR'),7)+543)||' '||SUBSTR(C.CFSP_UPD_TIME,1,2)||':'||SUBSTR(C.CFSP_UPD_TIME,3,2)||':'||SUBSTR(C.CFSP_UPD_TIME,5,2)))   FROM   CS_FSD_SP C WHERE  C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = P.APP_NO) */" + NL +
                  "               (SELECT MAX(DECODE(CS.CSP_UPD_DATE,NULL,NULL,SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(CS.CSP_UPD_DATE,'DD/MM/RRRR'),7)+543)||' '||TO_CHAR(CS.CSP_UPD_DATE,'HH24:MI:SS')))  FROM CS_SUSPENSE_VIEW CS WHERE CS.CSP_APP_NO = P.APP_NO  AND ( (CS.csp_disabled = 'N') OR (CS.csp_disabled = 'Y' AND CS.csp_cancel_type = 'N')))" + NL +
                  "      END " + NL +
                  "   ) AS CSP_UPD_DATE_TIME" + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT MAX(DECODE(R.IBBL_TRN_DATE,NULL,NULL,SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(R.IBBL_TRN_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT MAX(DECODE(H.IBRH_TRN_DATE,NULL,NULL,SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(H.IBRH_TRN_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_BBL_REF_H H WHERE H.IBRH_APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS IBBL_TRN_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (SELECT DECODE(B.IMB_APPSYS_DATE,NULL,NULL,SUBSTR(B.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(B.IMB_APPSYS_DATE,3,2)||'/25'||SUBSTR(B.IMB_APPSYS_DATE,1,2)) FROM IS_APPM_BSC B WHERE B.IMB_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "              (SELECT DECODE(A.APP_DT,NULL,NULL,SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(A.APP_DT,'DD/MM/RRRR'),7)+543))  FROM FSD_GM_APP A WHERE A.POLICY =  P.POLICY AND A.APP_NO = P.APP_NO) " + NL +
                  "      END " + NL +
                  "   ) AS APPSYS_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'AP')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'AP') " + NL +
                  "      END " + NL +
                  "    ) AS AP " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'CO')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'CO') " + NL +
                  "      END " + NL +
                  "    ) AS CO " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "              (WEB_PKG.GET_STSDATE(P.APP_NO,'MO')) " + NL +
                  "          ELSE " + NL +
                  "              WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'MO') " + NL +
                  "      END " + NL +
                  "    ) AS MO " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  "          ELSE " + NL +
                  "             (SELECT DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  "       END " + NL +
                  "   ) AS SEND_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DISTINCT  MAX(DECODE(M.IMM2_PSTS_DATE,NULL,NULL,SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(M.IMM2_PSTS_DATE,'DD/MM/RRRR'),7)+543)))  FROM IS_APPM_MO2 M WHERE M.IMM2_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "             (SELECT MAX(DECODE(M.RECEIVE_DT,NULL,NULL,SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(M.RECEIVE_DT,'DD/MM/RRRR'),7)+543))) FROM FSD_GM_MOCODE M WHERE M.APP_NO = P.APP_NO AND M.POLICY = P.POLICY) " + NL +
                  "       END " + NL +
                  "   ) AS IMM2_PSTS_DATE " + NL +
                  "  ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DISTINCT MAX(DECODE(B.ISB_STS_DATE,NULL,NULL,SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(B.ISB_STS_DATE,'DD/MM/RRRR'),7)+543))) FROM IS_APPS_BSC B WHERE B.ISB_STS = 'IF' AND B.ISB_APP_NO = P.APP_NO) " + NL +
                  "          ELSE " + NL +
                  "             NULL " + NL +
                  "       END " + NL +
                  "   ) AS ISB_STS_DATE " + NL +
                  "  FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "  WHERE p.policy is not null "+condition + wherePolicyDate + whereInstallDate + WhereBranch + whereBanc + wherePlan + whereSend +
                  ") TB " + NL + whereCase + NL +
                  "ORDER BY TB.INSTALL_DT,TB.POLICY ASC ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetMemoList()
        {
            string sql;
            sql = "SELECT * " +
                  "FROM " +
                  "( " +
                  "    SELECT 'ALL' AS CODE2,'ทั้งหมด' AS DESCRIPTION " +
                  "    FROM DUAL " +
                  "    UNION " +
                  "    SELECT C.CODE2,C.DESCRIPTION " +
                  "    FROM ZTB_CONSTANT3 C  " +
                  "    WHERE C.COL_NAME = 'MO_FLG' " +
                  ") TB " +
                  "ORDER BY TB.CODE2 ASC";
            //OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPayIn(string appNo)
        {
            string sql = "";
            //sql = "SELECT C.CSP_CALPREM " +
            //      ",C.CSP_SP_NO " +
            //      ",TO_CHAR(C.CSP_UPD_DATE,'DD/MM/RRRR') AS UPD_DATE " +
            //      ",SUBSTR(C.CSP_UPD_TIME,1,2)||':'||SUBSTR(C.CSP_UPD_TIME,3,2)||':'||SUBSTR(C.CSP_UPD_TIME,5,2) AS UPD_TIME " +
            //      ",(SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 = C.CSP_OPTION) AS PAY_BY " +
            //      ",C.CSP_PAYIN_DATE,C.CSP_APP_NO " +
            //      "FROM CS_SUSPENSE C " +
            //      "WHERE C.CSP_CANCEL IS NULL " +
            //      "AND C.CSP_APP_NO = SUBSTR('" + appNo + "',2) " +
            //      "ORDER BY C.CSP_PAYIN_DATE ASC";
            sql = "SELECT   S.CSP_CALPREM AS CSP_CALPREM,S.CSP_SP_NO AS CSP_SP_NO,S.CSP_UPD_DATE AS UPD_DATE,NULL AS UPD_TIME " + NL +
                  ",(SELECT SS.DESCRIPTION FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY " + NL +
                  ",S.CSP_PAYIN_DATE AS CSP_PAYIN_DATE,S.CSP_APP_NO AS CSP_APP_NO " + NL +
                  "FROM   CS_SUSPENSE_VIEW S " + NL +
                  "WHERE   S.CSP_APP_NO = '" + appNo + "' " +
                  "AND ( (S.csp_disabled = 'N') OR (S.csp_disabled = 'Y' AND S.csp_cancel_type = 'N')) " + NL +
                  "ORDER BY S.CSP_PAYIN_DATE ASC";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPayInBanc(string appNo)
        {
            string sql = "";
            //sql = "SELECT C.CFSP_CALPREM,C.CFSP_SP_NO,TO_CHAR(C.CFSP_UPD_DT,'DD/MM/')||(TO_CHAR(C.CFSP_UPD_DT,'RRRR')+543)||' '||SUBSTR(C.CFSP_UPD_TIME,1,2)||':'||SUBSTR(C.CFSP_UPD_TIME,3,2)||':'||SUBSTR(C.CFSP_UPD_TIME,5,2) AS UPD_DATE " + NL +
            //      ",(SELECT S.DESCRIPTION FROM ZTB_CONSTANT2 S WHERE S.COL_NAME = 'PAYMENT_OPT' AND S.CODE2 =C.CFSP_OPTION) AS PAY_BY " + NL +
            //      ",TO_CHAR(C.CFSP_PAY_DT,'DD/MM/')||(TO_CHAR(C.CFSP_PAY_DT,'RRRR')+543) AS PAYIN_DATE,C.CFSP_APP_NO " + NL +
            //      "FROM CS_FSD_SP C " + NL +
            //      "WHERE C.CFSP_CANCEL IS NULL " + NL +
            //      "AND C.CFSP_APP_NO = '" + appNo + "' " + NL +
            //      "ORDER BY C.CFSP_PAY_DT ASC";
            sql = "SELECT   S.CSP_CALPREM AS CFSP_CALPREM,S.CSP_SP_NO AS CFSP_SP_NO,S.CSP_UPD_DATE AS UPD_DATE " + NL +
                  ",(SELECT SS.DESCRIPTION FROM ZTB_CONSTANT2 SS WHERE SS.COL_NAME = 'PAYMENT_OPT' AND SS.CODE2 = S.CSP_PAY_OPTION) AS PAY_BY " + NL +
                  ",S.CSP_PAYIN_DATE AS PAYIN_DATE,S.CSP_APP_NO AS CFSP_APP_NO " + NL +
                  "FROM   CS_SUSPENSE_VIEW S " + NL +
                  "WHERE   S.CSP_APP_NO = '" + appNo + "' " + NL +
                  "AND ( (S.csp_disabled = 'N') OR (S.csp_disabled = 'Y' AND S.csp_cancel_type = 'N')) " + NL +
                  "ORDER BY S.CSP_PAYIN_DATE ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyDetail(string PolNo)
        {
            OleDbCommand com = new OleDbCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_POLICY_DETAIL";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = _MyConn;
            OleDbDataAdapter da = new OleDbDataAdapter(com);
            da.SelectCommand.Parameters.Add("POLICY_IN", OleDbType.VarChar).Value = PolNo;
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetSalecoByPol(string dep, string polNo)
        {
            string sql;
            if (dep == "D")
            {
                sql = "SELECT S.DESCRIPTION " +
                      ",WEB_PKG.GET_SALECO_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME " +
                      ",WEB_PKG.GET_MKG_BYAGENT(S.BBL_AGENTCODE) AS MKG_NAME " +
                      "FROM GBBL_STRUCT S " +
                      "WHERE S.BBL_AGENTCODE = (SELECT P.ASSIGN_AGENT FROM P_POLICY P WHERE P.POLICY = '" + polNo + "')";
            }
            else
            {
                sql = "SELECT D.AGENTCODE,D.UPLINE " +
                      ",NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = D.AGENTCODE),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = D.AGENTCODE))||' (สังกัด '||(SELECT O.DESCRIPTION AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = (NVL((SELECT OFFICE FROM GAG_DETAIL WHERE AGENTCODE = D.AGENTCODE),(SELECT OFFICE FROM GAG_INACTIVE WHERE AGENTCODE = D.AGENTCODE ))))||')' AS AGENT_NAME " +
                      ",NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = D.UPLINE),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = D.UPLINE)) AS UPLINE_NAME " +
                      "FROM GAG_DETAIL D " +
                      "WHERE D.AGENTCODE = (SELECT P.ASSIGN_AGENT FROM P_POLICY P WHERE P.POLICY = '" + polNo + "')";
            }

            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolByPremium(string PolNo)
        {
            OleDbCommand com = new OleDbCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_PREMIUMBYPOLICY";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = _MyConn;
            OleDbDataAdapter da = new OleDbDataAdapter(com);
            da.SelectCommand.Parameters.Add("POLICY_IN", OleDbType.VarChar).Value = PolNo;
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetAppByAppno(string TypeSearch, string ValueSearch)
        {
            string sql = "";
            string whereApp = "";
            string orderByApp = "";
            if (TypeSearch == "1")
            {
                whereApp = " AND LPAD(A.APP_NO,8,'0') = " + ValueSearch + " ";
                orderByApp = " ";
            }
            else if (TypeSearch == "2")
            {
                whereApp = " AND A.NAME LIKE '%" + ValueSearch + "%' ";
                orderByApp = " ORDER BY A.NAME ASC ";
            }
            sql = "SELECT LPAD(A.APP_NO,8,'0') AS APP_NO,A.NAME,A.PLAN,A.SUMM,NVL(A.PREMIUM+A.PREM+A.RIDER,0) AS PREMIUM,A.POLNO " +
                  ",DECODE(A.STATUS,'IF',DECODE(A.FLG_TYPE,'A','ออกเลขกรมธรรม์','ออกกรมธรรม์แล้ว'),NVL((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS),'กำลังพิจารณา')) AS STATUS " +
                  ",A.P_MODE,NVL(A.CALPREM,0) AS CALPREM,DECODE(B.IAPR_APP_NO,NULL,'A','B') AS KEY_IN " +
                  ",A.FLG_TYPE,A.AGENT_CODE,A.UPLINE,A.OFFICE,A.STATUS AS STATUS_CODE " +
                  ",WEB_PKG.GET_MEMODESC(A.APP_NO) AS STATUS_MEMO  " +
                  ",WEB_PKG.GET_MEMOCODE(A.APP_NO) AS MEMO_CODE " +
                  ",WEB_PKG.GET_CODESC(A.APP_NO) AS STATUS_CO " +
                  ",DECODE(A.BANK_ASS,'Y','D',DECODE (A.PL_BLOCK,'S', 'C',DECODE (A.REGION, 'BKK', 'A', 'B'))) AS DEP " +
                  ",DECODE(A.BANK_ASS,'Y',(SELECT 'ธนาคารกรุงเทพ '||S.DESCRIPTION||'('||BBL_BRANCH||')'||'/'||(WEB_PKG.GET_REGION_BYAGENT(A.AGENT_CODE)) AS DESCRIPTION FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = A.AGENT_CODE) " +
                  "        ,DECODE (A.PL_BLOCK,'S', 'สะสมเงินเดือน',(SELECT O.DESCRIPTION||'('||O.OFFICE||')' AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = A.OFFICE)) " +
                  "       ) AS BRANCH " +
                  ",(SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO)) AS SENDBY " +
                  ",(SELECT DECODE((SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,'วันที่นำส่งจากประกันชีวิต ',(SELECT DISTINCT 'วันที่ส่งไปรษณีย์ ' FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO)) AS SEND_PLACE " +
                  ",(SELECT DECODE((SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,S.SEND_DT,(SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO)) AS SEND_DATE " +
                  ",(SELECT P.STATUS FROM P_POLICY P WHERE P.POLICY = A.POLNO) AS POL_STATUS " +
                  "FROM WEB_APP_ALL A, IS_APP_REGION B " +
                  "WHERE A.APP_NO = B.IAPR_APP_NO(+) ";
            sql = sql + whereApp + orderByApp;
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolByPolno(string TypeSearch, string ValueSearch)
        {
            string sql = "";
            string wherePol = "";
            string orderByPol = "";
            if (TypeSearch == "1")
            {
                wherePol = " AND P.POLICY = " + ValueSearch + " ";
                orderByPol = " ";
            }
            else if (TypeSearch == "2")
            {
                wherePol = " AND N.NAME||'  '||N.SURNAME LIKE '%" + ValueSearch + "%' ";
                orderByPol = " ORDER BY N.NAME,N.SURNAME  ASC ";
            }
            sql = "SELECT P.POLICY,P.APP_NO,P.SUMM,N.NAME,N.SURNAME " +
                  ",NVL(POLICY.WEB_PKG.GET_PREMIUM_NXTDUE_DTN(P.POLICY),0) + NVL(POLICY.WEB_PKG.GET_RIDER_PREMIUM_NXTDUE_DTN(P.POLICY,P.STATUS,P.NXTDUE_DT),0) AS PREMIUM " +
                  ",(SELECT Z.BLA_TITLE AS TITLE FROM ZTB_PLAN Z  WHERE Z.PL_BLOCK = P.PL_BLOCK AND Z.PL_TYPE = P.PL_TYPE AND Z.PL_CODE = P.PL_CODE AND Z.PL_CODE2 = P.PL_CODE2) AS TITLE " +
                  ",(WEB_PKG.GET_STATUS(P.POLICY,P.STATUS,P.POL_YR,P.POL_LT)) AS STATUS_DES " +
                  ",(SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'P_MODE' AND C.CODE2 = LPAD(P.P_MODE,2,'0')) AS P_MODE " +
                  ",(SELECT Z.DESCRIPTION FROM  ZTB_OFFICE Z WHERE Z.OFFICE = (SELECT GA.OFFICE FROM GAG_DETAIL GA WHERE GA.AGENTCODE = P.ASSIGN_AGENT)) AS OFFICE " +
                  ",DECODE(P.BANCASS,'Y','D',DECODE(P.PL_BLOCK,'S','C',DECODE(O.REGION,'BKK','A','B'))) AS POLICY_TYPE " +
                  ",DECODE(P.BANCASS,'Y',(SELECT 'ธนาคารกรุงเทพ '||S.DESCRIPTION||'('||BBL_BRANCH||')'||'/'||(WEB_PKG.GET_REGION_BYAGENT(P.SALE_AGENT)) AS DESCRIPTION FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = P.SALE_AGENT) " +
                  "        ,DECODE (P.PL_BLOCK,'S', 'สะสมเงินเดือน',(SELECT Z.DESCRIPTION FROM  ZTB_OFFICE Z WHERE Z.OFFICE = (SELECT GA.OFFICE FROM GAG_DETAIL GA WHERE GA.AGENTCODE = P.ASSIGN_AGENT))) " +
                  "       ) AS BRANCH " +
                  ",P.STATUS " +
                  ",(SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY)) AS SENDBY " +
                  ",(SELECT DECODE((SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,'วันที่นำส่งจากประกันชีวิต ',(SELECT DISTINCT 'วันที่ส่งไปรษณีย์ ' FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY)) AS SEND_PLACE " +
                  ",(SELECT DECODE((SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR),NULL,S.SEND_DT,(SELECT DISTINCT D.DATEEND FROM DMIMPORTDATA D WHERE D.DMREF = S.REF_DEV||LPAD(S.POLICY,9,'0')||S.REF_YR)) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY)) AS SEND_DATE " +
                  "FROM P_POLICY P,ZTB_OFFICE O,P_NAME N " +
                  "WHERE P.LICENSE_OFC = O.OFFICE " +
                  "AND P.NAMECODE = N.NAMECODE ";
            sql = sql + wherePol + orderByPol;
            OleDbDataAdapter da = new OleDbDataAdapter(sql, _MyConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetBkkRegion(string Dep)
        {
            string sql = "";
            if (Dep == "D")
            {
                sql = "SELECT TB.OFFICE,TB.DESCRIPTION " +
                      "FROM " +
                      "( " +
                      "    SELECT '00' AS OFFICE,'แสดงข้อมูลทุกสาขา' AS DESCRIPTION,'A' AS FLG " +
                      "    FROM DUAL " +
                      "    UNION ALL " +
                      "    SELECT S.BBL_AGENTCODE AS OFFICE,S.DESCRIPTION,'B' AS FLG " +
                      "    FROM GBBL_STRUCT S " +
                      "    WHERE S.CLASS = 'B' " +
                      "    AND S.BBL_AGENTCODE IS NOT NULL " +
                      ") TB " +
                      "ORDER BY TB.FLG,TB.DESCRIPTION";
            }
            else
            {
                sql = "SELECT TB.OFFICE,TB.DESCRIPTION " +
                      "FROM " +
                      "( " +
                      "    SELECT '00' AS OFFICE,'แสดงข้อมูลทุกสาขา' AS DESCRIPTION,'A' AS FLG " +
                      "    FROM DUAL " +
                      "    UNION ALL " +
                      "    SELECT O.OFFICE,O.DESCRIPTION,'B' AS FLG " +
                      "    FROM ZTB_OFFICE O " +
                      "    WHERE O.TMN_DT > SYSDATE " +
                      ") TB " +
                      "ORDER BY TB.FLG,TB.DESCRIPTION";
            }
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetBanc(string StrYear, string StrMonth, string StAppSysDate, string EdAppoSysDate)
        {
            string myYear = StrYear.Substring(2, 2);
            string CStartDtAppSys = "";
            string CEndDtAppSys = "";
            string sql = "";
            string WhereAppSysDate = "";

            if ((StAppSysDate != "") && (EdAppoSysDate != ""))
            {
                CStartDtAppSys = manage.GetStrDate(StAppSysDate);
                CEndDtAppSys = manage.GetStrDate(EdAppoSysDate);
                WhereAppSysDate = " AND A.APPSYS_DATE BETWEEN " + CStartDtAppSys + " AND " + CEndDtAppSys + " " + NL;
            }
            else
            {
                WhereAppSysDate = "";
            }


            sql = "SELECT NVL(SUM(DECODE(TB.BANC_TYPE,'G',1,0)),0) AS AMOUNT_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'G',TB.SUMM,0)),0) AS SUMM_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'G',TB.PREMIUM,0)),0) AS PREMIUM_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',1,0)),0) AS AMOUNT_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',TB.SUMM,0)),0) AS SUMM_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',TB.PREMIUM,0)),0) AS PREMIUM_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',1,0)),0) AS AMOUNT_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',TB.SUMM,0)),0) AS SUMM_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',TB.PREMIUM,0)),0) AS PREMIUM_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',1,0)),0) AS AMOUNT_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',TB.SUMM,0)),0) AS SUMM_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',TB.PREMIUM,0)),0) AS PREMIUM_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',1,0)),0) AS AMOUNT_HL " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',TB.SUMM,0)),0) AS SUMM_HL " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',TB.PREMIUM,0)),0) AS PREMIUM_HL " + NL +
                  "FROM " + NL +
                  "( " + NL +
                  "  SELECT A.APP_NO,A.SUMM,(A.PREMIUM + A.PREM + A.RIDER) AS PREMIUM " + NL +
                  "  ,( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                  "    ELSE " + NL +
                  "      (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_BLOCK = 'B' AND B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                  "    END " + NL +
                  "   ) AS BANC_TYPE " + NL +
                  "  ,A.FLG_TYPE " + NL +
                  "  FROM WEB_APP_ALL_NEW A " + NL +
                  "  WHERE (A.BANK_ASS = 'Y' OR A.FLG_TYPE = 'C') " + NL +
                  "  AND SUBSTR(A.APP_DATE,1,2) = '" + myYear + "' " + NL +
                  "  AND SUBSTR(A.APP_DATE,3,2) = '" + StrMonth + "' " + WhereAppSysDate + NL +
                  ") TB";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetBancDetail(string AppNo, string PolNo)
        {
            string sql = "";
            sql = "SELECT A.APP_NO,A.POLICY,A.CERT_NO,A.APP_DT,A.ENTRY_DT,NULL AS CTM_TYPE " + "\n" +
                  ",N.PRENAME||' '||N.NAME AS NAMES " + "\n" +
                  ",N.BIRTH_DT,DECODE(N.SEX,'0','หญิง','1','ชาย') AS SEX,N.IDCARD_NO " + "\n" +
                  ",/*POLICY.ORD_PRODUCT.FCL_AGE(N.BIRTH_DT,SYSDATE,SYSDATE)*/A.ENTRY_AGE AS AGE " + "\n" +
                  ",N.ADDR1||' '||N.ADDR2||' '||N.ADDR3||' '||N.POST_CODE AS ADR " + "\n" +
                  ",N.TEL_NO AS TEL " + "\n" +
                  ",NVL((SELECT NVL(DESCRIPTION,'กำลังดำเนินการ') " + "\n" +
                  " FROM ZTB_CONSTANT2 " + "\n" +
                  " WHERE COL_NAME = 'LET_TYPE' " + "\n" +
                  " AND CODE2 =   (CASE " + "\n" +
                  "                WHEN (A.STATUS = '10') AND (A.CERT_NO > 0) " + "\n" +
                  "                THEN " + "\n" +
                  "                   'IF' " + "\n" +
                  "                WHEN (A.STATUS = '35') THEN " + "\n" +
                  "                 CASE " + "\n" +
                  "                     WHEN (SELECT   M.LET_TYPE  FROM   FSD_GM_MO M  WHERE   M.LET_CANCEL IS NULL  AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) IN('NT','CC','DC','PP','NO') " + "\n" +
                  "                         THEN " + "\n" +
                  "                             DECODE ((SELECT   M.LET_TYPE FROM   FSD_GM_MO M WHERE   M.LET_CANCEL IS NULL AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO),'NO','NT',(SELECT   M.LET_TYPE  FROM   FSD_GM_MO M  WHERE   M.LET_CANCEL IS NULL  AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO)) " + "\n" +
                  "                     ELSE " + "\n" +
                  "                             'NT' " + "\n" +
                  "                     END " + "\n" +
                  "                WHEN (SELECT   M.LET_TYPE " + "\n" +
                  "                      FROM   FSD_GM_MO M " + "\n" +
                  "                      WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                      AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                       FROM FSD_GM_MO M " + "\n" +
                  "                                       WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                       AND M.POLICY = A.POLICY " + "\n" +
                  "                                       AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                      AND M.POLICY = A.POLICY " + "\n" +
                  "                      AND M.APP_NO = A.APP_NO) IS NOT NULL " + "\n" +
                  "                THEN " + "\n" +
                  "                      DECODE ( " + "\n" +
                  "                              (SELECT   M.LET_TYPE " + "\n" +
                  "                               FROM   FSD_GM_MO M " + "\n" +
                  "                               WHERE   M.LET_CANCEL IS NULL " + "\n" +
                  "                               AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                                FROM FSD_GM_MO M " + "\n" +
                  "                                                WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                                AND M.POLICY = A.POLICY " + "\n" +
                  "                                                AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                               AND M.POLICY = A.POLICY " + "\n" +
                  "                               AND M.APP_NO = A.APP_NO),'NO','NT','CG','CO',(SELECT M.LET_TYPE " + "\n" +
                  "                                                                  FROM FSD_GM_MO M " + "\n" +
                  "                                                                  WHERE   M.LET_CANCEL IS NULL " + "\n" +
                  "                                                                  AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                                                                   FROM FSD_GM_MO M " + "\n" +
                  "                                                                                   WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                                                                   AND M.POLICY = A.POLICY " + "\n" +
                  "                                                                                   AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                                                                  AND M.POLICY = A.POLICY " + "\n" +
                  "                                                                  AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                              ) " + "\n" +
                  "                ELSE " + "\n" +
                  "                   (CASE " + "\n" +
                  "                       WHEN (A.STATUS = '10') AND (A.CERT_NO = 0) THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '00') THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '20') THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '35') THEN 'CC' " + "\n" +
                  "                    END) " + "\n" +
                  "             END)),'กำลังดำเนินการ') AS STS " + "\n" +
                  "           ,(CASE " + "\n" +
                  "                WHEN (A.STATUS = '10') AND (A.CERT_NO > 0) " + "\n" +
                  "                THEN " + "\n" +
                  "                   'IF' " + "\n" +
                  "                WHEN (A.STATUS = '35') THEN " + "\n" +
                  "                 CASE " + "\n" +
                  "                     WHEN (SELECT   M.LET_TYPE  FROM   FSD_GM_MO M  WHERE   M.LET_CANCEL IS NULL  AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) IN('NT','CC','DC','PP','NO') " + "\n" +
                  "                         THEN " + "\n" +
                  "                             DECODE ((SELECT   M.LET_TYPE FROM   FSD_GM_MO M WHERE   M.LET_CANCEL IS NULL AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO),'NO','NT',(SELECT   M.LET_TYPE  FROM   FSD_GM_MO M  WHERE   M.LET_CANCEL IS NULL  AND M.LET_PRT = (SELECT   MAX (M.LET_PRT) FROM   FSD_GM_MO M WHERE M.LET_CANCEL IS NULL AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO) AND M.POLICY = A.POLICY AND M.APP_NO = A.APP_NO)) " + "\n" +
                  "                     ELSE " + "\n" +
                  "                             'NT' " + "\n" +
                  "                     END " + "\n" +
                  "                WHEN (SELECT   M.LET_TYPE " + "\n" +
                  "                      FROM   FSD_GM_MO M " + "\n" +
                  "                      WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                      AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                       FROM FSD_GM_MO M " + "\n" +
                  "                                       WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                       AND M.POLICY = A.POLICY " + "\n" +
                  "                                       AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                      AND M.POLICY = A.POLICY " + "\n" +
                  "                      AND M.APP_NO = A.APP_NO) IS NOT NULL " + "\n" +
                  "                THEN " + "\n" +
                  "                      DECODE ( " + "\n" +
                  "                              (SELECT   M.LET_TYPE " + "\n" +
                  "                               FROM   FSD_GM_MO M " + "\n" +
                  "                               WHERE   M.LET_CANCEL IS NULL " + "\n" +
                  "                               AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                                FROM FSD_GM_MO M " + "\n" +
                  "                                                WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                                AND M.POLICY = A.POLICY " + "\n" +
                  "                                                AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                               AND M.POLICY = A.POLICY " + "\n" +
                  "                               AND M.APP_NO = A.APP_NO),'NO','NT','CG','CO',(SELECT M.LET_TYPE " + "\n" +
                  "                                                                  FROM FSD_GM_MO M " + "\n" +
                  "                                                                  WHERE   M.LET_CANCEL IS NULL " + "\n" +
                  "                                                                  AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                                                                   FROM FSD_GM_MO M " + "\n" +
                  "                                                                                   WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                                                                   AND M.POLICY = A.POLICY " + "\n" +
                  "                                                                                   AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                                                                  AND M.POLICY = A.POLICY " + "\n" +
                  "                                                                  AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                              ) " + "\n" +
                  "                ELSE " + "\n" +
                  "                   (CASE " + "\n" +
                  "                       WHEN (A.STATUS = '10') AND (A.CERT_NO = 0) THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '00') THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '20') THEN NULL " + "\n" +
                  "                       WHEN (A.STATUS = '35') THEN 'CC' " + "\n" +
                  "                    END) " + "\n" +
                  "             END) AS STATUS_CODE " + "\n" +
                  "           ,(CASE " + "\n" +
                  "                WHEN (SELECT C.STATUS_DT FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') IS NOT NULL THEN\n" + 
                  "                    (SELECT C.STATUS_DT FROM U_APPLICATION_ID AA,U_APPLICATION B,U_STATUS_ID C WHERE AA.APP_NO = A.APP_NO AND AA.CHANNEL_TYPE = 'GM' AND AA.APP_ID = B.APP_ID AND B.UAPP_ID = C.UAPP_ID AND C.TMN = 'N') \n" +
                  "                WHEN (A.STATUS = '10') AND (A.CERT_NO > 0) " + "\n" +
                  "                THEN " + "\n" +
                  "                   (SELECT M.ENTRY_DT AS ENTRY_DT FROM FSD_GM_MAST M WHERE M.CERT_NO = A.CERT_NO AND M.POLICY = A.POLICY) " + "\n" +
                  "                WHEN (SELECT   M.LET_TYPE " + "\n" +
                  "                      FROM   FSD_GM_MO M " + "\n" +
                  "                      WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                      AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                       FROM FSD_GM_MO M " + "\n" +
                  "                                       WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                       AND M.POLICY = A.POLICY " + "\n" +
                  "                                       AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                      AND M.POLICY = A.POLICY " + "\n" +
                  "                      AND M.APP_NO = A.APP_NO) IS NOT NULL " + "\n" +
                  "                THEN " + "\n" +
                  "                     (SELECT   M.LET_PRT " + "\n" +
                  "                      FROM   FSD_GM_MO M " + "\n" +
                  "                      WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                      AND M.LET_PRT = (SELECT MAX (M.LET_PRT) " + "\n" +
                  "                                       FROM FSD_GM_MO M " + "\n" +
                  "                                       WHERE M.LET_CANCEL IS NULL " + "\n" +
                  "                                       AND M.POLICY = A.POLICY " + "\n" +
                  "                                       AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                      AND M.POLICY = A.POLICY " + "\n" +
                  "                      AND M.APP_NO = A.APP_NO) " + "\n" +
                  "                ELSE " + "\n" +
                  "                 A.ENTRY_DT " + "\n" +
                  "                END) AS STATUS_DT " + "\n" +
                  ",(SELECT B.DESCRIPTION FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL AND B.POLICY = A.POLICY) AS POLICY_TYP \n" +
                  ",A.SUMM " + "\n" +
                  ",(SELECT Z2.DESCRIPTION FROM ZTB_CONSTANT2 Z2 WHERE  Z2.COL_NAME = 'P_MODE' AND Z2.CODE2 = A.AMODE) AS P_MODE " + "\n" +
                  ",A.CAL_PREM AS PREMIUM " + "\n" +
                  "/*,A.EXTRA_PREM AS PREM*/ " + "\n" +
                  ",DECODE(NVL(A.EXTRA_PREM,0),0,NVL((SELECT M.EM_PREMIUM FROM FSD_GM_MOEMPREM M WHERE M.POLICY = A.POLICY  AND M.APP_NO = TO_CHAR (A.APP_NO) AND M.SEQ = (SELECT MAX(SEQ) FROM FSD_GM_MOEMPREM WHERE POLICY  = A.POLICY  AND APP_NO = TO_CHAR (A.APP_NO))),0),NVL(A.EXTRA_PREM,0)) AS PREM " + "\n" +
                  ",0 AS RIDER " + "\n" +
                  "/*,(SELECT   NVL (SUM (C.CFSP_CALPREM), 0) FROM   CS_FSD_SP C WHERE C.CFSP_CANCEL IS NULL AND C.CFSP_APP_NO = A.APP_NO)AS CALPREM */" + "\n" +
                  ",(SELECT   NVL (SUM (S.csp_calprem), 0) AS csp_calprem FROM   CS_SUSPENSE_VIEW S WHERE   S.CSP_APP_NO = A.APP_NO AND (S.csp_clearing_status <> 'N') AND ( (S.csp_disabled = 'N') OR (S.csp_disabled = 'Y' AND S.csp_cancel_type = 'N'))) AS CALPREM \n" +
                  ",A.ASS_TERM,A.PAY_TERM " + "\n" +
                  ",LPAD (A.AGT_CODE, 8, '0') AS AGENT_CODE " + "\n" +
                  ",LPAD (A.AGT_CODE, 8, '0') AS UPLINE " + "\n" +
                  ",(SELECT S.DESCRIPTION  FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = LPAD (A.AGT_CODE, 8, '0')) AS BRANCH " + "\n" +
                  ",( " + "\n" +
                  "  SELECT S.SEND_DT " + "\n" +
                  "  FROM FSD_GM_MAST_SEND S " + "\n" +
                  "  WHERE S.POLICY = A.POLICY " + "\n" +
                  "  AND S.CERT_NO = A.CERT_NO " + "\n" +
                  "  AND S.SEND_FAIL = 'N' " + "\n" +
                  "  AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLICY AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + "\n" +
                  " ) AS SEND_DT " + "\n" +
                  "FROM FSD_GM_APP A,FSD_GM_NAME N " + "\n" +
                  "WHERE A.CUST_NO = N.CUST_NO " + "\n" +
                  "AND A.POLICY = '" + PolNo + "' " + "\n" +
                  "AND A.APP_NO = '" + AppNo + "' ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetCoBancDetail(string AppNo, string PolNo)
        {
            string sql = "";
            //sql = "SELECT TB.TITLE " +
            //      "FROM " +
            //      "( " +
            //      "SELECT " +
            //      "( " +
            //      "CASE " +
            //      "    WHEN M.CVEX IS NOT NULL THEN " +
            //      "        'ชำระเบี้ยประกันภัยเพิ่มจำนวน '||(M.SUMM*M.CVEX)/1000||' บาท' " +
            //      "    ELSE " +
            //      "        NULL " +
            //      "END " +
            //      ") AS TITLE " +
            //      "FROM FSD_GM_MO M " +
            //      "WHERE M.LET_TYPE IN ('CO','CG') " +
            //      "AND M.POLICY = '" + PolNo + "' " +
            //      "AND M.APP_NO = '" + AppNo + "' " +
            //      "UNION ALL " +
            //      "SELECT " +
            //      "( " +
            //      "CASE " +
            //      "    WHEN (M.SUMMS IS NOT NULL) AND (M.YRS IS NOT NULL) THEN " +
            //      "        'คงระยะเวลาประกันไว้ '||M.YRS||' ปี เหมือนเดิม แต่ขอลดทุนประกันเหลือ '||M.SUMMS||' บาท' " +
            //      "    ELSE " +
            //      "        NULL " +
            //      "END " +
            //      ") AS TITLE " +
            //      "FROM FSD_GM_MO M " +
            //      "WHERE M.LET_TYPE IN ('CO','CG') " +
            //      "AND M.POLICY = '" + PolNo + "' " +
            //      "AND M.APP_NO = '" + AppNo + "' " +
            //      "UNION ALL " +
            //      "SELECT " +
            //      "( " +
            //      "CASE " +
            //      "    WHEN (M.SUMMF IS NOT NULL) AND (M.YRF IS NOT NULL) THEN " +
            //      "        'ขอลดระยะเวลาประกันเหลือ '||M.YRF||' ปี และขอลดทุนประกันเหลือ '||M.SUMMF||' บาท' " +
            //      "    ELSE " +
            //      "        NULL " +
            //      "END " +
            //      ") AS TITLE " +
            //      "FROM FSD_GM_MO M " +
            //      "WHERE M.LET_TYPE IN ('CO','CG') " +
            //      "AND M.POLICY = '" + PolNo + "' " +
            //      "AND M.APP_NO = '" + AppNo + "' " +
            //      "UNION ALL " + NL +
            //      "SELECT G.EXCLUDE AS TITLE " + NL +
            //      "FROM U_APPLICATION_ID A,U_APPLICATION B,W_UNDERWRITE_APPLICATION C,W_SUBUNDERWRITE_ID D,W_SUMMARY E,U_EXCLUSION_ID F,U_EXCLUSION_DETAIL G" + NL +
            //      "WHERE A.APP_ID = B.APP_ID" + NL +
            //      "AND B.UAPP_ID = C.UAPP_ID" + NL +
            //      "AND C.UNDERWRITE_ID = D.UNDERWRITE_ID" + NL +
            //      "AND D.SUBUNDERWRITE_ID = E.SUBUNDERWRITE_ID" + NL +
            //      "AND E.SUMMARY_ID = F.SUMMARY_ID" + NL +
            //      "AND F.UEXCLUDE_ID = G.UEXCLUDE_ID" + NL +
            //      "AND G.TMN = 'N'" + NL +
            //      "AND F.TMN = 'N'" + NL +
            //      "AND E.TMN = 'N'" + NL +
            //      "AND D.CUSTOMER_TYPE = 'C'" + NL +
            //      "AND A.CHANNEL_TYPE = 'GM'" + NL +
            //      "AND A.APP_NO = '" + AppNo + "' " + NL +
            //      ") TB " +
            //      "WHERE TB.TITLE IS NOT NULL ";
            sql = "SELECT TB.TITLE " + NL +
                  "FROM " + NL +
                  "( " + NL +
                  "    SELECT C.EF_PREMIUM AS PREM, 'เพิ่มเบี้ยประกันภัยเพิ่มพิเศษชั่วคราว จำนวน '||C.EF_PREMIUM ||' บาท/งวด เป็นเวลา '||C.EF_PAY_TERM||' ปี' AS TITLE,'X' AS STATUS,'1' AS FLG,NULL AS STATUS_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,U_LFEF C " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE IN('GM','HL') " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.TMN = 'N' " + NL +
                  "    UNION " + NL +
                  "    SELECT C.EM_PREMIUM AS PREM,'เพิ่มเบี้ยประกันภัยเพิ่มพิเศษ จำนวน '||C.EM_PREMIUM||' บาท/งวด' AS TITLE,'X' AS STATUS,'1' AS FLG,NULL AS STATUS_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,U_LFEM C " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE = 'GM' " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.TMN = 'N' " + NL +
                  "    UNION " + NL +
                  "    SELECT C.XOCP_PREMIUM AS PREM, 'เพิ่มเบี้ยประกันภัยเพิ่มพิเศษจากอาชีพจำนวน '||C.XOCP_PREMIUM ||' บาท/งวด' AS TITLE,'X' AS STATUS,'1' AS FLG,NULL AS STATUS_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,U_LFXOCP C " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE = 'GM' " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.TMN = 'N' " + NL +
                  "    UNION " + NL +
                  "    SELECT D.XTR_PREMIUM AS PREM,'เพิ่มเบี้ยประกันภัยเพิ่มพิเศษ '||(SELECT P.TITLE_ABBR FROM ZTB_PLAN P WHERE P.PL_BLOCK = 'R' AND P.PL_CODE = C.PL_CODE  AND PL_CODE2 = C.PL_CODE2)||' จำนวน '||D.XTR_PREMIUM ||' บาท/งวด' AS TITLE,'X' AS STATUS,'2' AS FLG,NULL AS STATUS_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,U_RIDER_ID C,U_RDXTR D " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE = 'GM' " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.TMN = 'N' " + NL +
                  "    AND C.URIDER_ID = D.URIDER_ID(+) " + NL +
                  "    AND D.TMN = 'N' " + NL +
                  "    UNION " + NL +
                  "    SELECT  D.EM_PREMIUM AS PREM,'เพิ่มเบี้ยประกันภัยเพิ่มพิเศษ '||(SELECT P.TITLE_ABBR FROM ZTB_PLAN P WHERE P.PL_BLOCK = 'R' AND P.PL_CODE = C.PL_CODE  AND PL_CODE2 = C.PL_CODE2)||' จำนวน '||D.EM_PREMIUM ||' บาท/งวด' AS TITLE,'X' AS STATUS,'2' AS FLG,NULL AS STATUS_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,U_RIDER_ID C,U_RDEM D " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE = 'GM' " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.TMN = 'N' " + NL +
                  "    AND C.URIDER_ID = D.URIDER_ID(+) " + NL +
                  "    AND D.TMN = 'N' " + NL +
                  "    UNION " + NL +
                  "    SELECT 9999999 AS PREM,G.EXCLUDE AS TITLE,G.ADMIT_FLG,'3' AS FLG,G.ADMIT_DT " + NL +
                  "    FROM U_APPLICATION_ID A,U_APPLICATION B,W_UNDERWRITE_APPLICATION C,W_SUBUNDERWRITE_ID D,W_SUMMARY E,U_EXCLUSION_ID F,U_EXCLUSION_DETAIL G " + NL +
                  "    WHERE A.APP_ID = B.APP_ID " + NL +
                  "    AND A.APP_NO = '" + AppNo + "' " + NL +
                  "    AND A.CHANNEL_TYPE = 'GM' " + NL +
                  "    AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "    AND C.UNDERWRITE_ID = D.UNDERWRITE_ID " + NL +
                  "    AND D.CUSTOMER_TYPE = 'C' " + NL +
                  "    AND D.SUBUNDERWRITE_ID = E.SUBUNDERWRITE_ID " + NL +
                  "    AND E.SUMMARY_ID = F.SUMMARY_ID " + NL +
                  "    AND F.TMN = 'N' " + NL +
                  "    AND F.UEXCLUDE_ID = G.UEXCLUDE_ID " + NL +
                  "    AND G.TMN = 'N' " + NL +
                  ") TB" + NL +
                  "WHERE TB.TITLE IS NOT NULL " + NL +
                  "ORDER BY TB.FLG ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn); 
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetMoBancDetail(string AppNo, string PolNo)
        {
            string sql = "";
            //sql = "SELECT M.PEND_SEQ,M.PEND_CODE,M.PEND_DT,P.DESCRIPTION||' '||M.PEND_REMARK AS DESCRIPTION,M.RECEIVE_DT,P.MO_FLG " +
            //      "FROM FSD_GM_MOCODE M,ATB_PEND_REQM P " +
            //      "WHERE M.PEND_CODE = P.PEND_CODE " +
            //      "AND M.POLICY = '" + PolNo + "' " +
            //      "AND M.APP_NO = '" + AppNo + "' " +
            //      "ORDER BY M.PEND_SEQ ASC";
            //sql = "SELECT TB.PEND_SEQ,TB.PEND_CODE,TB.PEND_DT,TB.DESCRIPTION,TB.RECEIVE_DT,TB.MO_FLG " + NL +
            //      "FROM" + NL +
            //      "(" + NL +
            //      "    SELECT M.PEND_SEQ,M.PEND_CODE,M.PEND_DT,P.DESCRIPTION||' '||M.PEND_REMARK AS DESCRIPTION,M.RECEIVE_DT,P.MO_FLG " + NL +
            //      "    FROM FSD_GM_MOCODE M,ATB_PEND_REQM P " + NL +
            //      "    WHERE M.PEND_CODE = P.PEND_CODE " + NL +
            //      "    AND M.POLICY = '" + PolNo + "' " + NL +
            //      "    AND M.APP_NO = '" + AppNo + "' " + NL +
            //      "    UNION" + NL +
            //      "    SELECT G.PRINT_SEQ AS PEND_SEQ ,G.PEND_CODE AS PEND_CODE,F.MEMO_TRN_DT AS PEND_DT,H.DESCRIPTION AS DESCRIPTION,G.PEND_STATUS_DT AS RECEIVE_DT,NULL AS MO_FLG" + NL +
            //      "    FROM U_APPLICATION_ID A,U_APPLICATION B,W_UNDERWRITE_APPLICATION C,W_SUBUNDERWRITE_ID D,W_SUMMARY E,U_MEMO_ID F,U_MEMO_DETAIL G,AUTB_PEND_REQM H" + NL +
            //      "    WHERE A.APP_ID = B.APP_ID" + NL +
            //      "    AND B.UAPP_ID = C.UAPP_ID" + NL +
            //      "    AND C.UNDERWRITE_ID = D.UNDERWRITE_ID" + NL +
            //      "    AND D.SUBUNDERWRITE_ID = E.SUBUNDERWRITE_ID" + NL +
            //      "    AND E.SUMMARY_ID = F.SUMMARY_ID" + NL +
            //      "    AND F.UMEMO_ID = G.UMEMO_ID" + NL +
            //      "    AND G.PEND_CODE = H.PEND_CODE" + NL +
            //      "    AND F.TMN = 'N'" + NL +
            //      "    AND E.TMN = 'N'" + NL +
            //      "    AND D.CUSTOMER_TYPE = 'C'" + NL +
            //      "    AND A.CHANNEL_TYPE = 'GM'" + NL +
            //      "    AND A.APP_NO = '" + AppNo + "' " + NL +
            //      ") TB" + NL +
            //      "ORDER BY TB.PEND_SEQ ASC";
            sql = "SELECT G.PRINT_SEQ AS PEND_SEQ,G.PEND_CODE,F.MEMO_TRN_DT AS PEND_DT,(SELECT QM.DESCRIPTION FROM AUTB_PEND_REQM QM WHERE QM.PEND_CODE = G.PEND_CODE)||' '||I.PEND_DESCRIPTION AS DESCRIPTION,G.PEND_STATUS_DT AS RECEIVE_DT,NULL AS MO_FLG " + NL +
                  "FROM U_APPLICATION_ID A,U_APPLICATION B,W_UNDERWRITE_APPLICATION C,W_SUBUNDERWRITE_ID D,W_SUMMARY E,U_MEMO_ID F,U_MEMO_DETAIL G,U_STATUS_ID H,U_MEMO_DESCRIPTION I " + NL +
                  "WHERE A.APP_ID = B.APP_ID " + NL +
                  "AND A.APP_NO = '" + AppNo + "' " + NL +
                  "AND A.CHANNEL_TYPE IN('GM','HL') " + NL +
                  "AND B.UAPP_ID = C.UAPP_ID " + NL +
                  "AND C.UNDERWRITE_ID = D.UNDERWRITE_ID " + NL +
                  "AND D.CUSTOMER_TYPE = 'C' " + NL +
                  "AND D.SUBUNDERWRITE_ID = E.SUBUNDERWRITE_ID " + NL +
                  "AND E.TMN = 'N' " + NL +
                  "AND E.SUMMARY_ID = F.SUMMARY_ID " + NL +
                  "AND F.TMN = 'N' " + NL +
                  "AND F.UMEMO_ID = G.UMEMO_ID " + NL +
                  "AND B.UAPP_ID = H.UAPP_ID " + NL +
                  "AND H.TMN = 'N' " + NL +
                  "AND G.UMEMO_ID = I.UMEMO_ID(+) " + NL +
                  "AND G.PEND_CODE = I.PEND_CODE(+) " + NL +
                  "ORDER BY G.PRINT_SEQ ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetStatusApp()
        {
            string sql = "";
            sql = "SELECT TB.STATUS,TB.DESCRIPTION " +
                  "FROM " +
                  "( " +
                  "    SELECT 'ALL' AS STATUS,'เลือกสถานะทั้งหมด' AS DESCRIPTION,'A' AS FLG " +
                  "    FROM DUAL " +
                  "    UNION ALL " +
                  "    SELECT DECODE(T.TCT_CODE2,'  ','PO',T.TCT_CODE2),T.TCT_DESC AS DESCRIPTION,'B' AS FLG " +
                  "    FROM TB_CONSTANT T " +
                  "    WHERE T.TCT_CODE1 = 'AS' " +
                  ") TB " +
                  "ORDER BY TB.FLG,TB.DESCRIPTION ASC";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPlan(string depCode, string BancCode)
        {
            string sql = "";
            string vDep = "";
            string whereBanc = " ";
            if ((depCode == "A") || (depCode == "B"))
            {
                vDep = "A";
            }
            else if (depCode == "C")
            {
                vDep = "S";
            }
            if (depCode == "D")
            {
                if (BancCode == "ALL")
                {
                    whereBanc = " ";
                }
                else
                {
                    whereBanc = " AND B.BANC_TYPE = '" + BancCode + "' " + NL;
                }
                sql = "SELECT TB.PLAN_CODE,TB.TITLE " + NL +
                       "FROM " + NL +
                       "( " + NL +
                       "    SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " + NL +
                       "    FROM DUAL " + NL +
                       "    UNION ALL " + NL +
                       "    SELECT 'A'||B.OPL_CODE||NVL(B.OPL_CODE2,0) AS PLAN_CODE  " + NL +
                       "    ,B.DESCRIPTION||'(A/'||B.OPL_CODE||'/'||NVL(B.OPL_CODE2,0)||')' AS TITLE,'B' AS FLG  " + NL +
                       "    FROM ZTB_PLAN_BANC B " + NL +
                       "    WHERE B.POLICY IS NULL " + whereBanc + NL +
                       "    UNION ALL " + NL +
                       "    SELECT B.POLICY AS PLAN_CODE,B.DESCRIPTION||'('||B.POLICY||')' AS TITLE,'C' AS FLG " + NL +
                       "    FROM ZTB_PLAN_BANC B " + NL +
                       "    WHERE B.POLICY IS NOT NULL " + whereBanc + NL +
                       ") TB " + NL +
                       "ORDER BY TB.FLG,TB.TITLE ASC";


            }
            else if (depCode == "ALL")
            {
                sql = "SELECT TB.PLAN_CODE,TB.TITLE " + NL +
                      "FROM " + NL +
                      "( " + NL +
                      "    SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " + NL +
                      "    FROM DUAL " + NL +
                      "    UNION ALL " + NL +
                      "    SELECT DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||P.OPL_CODE||NVL(P.OPL_CODE2,0) AS PLAN_CODE " + NL +
                      "    ,P.BLA_TITLE||'('||DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||'/'||P.OPL_CODE||'/'||NVL(P.OPL_CODE2,0)||')' AS TITLE,'B' AS FLG " + NL +
                      "    FROM  ZTB_PLAN P " + NL +
                      "    WHERE P.PL_BLOCK IN('A','S','B') " + NL +
                      "    UNION ALL " + NL +
                      "    SELECT B.POLICY AS PLAN_CODE,B.DESCRIPTION||'('||B.POLICY||')' AS TITLE,'C' AS FLG " + NL +
                      "    FROM ZTB_PLAN_BANC B " + NL +
                      "    WHERE B.POLICY IS NOT NULL " + NL +
                      ") TB " + NL +
                      "ORDER BY TB.FLG,TB.TITLE ASC";
            }
            else
            {
                if (depCode == "E")
                {
                    sql = "SELECT TB.PLAN_CODE,TB.TITLE " + NL +
                          "FROM " + NL +
                          "( " + NL +
                          "  SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " + NL +
                          "  FROM DUAL " + NL +
                          "  UNION ALL " + NL +
                          "  SELECT DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||P.OPL_CODE||NVL(P.OPL_CODE2,0) AS PLAN_CODE " + NL +
                          "  ,P.BLA_TITLE||'('||DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||'/'||P.OPL_CODE||'/'||NVL(P.OPL_CODE2,0)||')' AS TITLE,'B' AS FLG " + NL +
                          "  FROM ZTB_PLAN P,ZTB_PLAN_MARKETING_TYPE M " + NL +
                          "  WHERE P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 = M.PLAN_ID " + NL +
                          "  AND M.MARKETING_TYPE = 'TEL'" + NL +
                          ") TB " + NL +
                          "ORDER BY TB.FLG,TB.TITLE ASC";                    
                }
                else
                {
                    sql = "SELECT TB.PLAN_CODE,TB.TITLE " + NL +
                          "FROM " + NL +
                          "( " + NL +
                          "  SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " + NL +
                          "  FROM DUAL " + NL +
                          "  UNION ALL " + NL +
                          "  SELECT DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||P.OPL_CODE||NVL(P.OPL_CODE2,0) AS PLAN_CODE " + NL +
                          "  ,P.BLA_TITLE||'('||DECODE(P.PL_BLOCK,'B','A',P.PL_BLOCK)||'/'||P.OPL_CODE||'/'||NVL(P.OPL_CODE2,0)||')' AS TITLE,'B' AS FLG " + NL +
                          "  FROM ZTB_PLAN P,ZTB_PLAN_MARKETING_TYPE M " + NL +
                          "  WHERE P.PL_BLOCK = '" + vDep + "' " + NL +
                          "  AND P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 = M.PLAN_ID " + NL +
                          "  AND M.MARKETING_TYPE <> 'TEL'" + NL +
                          ") TB " + NL +
                          "ORDER BY TB.FLG,TB.TITLE ASC";
                }

            }
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }
        public DataTable GetPlanPolicy(string depCode, string BancCode)
        {
            string sql = "";
            string vDep = "";
            string whereBanc = " ";
            if ((depCode == "A") || (depCode == "B"))
            {
                vDep = "A";
            }
            else if (depCode == "C")
            {
                vDep = "S";
            }
            if (depCode == "D")
            {
                if (BancCode == "ALL")
                {
                    whereBanc = " ";
                }
                else
                {
                    whereBanc = " AND B.BANC_TYPE = '" + BancCode + "' ";
                }
                sql = "SELECT TB.PLAN_CODE,TB.TITLE " + NL +
                       "FROM " + NL +
                       "( " + NL +
                       "    SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " + NL +
                       "    FROM DUAL " + NL +
                       "    UNION ALL " + NL +
                       "    SELECT B.PL_BLOCK||B.PL_CODE||B.PL_CODE2 AS PLAN_CODE  " + NL +
                       "    ,B.DESCRIPTION||'('||B.PL_BLOCK||'/'||B.PL_CODE||'/'||B.PL_CODE2||')' AS TITLE,'B' AS FLG  " + NL +
                       "    FROM ZTB_PLAN_BANC B " + NL +
                       "    WHERE B.POLICY IS NULL " + whereBanc + NL +
                       "    UNION ALL " + NL +
                       "    SELECT B.POLICY AS PLAN_CODE,B.DESCRIPTION||'('||B.POLICY||')' AS TITLE,'C' AS FLG " + NL +
                       "    FROM ZTB_PLAN_BANC B " + NL +
                       "    WHERE B.POLICY IS NOT NULL " + whereBanc + NL +
                       ") TB " + NL +
                       "ORDER BY TB.FLG,TB.TITLE ASC";


            }
            else if (depCode == "ALL")
            {
                sql = "SELECT TB.PLAN_CODE,TB.TITLE " +
                      "FROM " +
                      "( " +
                      "    SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " +
                      "    FROM DUAL " +
                      "    UNION ALL " +
                      "    SELECT P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 AS PLAN_CODE " +
                      "    ,P.BLA_TITLE||'('||P.PL_BLOCK||'/'||P.PL_CODE||'/'||P.PL_CODE2||')' AS TITLE,'B' AS FLG " +
                      "    FROM  ZTB_PLAN P " +
                      "    WHERE P.PL_BLOCK IN('A','S','B') " +
                      "    UNION ALL " +
                      "    SELECT B.POLICY AS PLAN_CODE,B.DESCRIPTION||'('||B.POLICY||')' AS TITLE,'C' AS FLG " +
                      "    FROM ZTB_PLAN_BANC B " +
                      "    WHERE B.POLICY IS NOT NULL " +
                      ") TB " +
                      "ORDER BY TB.FLG,TB.TITLE ASC";
            }
            else
            {
                if (depCode == "E")
                {
                    sql = "SELECT TB.PLAN_CODE,TB.TITLE " +
                          "FROM " +
                          "( " +
                          "  SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " +
                          "  FROM DUAL " +
                          "  UNION ALL " +
                          "  SELECT P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 AS PLAN_CODE " +
                          "  ,P.BLA_TITLE||'('||P.PL_BLOCK||'/'||P.PL_CODE||'/'||P.PL_CODE2||')' AS TITLE,'B' AS FLG " +
                          "  FROM ZTB_PLAN P,ZTB_PLAN_MARKETING_TYPE M " +
                          "  WHERE P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 = M.PLAN_ID " +
                          "  AND M.MARKETING_TYPE = 'TEL' " +
                          ") TB " +
                          "ORDER BY TB.FLG,TB.TITLE ASC";                    
                }
                else
                {
                    sql = "SELECT TB.PLAN_CODE,TB.TITLE " +
                          "FROM " +
                          "( " +
                          "  SELECT 'ALL' AS PLAN_CODE,'เลือกแบบประกันทั้งหมด' AS TITLE,'A' AS FLG " +
                          "  FROM DUAL " +
                          "  UNION ALL " +
                          "  SELECT P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 AS PLAN_CODE " +
                          "  ,P.BLA_TITLE||'('||P.PL_BLOCK||'/'||P.PL_CODE||'/'||P.PL_CODE2||')' AS TITLE,'B' AS FLG " +
                          "  FROM ZTB_PLAN P,ZTB_PLAN_MARKETING_TYPE M " +
                          "  WHERE P.PL_BLOCK = '" + vDep + "' " +
                          "  AND P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 = M.PLAN_ID " + 
                          "  AND M.MARKETING_TYPE <> 'TEL' " +
                          ") TB " +
                          "ORDER BY TB.FLG,TB.TITLE ASC";
                }

            }
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }
        public DataTable GetAppSearch(string depCode, string planCode, string statusCode, string name, string appNo, string polNo, string summSt, string summEnd, string bblDtSt, string bblDtEnd, string appDtSt, string appDtEnd, string appSysDtSt, string appSysDtEnd, string banc)
        {
            string sql = "";
            string whereDep = " ";
            string wherePlan = " ";
            string whereStatus = " ";
            string whereName = " ";
            string whereAppNo = " ";
            string wherePolNo = " ";
            string whereSumm = " ";
            string whereBblDt = " ";
            string whereAppDt = " ";
            string whereAppSysDt = " ";
            string orderBy = " ";
            string cBblDtSt = "";
            string cBblDtEnd = "";
            string cAppDtSt = "";
            string cAppDtEnd = "";
            string cAppSysDtSt = "";
            string cAppSysDtEnd = "";
            string whereBanc = " ";

            orderBy = " ORDER BY A.APP_NO ASC " + NL;

            if ((bblDtSt != "") && (bblDtEnd != ""))
            {
                cBblDtSt = " TO_DATE('" + manage.GetDateFomatEN(bblDtSt) + "','DD/MM/RRRR') ";
                cBblDtEnd = " TO_DATE('" + manage.GetDateFomatEN(bblDtEnd) + "','DD/MM/RRRR') ";
                whereBblDt = " AND " + NL +
                             " ( " + NL +
                             "    CASE " + NL +
                             "    WHEN (A.FLG_TYPE = 'C') THEN " + NL +
                             "      (SELECT MAX(H.IBRH_APP_DATE) FROM IS_BBL_REF_H H WHERE H.IBRH_APP_NO = A.APP_NO) " + NL +
                             "    ELSE " + NL +
                             "      (SELECT MAX(R.IBBL_APP_DATE) FROM IS_BBL_REF R WHERE R.IBBL_APP_NO = A.APP_NO) " + NL +
                             "    END " + NL +
                             " ) BETWEEN " + cBblDtSt + " AND " + cBblDtEnd + " " + NL;
            }
            else
            {

                whereBblDt = " ";
            }

            if ((appDtSt != "") && (appDtEnd != ""))
            {
                cAppDtSt = manage.GetStrDate(appDtSt);
                cAppDtEnd = manage.GetStrDate(appDtEnd);
                whereAppDt = " AND A.APP_DATE BETWEEN '" + cAppDtSt + "' AND '" + cAppDtEnd + "' " + NL;
            }
            else
            {
                whereAppDt = " ";

            }

            if ((appSysDtSt != "") && (appSysDtEnd != ""))
            {
                cAppSysDtSt = manage.GetStrDate(appSysDtSt);
                cAppSysDtEnd = manage.GetStrDate(appSysDtEnd);
                whereAppSysDt = " AND A.APPSYS_DATE BETWEEN '" + cAppSysDtSt + "' AND '" + cAppSysDtEnd + "' " + NL;
            }
            else
            {
                whereAppSysDt = " ";
            }

            if (depCode == "ALL")
            {
                whereDep = " ";
            }
            else
            {
                whereDep = " AND " + NL +
                           "( " + NL +
                           "  CASE " + NL +
                           "  WHEN (A.FLG_TYPE <> 'C') AND (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                           "  WHEN (A.FLG_TYPE <> 'C') AND (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                           "  WHEN (A.FLG_TYPE <> 'C') AND (A.MKG_TYPE = 'TEL') THEN 'E' " + NL +
                           "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION = 'BKK') THEN 'A' " + NL +
                           "  WHEN (A.FLG_TYPE <> 'C') AND (A.REGION <> 'BKK') THEN 'B' " + NL +
                           "  WHEN (A.FLG_TYPE = 'C') AND (A.POLNO IN(SELECT B.POLICY FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL)) THEN 'D' " + NL +
                           "  END " + NL +
                           ") = '" + depCode + "' " + NL;
            }

            string chkPlan = "";
            chkPlan = planCode.Substring(0, 1);
            if (planCode == "ALL")
            {
                wherePlan = " ";
            }
            else
            {
                if (chkPlan == "8" || chkPlan == "H")
                {
                    wherePlan = " AND A.POLNO = '" + planCode + "' " + NL;
                }
                else
                {
                    wherePlan = " AND A.PL_BLOCK||A.OPL_CODE||A.OPL_CODE2 = '" + planCode + "' " + NL;
                }
            }

            if (statusCode == "ALL")
            {
                whereStatus = " ";
            }
            else if (statusCode == "PO")
            {
                whereStatus = " AND A.STATUS IS NULL " + NL;
            }
            else
            {
                whereStatus = " AND A.STATUS = '" + statusCode + "' " + NL;
            }

            if (name == "")
            {
                whereName = " ";
            }
            else
            {
                whereName = " AND A.NAME LIKE '%" + name + "%' " + NL;
                orderBy = " ORDER BY A.NAME ASC " + NL;
                whereBblDt = " ";
                whereAppDt = " ";
                whereAppSysDt = " ";
            }

            if (appNo == "")
            {
                whereAppNo = " ";
            }
            else
            {
                whereAppNo = " AND A.APP_NO = '" + appNo + "'" + NL;
                whereBblDt = " ";
                whereAppDt = " ";
                whereAppSysDt = " ";
            }

            if (polNo == "")
            {
                wherePolNo = " ";
            }
            else
            {
                wherePolNo = " AND A.POLNO = '" + polNo + "' " + NL;
                whereBblDt = " ";
                whereAppDt = " ";
                whereAppSysDt = " ";
            }

            if ((summSt != "") && (summEnd != ""))
            {
                whereSumm = " AND A.SUMM BETWEEN " + summSt + " AND " + summEnd + " " + NL;
            }
            else
            {
                whereSumm = " ";
            }

            if (banc == "ALL")
            {
                whereBanc = " ";
            }
            else
            {
                whereBanc = " AND " + NL +
                            " ( " + NL +
                            "    CASE " + NL +
                            "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                            "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.OPL_CODE = A.OPL_CODE AND B.OPL_CODE2 = A.OPL_CODE2) " + NL +
                            "      ELSE " + NL +
                            "        (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = A.POLNO) " + NL +
                            "    END " + NL +
                            " ) = '" + banc + "' " + NL;
            }

            sql = "SELECT " + NL +
                  "/*( " + NL +
                  "    CASE " + NL +
                  "    WHEN ASCII(SUBSTR(A.APP_NO,0,1)) >= 65 THEN " + NL +
                  "        A.APP_NO " + NL +
                  "    ELSE " + NL +
                  "        LPAD(A.APP_NO,8,'0') " + NL +
                  "    END " + NL +
                  ") AS APP_NO */" + NL +
                  "A.APP_NO AS APP_NO " + NL +
                  ",A.POLNO,A.NAME,DECODE(B.IAPR_APP_NO,NULL,'A','B') AS KEY_IN,A.CARD_NO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SENDBY " + NL +
                  ",'วันที่นำส่งจากประกันชีวิต ' AS SEND_PLACE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT DISTINCT S.SEND_DT FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = A.POLNO AND S.SEND_FAIL IS NULL)) " + NL +
                  "      ELSE " + NL +
                  "        ( " + NL +
                  "          SELECT S.SEND_DT " + NL +
                  "          FROM FSD_GM_MAST_SEND S " + NL +
                  "          WHERE S.POLICY = A.POLNO " + NL +
                  "          AND S.CERT_NO = A.CERT_NO " + NL +
                  "          AND S.SEND_FAIL = 'N' " + NL +
                  "          AND S.SEQ = (SELECT MAX(S.SEQ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "          AND S.UPD_DT = (SELECT MAX(S.UPD_DT ) FROM FSD_GM_MAST_SEND S WHERE S.POLICY = A.POLNO AND S.CERT_NO = A.CERT_NO AND S.SEND_FAIL = 'N') " + NL +
                  "        ) " + NL +
                  "    END " + NL +
                  " ) AS SEND_DATE " + NL +
                  ",NVL(A.SUMM,0) AS SUMM,NVL(A.PREMIUM+A.PREM+A.RIDER,0) AS PREMIUM,NVL(A.CALPREM,0) AS CALPREM,A.P_MODE,A.PLAN " + NL +
                  ",(  " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.FLG_TYPE = 'A') AND (A.STATUS = 'IF') THEN " + NL +
                  "      'ส่งเสนอผู้พิจารณา' " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      (SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT T.TCT_DESC FROM TB_CONSTANT T WHERE T.TCT_CODE1 = 'AS' AND T.TCT_CODE2 = A.STATUS) IS NULL) AND (A.FLG_TYPE <> 'C') THEN " + NL +
                  "      'กำลังพิจารณา '||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS NOT NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      (SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) " + NL +
                  "    WHEN ((SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'LET_TYPE' AND C.CODE2 = A.STATUS) IS  NULL) AND (A.FLG_TYPE = 'C') THEN " + NL +
                  "      'กำลังพิจารณา'||DECODE(A.UNDER_WRITE,'A','(ADMIN)','U','(UNDERWRITE)') " + NL +
                  "    END " + NL +
                  " ) AS STATUS " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "    WHEN (A.BANK_ASS = 'Y') AND (A.FLG_TYPE <> 'C') THEN " +
                  "      (SELECT S.DESCRIPTION||'('||BBL_BRANCH||')'||'/'||(WEB_PKG.GET_REGION_BYAGENT(A.AGENT_CODE)) AS DESCRIPTION FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.BANK_ASS IS NULL) AND (A.FLG_TYPE = 'C') THEN " +
                  "      (SELECT S.DESCRIPTION||'('||BBL_BRANCH||')'||'/'||(WEB_PKG.GET_REGION_BYAGENT(A.AGENT_CODE)) AS DESCRIPTION FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = A.AGENT_CODE) " + NL +
                  "    WHEN (A.PL_BLOCK = 'S') THEN " + NL +
                  "      'สำนักงานใหญ่' " + NL +
                  "    ELSE " + NL +
                  "      (SELECT O.DESCRIPTION||'('||O.OFFICE||')' AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = A.OFFICE) " + NL +
                  "    END " + NL +
                  " ) AS BRANCH " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN (A.FLG_TYPE <> 'C') THEN " + NL +
                  "        (SELECT P.STATUS FROM P_POLICY P WHERE P.POLICY = A.POLNO) " + NL +
                  "      ELSE " + NL +
                  "        (SELECT M.STATUS FROM FSD_GM_MAST M WHERE M.POLICY = A.POLNO  AND M.APP_NO = A.APP_NO) " + NL +
                  "    END " + NL +
                  " ) AS POL_STATUS " + NL +
                  ",A.STATUS AS STATUS_CODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_MEMODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        WEB_PKG.GET_MEMOBANCDESC(A.APP_NO,A.POLNO) " + NL +
                  "    END " + NL +
                  " ) AS STATUS_MEMO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "      WHEN(A.FLG_TYPE <> 'C') THEN " + NL +
                  "        WEB_PKG.GET_CODESC(A.APP_NO) " + NL +
                  "      ELSE " + NL +
                  "        NVL(WEB_PKG.GET_COBANCDESC(A.APP_NO,A.POLNO),'ไม่ให้ความคุ้มครองเพิ่มเติมสำหรับการทุพพลภาพถาวรสิ้นเชิง') " + NL +
                  "    END " + NL +
                  " ) AS STATUS_CO,A.FLG_TYPE,A.AGENT_CODE,A.UPLINE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN (A.BANK_ASS = 'Y') THEN 'D' " + NL +
                  "        WHEN (A.PL_BLOCK = 'S') THEN 'C' " + NL +
                  "        WHEN (A.REGION = 'BKK') THEN 'A' " + NL +
                  "        WHEN (A.REGION <> 'BKK') THEN 'B' " + NL +
                  "    END " + NL +
                  " ) AS FLG_DEP " + NL +
                  "FROM WEB_APP_ALL_NEW A,IS_APP_REGION B " + NL +
                  "WHERE A.APP_NO = B.IAPR_APP_NO(+) " + NL;

            sql = sql + whereDep + wherePlan + whereStatus + whereName + whereAppNo + wherePolNo + whereSumm + whereBblDt + whereAppDt + whereAppSysDt + whereBanc + orderBy;
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyBanc(string stDate, string endDate)
        {
            string sql = "";
            string cStDate = "";
            string cEndDate = "";
            cStDate = "TO_DATE('" + manage.GetDateFomatEN(stDate) + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            cEndDate = "TO_DATE('" + manage.GetDateFomatEN(endDate) + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            sql = "SELECT NVL(SUM(DECODE(TB.BANC_TYPE,'G',1,0)),0) AS AMOUNT_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'G',TB.SUMM,0)),0) AS SUMM_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'G',TB.PREMIUM,0)),0) AS PREMIUM_GAIN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',1,0)),0) AS AMOUNT_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',TB.SUMM,0)),0) AS SUMM_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'P',TB.PREMIUM,0)),0) AS PREMIUM_PLAN " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',1,0)),0) AS AMOUNT_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',TB.SUMM,0)),0) AS SUMM_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'H',TB.PREMIUM,0)),0) AS PREMIUM_HOME " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',1,0)),0) AS AMOUNT_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',TB.SUMM,0)),0) AS SUMM_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'C',TB.PREMIUM,0)),0) AS PREMIUM_CREDIT " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',1,0)),0) AS AMOUNT_HL " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',TB.SUMM,0)),0) AS SUMM_HL " + NL +
                  ",NVL(SUM(DECODE(TB.BANC_TYPE,'E',TB.PREMIUM,0)),0) AS PREMIUM_HL " + NL +
                  "FROM " + NL +
                  "( " + NL +
                  "    SELECT P.POLICY,P.SUMM,(P.PREMIUM1 + P.PREMIUM2) AS PREMIUM " + NL +
                  "    ,( " + NL +
                  "        CASE " + NL +
                  "            WHEN P.FLG_TYPE = 'B' THEN " + NL +
                  "                (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) " + NL +
                  "            ELSE " + NL +
                  "                (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE  B.PL_CODE = P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) " + NL +
                  "        END " + NL +
                  "     ) AS BANC_TYPE " + NL +
                  "    FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "    WHERE  P.POLICY_TYPE = 'D' " + NL +
                  "    AND P.POLICY_DT BETWEEN " + cStDate + " AND " + cEndDate + " " + NL +
                  ") TB ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetCleanCase(string startPolDT, string endPolDT, string typeDep, string branch, string bancType, string planType)
        {
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            string sql = "";
            string WhereBranch = "";
            string whereBanc = " ";
            string wherePlan = " ";
            StrStartPolDT = manage.GetDateFomatEN(startPolDT);
            StrEndPolDT = manage.GetDateFomatEN(endPolDT);
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR','nls_calendar=''gregorian''')";
            if (branch == "00")
            {
                WhereBranch = "";
            }
            else
            {
                if (typeDep == "D")
                {
                    WhereBranch = " AND P.SALE_AGENT = '" + branch + "' " + NL;
                }
                else
                {
                    WhereBranch = " AND P.APP_OFC = '" + branch + "' " + NL;
                }
            }

            if (bancType == "")
            {
                whereBanc = " ";
            }
            else if (bancType == "1")
            {
                whereBanc = " AND (P.FLG_TYPE = 'A' OR (P.FLG_TYPE = 'B' AND P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE = 'E')))  " + NL;
            }
            else if (bancType == "2")
            {
                whereBanc = " AND P.FLG_TYPE = 'B' AND   P.POLICY IN (SELECT PLAN_BANC.POLICY FROM ZTB_PLAN_BANC PLAN_BANC WHERE PLAN_BANC.POLICY IS NOT NULL AND PLAN_BANC.BANC_TYPE IN ('C','H')) " + NL;
            }

            if (planType == "")
            {
                wherePlan = " ";
            }
            else if ((planType == "G") || (planType == "P"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE  B.PL_CODE = P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) = '" + planType + "' " + NL;
            }
            else if ((planType == "H") || (planType == "C") || (planType == "E"))
            {
                wherePlan = " AND (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) = '" + planType + "' " + NL;
            }



            sql = "SELECT " + NL +
                  "NVL(SUM( " + NL +
                  "    CASE " + NL +
                  "      WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "        0 " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  "),0) AS AMOUNT_UNCLEAN, " + NL +
                  "NVL(SUM( " + NL +
                  "    CASE " + NL +
                  "      WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "        0 " + NL +
                  "      WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "        1 " + NL +
                  "      ELSE " + NL +
                  "        0 " + NL +
                  "    END " + NL +
                  "),0) AMOUNT_CLEAN, " + NL +
                  "NVL(SUM( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) > 0 THEN " + NL +
                  "      /*POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(TB.APPSYS_DATE,'DD/MM/RRRR'),TO_DATE(TO_CHAR(TB.SEND_DATE,'DD/MM/')||(TO_CHAR(TB.SEND_DATE,'RRRR')+543),'DD/MM/RRRR')) */" + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),TB.SEND_DATE)" + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  "),0) AS AMOUNT_DAY_UNCLEAN, " + NL +
                  "NVL(SUM( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    WHEN (DECODE(TB.AP,NULL,0,1)+DECODE(TB.CO,NULL,0,1)+DECODE(TB.MO,NULL,0,1)) = 0 THEN " + NL +
                  "      /*POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(TB.APPSYS_DATE,'DD/MM/RRRR'),TO_DATE(TO_CHAR(TB.SEND_DATE,'DD/MM/')||(TO_CHAR(TB.SEND_DATE,'RRRR')+543),'DD/MM/RRRR')) */" + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),TB.SEND_DATE)" + NL +
                  "    ELSE " + NL +
                  "      0 " + NL +
                  "  END " + NL +
                  "),0) AS AMOUNT_DAY_CLEAN, " + NL +
                  "/*NVL(COUNT(*),0) AS AMOUNT_TOTAL,*/ " + NL +
                  "NVL(SUM( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    ELSE " + NL +
                  "      1 " + NL +
                  "  END " + NL +
                  "),0)AS AMOUNT_TOTAL , " + NL +
                  "NVL(SUM( " + NL +
                  "  CASE " + NL +
                  "    WHEN TB.SEND_DATE IS NULL THEN " + NL +
                  "      0 " + NL +
                  "    ELSE " + NL +
                  "      /*POLICY.WEB_PKG.get_workingdaykpi(TO_DATE(TB.APPSYS_DATE,'DD/MM/RRRR'),TO_DATE(TO_CHAR(TB.SEND_DATE,'DD/MM/')||(TO_CHAR(TB.SEND_DATE,'RRRR')+543),'DD/MM/RRRR'))*/" + NL +
                  "      POLICY.WEB_PKG.get_workingdaykpi(DECODE(TB.APPSYS_DATE,NULL,NULL,TO_DATE(SUBSTR(TB.APPSYS_DATE,1,2)||'/'||SUBSTR(TB.APPSYS_DATE,4,2)||'/'||(SUBSTR(TB.APPSYS_DATE,7)-543),'DD/MM/RRRR')),TB.SEND_DATE)" + NL +
                  "  END " + NL +
                  "),0)AS AMOUNT_DAY_TOTAL " + NL +
                  "FROM " + NL +
                  "( " + NL +
                  "SELECT P.FLG_TYPE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (SELECT SUBSTR(B.IMB_APPSYS_DATE,5,2)||'/'||SUBSTR(B.IMB_APPSYS_DATE,3,2)||'/25'||SUBSTR(B.IMB_APPSYS_DATE,1,2) FROM IS_APPM_BSC B WHERE B.IMB_APP_NO = P.APP_NO) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT TO_CHAR(A.APP_DT,'DD/MM/')||(TO_CHAR(A.APP_DT,'RRRR')+543) FROM FSD_GM_APP A WHERE A.POLICY =  P.POLICY AND A.APP_NO = P.APP_NO)  " + NL +
                  "    END " + NL +
                  " ) AS APPSYS_DATE " + NL +
                  ",( " + NL +

                   "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "           (SELECT S.SEND_DT FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEQ = 1 AND ROWNUM = 1)" + NL +
                  "        WHEN INSTALL_DT >= TO_DATE('06/09/2016','DD/MM/RRRR','nls_calendar=''gregorian''') THEN  " + NL +
                  "           (SELECT S.SEND_DT FROM P_POLICY_SENDING S,P_APPL_ID APL WHERE S.POLICY_ID = APL.POLICY_ID AND APL.APP_NO=P.APP_NO AND S.SEQ = 1 AND ROWNUM=1 ) " + NL +
                  "        ELSE " + NL +
                  "           (SELECT S.SEND_DT FROM FSD_GM_MAST_SEND S WHERE  S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  "    END " + NL +
                  //"    CASE " + NL +
                  //"        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  //"           (SELECT S.SEND_DT FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEQ = 1 AND ROWNUM = 1)  " + NL +
                  //"        ELSE " + NL +
                  //"           (SELECT S.SEND_DT FROM FSD_GM_MAST_SEND S WHERE S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.SEQ = 1 AND ROWNUM = 1) " + NL +
                  //"     END " + NL +
                  " ) AS SEND_DATE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (WEB_PKG.GET_STSDATE(P.APP_NO,'AP')) " + NL +
                  "        ELSE " + NL +
                  "            WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'AP') " + NL +
                  "    END " + NL +
                  "  ) AS AP " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (WEB_PKG.GET_STSDATE(P.APP_NO,'CO')) " + NL +
                  "        ELSE " + NL +
                  "            WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'CO') " + NL +
                  "    END " + NL +
                  "  ) AS CO " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (WEB_PKG.GET_STSDATE(P.APP_NO,'MO')) " + NL +
                  "        ELSE " + NL +
                  "            WEB_PKG.GET_STSDATE_BANC(P.POLICY,P.APP_NO,'MO') " + NL +
                  "    END " + NL +
                  "  ) AS MO " + NL +
                  "FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "WHERE  P.POLICY_TYPE = '" + typeDep + "' " + NL +
                  "AND P.POLICY_DT BETWEEN " + StrStartPolDTO + " AND " + StrEndPolDTO + " " + NL + WhereBranch + NL + whereBanc + NL + wherePlan + NL + ") TB ";
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicyBancDetail(string polNo, string certNo)
        {
            string sql = "";
            sql = "SELECT LPAD(M.AGENT,8,'0') AS AGENT,M.POLICY,M.CERT_NO,M.STATUS AS STATUS_CODE " + NL +
                  ",(SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'STATUS' AND C.CODE2 = M.STATUS) AS STATUS_DESC " + NL +
                  ",(SELECT B.DESCRIPTION FROM ZTB_PLAN_BANC B WHERE B.POLICY = M.POLICY) AS PLAN_DESC " + NL +
                  ",N.NAME,N.IDCARD_NO,DECODE(N.SEX,'0','หญิง','1','ชาย') AS SEX,TO_CHAR(N.BIRTH_DT,'DD/MM/')||(TO_CHAR(N.BIRTH_DT,'RRRR')+543) AS BIRTH_DT " + NL +
                  ",/*POLICY.ORD_PRODUCT.FCL_AGE(N.BIRTH_DT,SYSDATE,SYSDATE)*/M.ENTRY_AGE AS AGE " + NL +
                  ",N.ADDR1||' '||N.ADDR2||' '||N.ADDR3||' '||N.POST_CODE AS ADDRESS,N.TEL_NO " + NL +
                  ",M.INFSUMM,M.SUMM,M.PAY_YR||'/'||M.PAY_LT AS YRLT,M.ASS_TERM,M.PAY_TERM " + NL +
                  ",(SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'P_MODE' AND C.CODE2 = LPAD(M.FMODE,2,'0')) AS P_MODE " + NL +
                  ",TO_CHAR(M.ENTRY_DT,'DD/MM/')||(TO_CHAR(M.ENTRY_DT,'RRRR')+543) AS ENTRY_DT " + NL +
                  ",TO_CHAR(M.MAT_DT,'DD/MM/')||(TO_CHAR(M.MAT_DT,'RRRR')+543) AS MAT_DT " + NL +
                  ",TO_CHAR(M.NXTDUE_DT,'DD/MM/')||(TO_CHAR(M.NXTDUE_DT,'RRRR')+543) AS NXTDUE_DT " + NL +
                  ",TO_CHAR(M.EFF_DT,'DD/MM/')||(TO_CHAR(M.EFF_DT,'RRRR')+543) AS EFF_DT " + NL +
                  ",NVL(M.PREMIUM,0) AS PREMIUM,NVL(M.XPREMIUM,0) AS XPREMIUM " + NL +
                  ",TO_NUMBER(N.OCC_CODE) AS OCC_CODE " + NL +
                  ",TO_CHAR(M.LASTPAY_DT,'DD/MM/')||(TO_CHAR(M.LASTPAY_DT,'RRRR')+543) AS LASTPAY_DT " + NL +
                  "FROM FSD_GM_MAST M,FSD_GM_NAME N " + NL +
                  "WHERE M.CUST_NO = N.CUST_NO " + NL +
                  "AND M.POLICY = '" + polNo + "' " + NL +
                  "AND M.CERT_NO = " + certNo + " ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetStatusPolicy()
        {
            string sql = "";
            sql = "SELECT TB.DESCRIPTION,TB.CODE2 " +
                  "FROM " +
                  "( " +
                  "    SELECT 'เลือกสถานะทั้งหมด' AS DESCRIPTION,'ALL' AS CODE2,'A' AS FLG " +
                  "    FROM DUAL " +
                  "    UNION ALL " +
                  "    SELECT C.DESCRIPTION,C.CODE2,'B' AS FLG " +
                  "    FROM ZTB_CONSTANT2 C " +
                  "    WHERE C.COL_NAME = 'STATUS' " +
                  ") TB " +
                  "ORDER BY TB.FLG,TB.CODE2 ASC ";

            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }
        public DataTable GetPolicySearch(string dep, string plan, string banc, string status, string name, string polNo, string summSt, string summEt, string policyDateSt, string policyDateEnd, string installDateSt, string installDateEnd)
        {
            string sql = "";
            string cPolicyDateSt = "";
            string cPolicyDateEnd = "";
            string wherePolicyDate = " ";
            string cInstallDateSt = "";
            string cInstallDateEnd = "";
            string whereInstallDate = " ";
            string whereDep = " ";
            string whereStatus = " ";
            string whereSumm = " ";
            string chkPlan = "";
            string wherePlan = " ";
            string whereBanc = " ";
            string wherePolicy = " ";
            string whereName = " ";

            if ((policyDateSt != "") && (policyDateEnd != ""))
            {
                cPolicyDateSt = manage.GetDateFomatEN(policyDateSt);
                cPolicyDateEnd = manage.GetDateFomatEN(policyDateEnd);
                wherePolicyDate = " AND P.POLICY_DT BETWEEN TO_DATE('" + cPolicyDateSt + "','DD/MM/RRRR') AND TO_DATE('" + cPolicyDateEnd + "','DD/MM/RRRR') " + NL;
            }
            else
            {
                wherePolicyDate = " ";
            }

            if ((installDateSt != "") && (installDateEnd != ""))
            {
                cInstallDateSt = manage.GetDateFomatEN(installDateSt);
                cInstallDateEnd = manage.GetDateFomatEN(installDateEnd);
                whereInstallDate = " AND P.INSTALL_DT BETWEEN TO_DATE('" + cInstallDateSt + "','DD/MM/RRRR') AND TO_DATE('" + cInstallDateEnd + "','DD/MM/RRRR') " + NL;
            }
            else
            {
                whereInstallDate = " ";
            }

            if ((summSt != "") && (summEt != ""))
            {
                whereSumm = " AND P.SUMM BETWEEN " + summSt + " AND " + summEt + " " + NL;
            }
            else
            {
                whereSumm = " ";
            }

            if (dep == "ALL")
            {
                whereDep = " ";
            }
            else
            {
                whereDep = " AND P.POLICY_TYPE = '" + dep + "' " + NL;
            }

            if (status == "ALL")
            {
                whereStatus = " ";
            }
            else
            {
                whereStatus = " AND P.STATUS = '" + status + "' " + NL;
            }

            chkPlan = plan.Substring(0, 1);
            if (plan == "ALL")
            {
                wherePlan = " ";
            }
            else
            {
                if (chkPlan == "8" || chkPlan == "H")
                {
                    wherePlan = " AND P.POLICY = '" + plan + "' " + NL;
                }
                else
                {
                    wherePlan = " AND P.PL_BLOCK||P.PL_CODE||P.PL_CODE2 = '" + plan + "' " + NL;
                }
            }

            if (banc == "ALL")
            {
                whereBanc = " ";
            }
            else
            {
                whereBanc = " AND " + NL +
                            " ( " + NL +
                            "    CASE " + NL +
                            "         WHEN P.FLG_TYPE = 'A' THEN " + NL +
                            "            (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.PL_CODE =  P.PL_CODE AND B.PL_CODE2 = P.PL_CODE2) " + NL +
                            "         ELSE " + NL +
                            "            (SELECT B.BANC_TYPE FROM ZTB_PLAN_BANC B WHERE B.POLICY = P.POLICY) " + NL +
                            "    END " + NL +
                            " ) = '" + banc + "' " + NL;
            }

            if (polNo == "")
            {
                wherePolicy = " ";
            }
            else
            {
                wherePolicy = " AND DECODE(P.FLG_TYPE,'A',P.POLICY,'B',P.CERT_NO) = '" + polNo + "' " + NL;
                wherePolicyDate = " ";
                whereInstallDate = " ";
            }
            if (name == "")
            {
                whereName = " ";
            }
            else
            {
                whereName = " AND " + NL +
                            " ( " + NL +
                            "    CASE " + NL +
                            "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                            "            (SELECT N.NAME||'  '||N.SURNAME AS NAME FROM P_NAME N WHERE N.NAMECODE = P.NAMECODE) " + NL +
                            "        ELSE " + NL +
                            "            (SELECT N.NAME FROM FSD_GM_NAME N WHERE N.CUST_NO = P.CUST_NO) " + NL +
                            "    END " + NL +
                            "  ) LIKE '%" + name + "%' " + NL;
                wherePolicyDate = " ";
                whereInstallDate = " ";
            }
            sql = "SELECT P.POLICY,P.CERT_NO,P.APP_NO,DECODE(P.FLG_TYPE,'A',P.POLICY,'B',P.POLICY||'/'||P.CERT_NO) AS POLNO,P.FLG_TYPE,P.POLICY_TYPE " + NL +
                  " ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT  DISTINCT MAX(DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)))  FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY))  " + NL +
                  "          ELSE " + NL +
                  "             (SELECT DECODE(S.SEND_DT,NULL,NULL,SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),1,2)||'/'||SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),4,2)||'/'||(SUBSTR(TO_CHAR(S.SEND_DT,'DD/MM/RRRR'),7)+543)) FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.UPD_DT = (SELECT MAX(S.UPD_DT) FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO))  " + NL +
                  "       END " + NL +
                  "   ) AS SEND_DATE " + NL +
                  " ,( " + NL +
                  "      CASE " + NL +
                  "          WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "             (SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY AND S.SEND_FAIL IS NULL AND S.SEQ = (SELECT MAX(S.SEQ) FROM P_POLICY_SENDING S WHERE S.POLICY = P.POLICY)) " + NL +
                  "          ELSE " + NL +
                  "             (SELECT DECODE(S.SEND_BY,'MC','ส่งไปรษณีย์ให้กับลูกค้า','MO','ส่งไปรษณีย์ให้กับสาขา','MB','ส่งไปรษณีย์ไปยังสาขาธนาคาร','SO','ส่งโดย Messenger ไปยังสาขา','UN','รับไปโดยหน่วย','BH','By Hand') FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO AND S.UPD_DT = (SELECT MAX(S.UPD_DT) FROM FSD_GM_MAST_SEND S WHERE S.SEND_FAIL = 'N' AND S.POLICY = P.POLICY AND S.CERT_NO = P.CERT_NO)) " + NL +
                  "       END " + NL +
                  "   ) AS SEND_BY " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (SELECT N.NAME||'  '||N.SURNAME AS NAME FROM P_NAME N WHERE N.NAMECODE = P.NAMECODE) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT N.NAME FROM FSD_GM_NAME N WHERE N.CUST_NO = P.CUST_NO) " + NL +
                  "    END " + NL +
                  " ) AS NAME " + NL +
                  ",P.SUMM,(P.PREMIUM1+P.PREMIUM2) AS PREMIUM " + NL +
                  ",(SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'P_MODE' AND C.CODE2 = LPAD(P.P_MODE,2,'0')) AS P_MODE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (SELECT Z.BLA_TITLE AS TITLE FROM ZTB_PLAN Z  WHERE Z.PL_BLOCK = P.PL_BLOCK  AND Z.PL_CODE = P.PL_CODE AND Z.PL_CODE2 = P.PL_CODE2) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT B.DESCRIPTION FROM ZTB_PLAN_BANC B WHERE B.POLICY IS NOT NULL AND B.POLICY = P.POLICY) " + NL +
                  "    END " + NL +
                  " ) AS TITLE " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.FLG_TYPE = 'A' THEN " + NL +
                  "            (WEB_PKG.GET_STATUS(P.POLICY,P.STATUS,P.POL_YR,P.POL_LT)) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT C.DESCRIPTION FROM ZTB_CONSTANT2 C WHERE C.COL_NAME = 'STATUS' AND C.CODE2 = P.STATUS) " + NL +
                  "    END " + NL +
                  " ) AS STATUS " + NL +
                  ",( " + NL +
                  "    CASE " + NL +
                  "        WHEN P.POLICY_TYPE <> 'D' THEN " + NL +
                  "            (SELECT Z.DESCRIPTION  FROM  ZTB_OFFICE Z  WHERE Z.OFFICE = DECODE(P.APP_OFC,'สสง','สนญ',P.APP_OFC)) " + NL +
                  "        ELSE " + NL +
                  "            (SELECT 'ธนาคารกรุงเทพ '||S.DESCRIPTION||'('||BBL_BRANCH||')'||'/'||(WEB_PKG.GET_REGION_BYAGENT(P.SALE_AGENT)) AS DESCRIPTION FROM GBBL_STRUCT S WHERE S.BBL_AGENTCODE = P.SALE_AGENT) " + NL +
                  "    END " + NL +
                  " ) AS BRANCH " + NL +
                  "FROM POLICY_ALL_MAIN_NEW P " + NL +
                  "WHERE 1 = 1 " + NL;

            sql = sql + wherePolicyDate + whereInstallDate + whereSumm + whereDep + whereStatus + wherePlan + whereBanc + wherePolicy + whereName;
            OracleDataAdapter da = new OracleDataAdapter(sql, oConn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

    }
}