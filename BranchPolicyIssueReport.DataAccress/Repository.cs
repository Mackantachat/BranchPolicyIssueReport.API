using BranchPolicyIssueReport.DataContract;
using DataAccessUtility;
using ITUtility;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using WebApp2;

namespace BranchPolicyIssueReport.DataAccress
{
    public class Repository: RepositoryBaseManagedCore
    {
        public Repository(string ConnectionName) : base(ConnectionName)
        {

        }
        public Repository(OracleConnection connection) : base(connection)
        {

        }

        public PolicyOfficeModel[] GetPolicyOffice(string startPolDT, string endPolDT, string monthVal)
        {
            PolicyOfficeModel[] returnValue = null;
            string StrStartPolDT = "";
            string StrEndPolDT = "";
            string StrStartPolDTO = "";
            string StrEndPolDTO = "";
            StringBuilder sql = new StringBuilder();
            string WhereMonth = "";
            if (monthVal == "00")
            {
                WhereMonth = " ";
            }
            else
            {
                WhereMonth = " AND TO_CHAR(B.POLICY_DT,'MM') = :monthVal";
            }
            //StrStartPolDT = startPolDT;
            //StrEndPolDT = endPolDT;
            StrStartPolDT = manage.GetDateFomatEN(startPolDT) + " 00:00:00";
            StrEndPolDT = manage.GetDateFomatEN(endPolDT) + " 23:59:59";
            StrStartPolDTO = "TO_DATE('" + StrStartPolDT + "','DD/MM/RRRR')";
            StrEndPolDTO = "TO_DATE('" + StrEndPolDT + "','DD/MM/RRRR')";

            sql.Append("select ROWNUM AS Id, POLICY_MONTH , POLICY_YEAR , AMOUNT_CE , SUMM_CE , PREMIUM_CE ,AMOUNT_EA ,SUMM_EA , PREMIUM_EA , AMOUNT_NO , SUMM_NO , PREMIUM_NO , AMOUNT_NE , SUMM_NE , PREMIUM_NE FROM (" +
                    "SELECT TO_CHAR(B.POLICY_DT,'MM') AS POLICY_MONTH,TO_CHAR(B.POLICY_DT,'YYYY') AS POLICY_YEAR \n" +
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
                  " AND B.POLICY_DT  BETWEEN TO_DATE(:StrStartPolDT,'DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''') AND TO_DATE(:StrEndPolDT,'DD/MM/RRRR HH24:MI:SS','nls_calendar=''gregorian''')" + " " + WhereMonth + "\n" +
                  " GROUP BY TO_CHAR(B.POLICY_DT,'MM'),TO_CHAR(B.POLICY_DT,'YYYY') \n" +
                  " ORDER BY TO_CHAR(B.POLICY_DT,'MM') ASC ) \n");

            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;
                cmd.Parameters.Add(new OracleParameter("StrStartPolDT", StrStartPolDT));
                cmd.Parameters.Add(new OracleParameter("StrEndPolDT", StrEndPolDT));
                if (monthVal != "00")
                {
                    cmd.Parameters.Add(new OracleParameter("monthVal", monthVal));
                }
                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyOfficeModel>().ToArray();
                    }
                }
            }
            return returnValue;
        }

        public PolicyOfficeOfMonthModel[] GetPolicyOfficeOfMonth(string txtregion, string startPolDT, string endPolDT)
        {
            PolicyOfficeOfMonthModel[] returnValue = null;

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

            if (!txtregion.Equals("ALL"))
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                //condition = condition + "D.REGION_CODE = " + Utility.SQLValueString(txtregion) + "\n"; //comment by Korn 
                condition = condition + "S.SECTION_CODE = " + Utility.SQLValueString(txtregion) + "\n"; //add by Korn
            }


            sql = "  SELECT ROWNUM AS Id,SECTION_CODE ,SECTION_NAME, NAME,SURNAME,SEQ ,AMOUNT_POLICY , SUMM_POLICY , PREMIUM_POLICY FROM (SELECT S.SECTION_CODE \n" +
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
                    " Order by S.SEQ ASC ) ";

            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;
                cmd.Parameters.Add(new OracleParameter("StrStartPolDT", StrStartPolDT));
                cmd.Parameters.Add(new OracleParameter("StrEndPolDT", StrEndPolDT));
             
                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyOfficeOfMonthModel>().ToArray();
                    }
                }
            }
            return returnValue;
        }

        public DataTable GetRegion()
        {
            string sql = "";

            //sql = "select * from ztb_bla_region order by region_code asc"; //comment by Korn
            sql = "select * from ZTB_BLA_SECTION order by seq asc"; //add by Korn
            OracleDataAdapter da = new OracleDataAdapter(sql, connection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public DataTable GetRegionSub(string secCode) //add by Korn
        {
            string sql = "";

            sql = "select * from ztb_bla_region2 where SECTION_CODE = '" + secCode + "' order by seq asc";
            OracleDataAdapter da = new OracleDataAdapter(sql, connection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;

        }

        public DataTable GetBkkRegionDdl(string Dep, string Ofc)
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
                      " where ztbofc.OFFICE=ofc.OFFICE  " + condition + conofc + " \n " +
                      " and (ztbofc.OFFICE not like 'สสง') \n" +
                      " ORDER BY ofc.DESCRIPTION ASC ";


            OracleDataAdapter da = new OracleDataAdapter(sql, connection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }

        public PolicyOfficeOfRegionModel[] GetPolicyOfficeOfRegion(string txtsection, string txtregion, string startPolDT, string endPolDT) 
        {
            PolicyOfficeOfRegionModel[] returnValue = null;
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

            if (!txtsection.Equals("ALL"))
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                condition = condition + "S.SECTION_CODE = " + Utility.SQLValueString(txtsection) + "\n";
            }


            if (!txtregion.Equals("ALL"))
            {
                if (condition != "")
                    condition = condition + "   AND ";
                else
                    condition = condition + "AND \n    ";

                condition = condition + "E.REGION_CODE = " + Utility.SQLValueString(txtregion) + "\n";
            }

            sql = "SELECT ROWNUM AS Id, SECTION_CODE, SECTION_NAME, REGION_CODE, REGION_NAME, NAME , SURNAME, SEQ, AMOUNT_POLICY, SUMM_POLICY, PREMIUM_POLICY  FROM (SELECT S.SECTION_CODE \n" +
                    " ,S.SECTION_NAME \n" +
                    " ,E.REGION_CODE \n" +
                    " ,E.REGION_NAME \n " +
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
                    " Order by E.SEQ ASC) ";

            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;
                cmd.Parameters.Add(new OracleParameter("StrStartPolDT", StrStartPolDT));
                cmd.Parameters.Add(new OracleParameter("StrEndPolDT", StrEndPolDT));

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyOfficeOfRegionModel>().ToArray();
                    }
                }
            }
            return returnValue;
        }

        public PolicyOfficeOfDepModel[] GetPolicyOfficeOfDep(string section, string region, string office, string startPolDT, string endPolDT)  //add by Korn
        {
            PolicyOfficeOfDepModel[] returnValue = null;
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


            if (!section.Equals("ALL")) //add new 
            {
                if (txtsection != "")
                    txtsection = txtsection + "   AND ";
                else
                    txtsection = txtsection + "AND \n    ";

                txtsection = txtsection + "S.SECTION_CODE = " + Utility.SQLValueString(section) + "\n";
            }
            if (!section.Equals("ALL")) //add new 
            {
                if (txtsection2 != "")
                    txtsection2 = txtsection2 + "   AND ";
                else
                    txtsection2 = txtsection2 + "AND \n    ";

                txtsection2 = txtsection2 + "SEC.SECTION_CODE = " + Utility.SQLValueString(section) + "\n";
            }

            if (!region.Equals("ALL"))
            {
                if (txtregion != "")
                    txtregion = txtregion + "   AND ";
                else
                    txtregion = txtregion + "AND \n    ";

                txtregion = txtregion + "D.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }

            if (!region.Equals("ALL"))
            {
                if (txtregion2 != "")
                    txtregion2 = txtregion2 + "   AND ";
                else
                    txtregion2 = txtregion2 + "AND \n    ";

                txtregion2 = txtregion2 + "TB1.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }


            if (!office.Equals("ALL"))
            {
                if (txtoffice != "")
                    txtoffice = txtoffice + "   AND ";
                else
                    txtoffice = txtoffice + "AND \n    ";

                txtoffice = txtoffice + "D.OFFICE = " + Utility.SQLValueString(office) + "\n";
            }
            if (office.Equals("ALL"))
            {
                if (txtoffice2 != "")
                    txtoffice2 = txtoffice2 + "    ";
                else
                    txtoffice2 = txtoffice2 + "\n    ";

                txtoffice2 = txtoffice2 +
                         " SELECT SUM(TB1.SEQ) AS SUM_TB1_SEQ ,decode(TB3.OFFICE,'สสง','สนญ',TB3.OFFICE) AS OFFICE,DECODE(TB3.DESCRIPTION,'สำนักงานใหญ่(สสง)','สำนักงานใหญ่',TB3.DESCRIPTION) AS DESCRIPTION \n" +
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
            if (office.Equals("ALL"))
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
                    "     FROM  PW_LIFE_SUMM A \n" +
                    "           ,P_INSTALL B \n" +
                    "           ,P_AGENT C \n" +
                    "           ,ZTB_BLA_REGION_OFFICE2 D \n" +
                    "           ,ZTB_BLA_REGION2 E \n" +
                    "           ,ZTB_BLA_REGION_ID F \n" +
                    "           ,ZTB_USER G,ZTB_OFFICE H \n" +
                    "           ,P_POLICY_SENDING I \n" +
                    "           ,IS_APPM_BSC J \n" +
                    "           ,ZTB_BLA_SECTION S \n" +
                    "     WHERE A.CHANNEL_TYPE = 'OD' \n" +
                    "           AND A.PL_BLOCK = 'A' \n" +
                    "           AND A.POLICY_ID = B.POLICY_ID \n" +
                    "           AND A.POLICY_ID = C.POLICY_ID \n" +
                    "           AND B.APP_OFC = D.OFFICE \n" +
                    "           AND D.REGION_CODE = E.REGION_CODE \n" +
                    "           AND E.SECTION_CODE = S.SECTION_CODE --add \n" +
                    "           AND S.SECTION_CODE = F.SECTION_CODE --e \n" +
                    "           AND F.REGION_CODE is null  --a \n" +
                    "           AND F.N_USERID <> '000199' --a \n" +
                    "           AND F.N_USERID <> '000198' --a \n" +
                    "           AND F.N_USERID <> '001541' \n" +
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




            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;
                cmd.Parameters.Add(new OracleParameter("StrStartPolDT", StrStartPolDT));
                cmd.Parameters.Add(new OracleParameter("StrEndPolDT", StrEndPolDT));

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyOfficeOfDepModel>().ToArray();
                    }
                }
            }
            return returnValue;
        }

        public PolicyOfficeOfDateModel[] GetPolicyOfficeOfDate(string region, string office, string startPolDT, string endPolDT)
        {
            PolicyOfficeOfDateModel[] returnValue = null;
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
            if (!region.Equals("ALL"))
            {
                if (txtregion != "")
                    txtregion = txtregion + "   AND ";
                else
                    txtregion = txtregion + "AND \n    ";

                txtregion = txtregion + "D.REGION_CODE = " + Utility.SQLValueString(region) + "\n";
            }
            if (!office.Equals("ALL"))
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
            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;
                cmd.Parameters.Add(new OracleParameter("StrStartPolDT", StrStartPolDT));
                cmd.Parameters.Add(new OracleParameter("StrEndPolDT", StrEndPolDT));

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyOfficeOfDateModel>().ToArray();
                    }
                }
            }
            return returnValue;
        }


        public PolicyListModel[] GetPolicyList(string startPolDT, string endPolDT, string startInstallDT, string endInstallDT, string typeDep, string branch, string bancType, string planType, string sendType, string caseType)
        {
            PolicyListModel[] returnValue = null;
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
            string NL = "\n";

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
                  "  WHERE p.policy is not null " + condition + wherePolicyDate + whereInstallDate + WhereBranch + whereBanc + wherePlan + whereSend +
                  ") TB " + NL + whereCase + NL +
                  "ORDER BY TB.INSTALL_DT,TB.POLICY ASC ";

            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<PolicyListModel>().ToArray();
                    }
                }
            }
            return returnValue;
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
            OracleDataAdapter da = new OracleDataAdapter(sql, connection);
            DataTable dt = new DataTable();
            da.Fill(dt);
            return dt;
        }


        public CR_POLICY_REC GetPolicyDetail(string PolNo)
        {
            CR_POLICY_REC returnValue = null;
            OracleCommand com = new OracleCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_POLICY_DETAIL";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = connection;
            com.BindByName = true;
            com.Parameters.Add("POL_REF", OracleDbType.RefCursor, ParameterDirection.InputOutput);
            com.Parameters.Add("POLICY_IN", PolNo);
            using (DataTable dt = Utility.FillDataTable(com))
            {
                if (dt.Rows.Count > 0)
                {
                    returnValue = dt.AsEnumerable<CR_POLICY_REC>().FirstOrDefault();
                }
            }
            return returnValue;
        }

        public Saleco GetSalecoByPolCaseD(string polNo)
        {
            Saleco returnValue = null;
            string sql;
        
                sql = "SELECT S.DESCRIPTION " +
                      ",WEB_PKG.GET_SALECO_BYAGENT(S.BBL_AGENTCODE) AS SALECO_NAME " +
                      ",WEB_PKG.GET_MKG_BYAGENT(S.BBL_AGENTCODE) AS MKG_NAME " +
                      "FROM GBBL_STRUCT S " +
                      "WHERE S.BBL_AGENTCODE = (SELECT P.ASSIGN_AGENT FROM P_POLICY P WHERE P.POLICY = '" + polNo + "')";


            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<Saleco>().FirstOrDefault();
                    }
                }
            }
            return returnValue;
        }

        public Agent GetSalecoByPol(string polNo)
        {
            Agent returnValue = null;
            string sql;
            sql = "SELECT D.AGENTCODE,D.UPLINE " +
                       ",NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = D.AGENTCODE),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = D.AGENTCODE))||' (สังกัด '||(SELECT O.DESCRIPTION AS DESCRIPTION FROM ZTB_OFFICE O WHERE O.OFFICE = (NVL((SELECT OFFICE FROM GAG_DETAIL WHERE AGENTCODE = D.AGENTCODE),(SELECT OFFICE FROM GAG_INACTIVE WHERE AGENTCODE = D.AGENTCODE ))))||')' AS AGENT_NAME " +
                       ",NVL((SELECT N.NAME||'  '||N.SURNAME AS AGENT_NAME FROM GAG_NAME N WHERE N.AGENTCODE = D.UPLINE),(SELECT O.DESCRIPTION FROM ZTB_OFFICE O WHERE LPAD(O.OLD_OFFICE,8,'0') = D.UPLINE)) AS UPLINE_NAME " +
                       "FROM GAG_DETAIL D " +
                       "WHERE D.AGENTCODE = (SELECT P.ASSIGN_AGENT FROM P_POLICY P WHERE P.POLICY = '" + polNo + "')";

            using (OracleCommand cmd = connection.CreateCommand())
            {
                cmd.CommandType = CommandType.Text;
                cmd.BindByName = true;

                cmd.CommandText = sql.ToString();
                using (DataTable dt = Utility.FillDataTable(cmd))
                {
                    if (dt.Rows.Count > 0)
                    {
                        returnValue = dt.AsEnumerable<Agent>().FirstOrDefault();
                    }
                }
            }
            return returnValue;
        }

        public CR_PREMIUM_REC[] GetPolByPremium(string PolNo)
        {
            CR_PREMIUM_REC[] returnValue = null;
            OracleCommand com = new OracleCommand();
            com.CommandText = "WEB_PKG.DO_QUERY_PREMIUMBYPOLICY";
            com.CommandType = CommandType.StoredProcedure;
            com.Connection = connection;
            com.BindByName = true;
            com.Parameters.Add("PREMIUM_REF", OracleDbType.RefCursor, ParameterDirection.InputOutput);
            com.Parameters.Add("POLICY_IN", PolNo);
            using (DataTable dt = Utility.FillDataTable(com))
            {
                if (dt.Rows.Count > 0)
                {
                    returnValue = dt.AsEnumerable<CR_PREMIUM_REC>().ToArray();
                }
            }
            return returnValue;
        }
    }
}