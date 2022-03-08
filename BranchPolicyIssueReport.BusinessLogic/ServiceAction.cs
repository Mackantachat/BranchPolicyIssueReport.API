using BranchPolicyIssueReport.DataAccress;
using BranchPolicyIssueReport.DataContract;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using WebApp2;
using System.Linq;

namespace BranchPolicyIssueReport.BusinessLogic
{
    public class ServiceAction
    {
        private readonly string connectionString;
        public ServiceAction(string ConnectionName) => this.connectionString = ConnectionName;

        public PolicyOfficeModel[] GetPolicyOffice(RequestPolicyModel sPolicy) => GetPolicyOffice(sPolicy, (Repository)null);

        private PolicyOfficeModel[] GetPolicyOffice(RequestPolicyModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyOfficeModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = "01/01/" + sPolicy.Year;
                var txtEndPolDT = "31/12/" + sPolicy.Year;
                datas = repository.GetPolicyOffice(txtStartPolDT, txtEndPolDT, sPolicy.Month);
                PolicyOfficeModel lastRowSum = new PolicyOfficeModel();
                lastRowSum.POLICY_MONTH_NAME = "รวม";
                lastRowSum.POLICY_MONTH = "00";
                lastRowSum.POLICY_YEAR = sPolicy.Year.ConvertStrYearTHToEN();
                decimal sumAmountCE, sumSumCE, sumPremiumCE, sumAmountEA, sumSumEA, sumPremiumEA, sumAmountNE, sumSumNE, sumPremiumNE, sumAmountNO, sumSumNO, sumPremiumNO, sumAmountTotal, sumSumTotal, sumPremiumTotal;
                sumAmountCE = sumSumCE = sumPremiumCE = sumAmountEA = sumSumEA = sumPremiumEA = sumAmountNE = sumSumNE = sumPremiumNE = sumAmountNO = sumSumNO = sumPremiumNO = sumAmountTotal = sumSumTotal = sumPremiumTotal = 0;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lastRowSum.Id = (datas.Length + 1).ToString();
                        }
                        data.POLICY_MONTH_NAME = manage.GetThaiMonth(data.POLICY_MONTH);
                        var amountTotal = decimal.Parse(data.AMOUNT_CE) + decimal.Parse(data.AMOUNT_EA) + decimal.Parse(data.AMOUNT_NE) + decimal.Parse(data.AMOUNT_NO);
                        data.AmountTotal = amountTotal.ToString("n0");
                        var summTotal = decimal.Parse(data.SUMM_CE) + decimal.Parse(data.SUMM_EA) + decimal.Parse(data.SUMM_NE) + decimal.Parse(data.SUMM_NO);
                        data.SummTotal = summTotal.ToString("n0");
                        var premiunTotal = decimal.Parse(data.PREMIUM_CE) + decimal.Parse(data.PREMIUM_EA) + decimal.Parse(data.PREMIUM_NE) + decimal.Parse(data.PREMIUM_NO);
                        data.PremiumTotal = premiunTotal.ToString("n0");

                        data.AMOUNT_CE = decimal.Parse(data.AMOUNT_CE).ToString("n0");
                        data.AMOUNT_EA = decimal.Parse(data.AMOUNT_EA).ToString("n0");
                        data.AMOUNT_NE = decimal.Parse(data.AMOUNT_NE).ToString("n0");
                        data.AMOUNT_NO = decimal.Parse(data.AMOUNT_NO).ToString("n0");
                        data.PREMIUM_CE = decimal.Parse(data.PREMIUM_CE).ToString("n0");
                        data.PREMIUM_EA = decimal.Parse(data.PREMIUM_EA).ToString("n0");
                        data.PREMIUM_NE = decimal.Parse(data.PREMIUM_NE).ToString("n0");
                        data.PREMIUM_NO = decimal.Parse(data.PREMIUM_NO).ToString("n0");
                        data.SUMM_CE = decimal.Parse(data.SUMM_CE).ToString("n0");
                        data.SUMM_EA = decimal.Parse(data.SUMM_EA).ToString("n0");
                        data.SUMM_NE = decimal.Parse(data.SUMM_NE).ToString("n0");
                        data.SUMM_NO = decimal.Parse(data.SUMM_NO).ToString("n0");

                        sumAmountCE += decimal.Parse(data.AMOUNT_CE);
                        sumSumCE += decimal.Parse(data.SUMM_CE);
                        sumPremiumCE += decimal.Parse(data.PREMIUM_CE);

                        sumAmountEA += decimal.Parse(data.AMOUNT_EA);
                        sumSumEA += decimal.Parse(data.SUMM_EA);
                        sumPremiumEA += decimal.Parse(data.PREMIUM_EA);

                        sumAmountNE += decimal.Parse(data.AMOUNT_NE);
                        sumSumNE += decimal.Parse(data.SUMM_NE);
                        sumPremiumNE += decimal.Parse(data.PREMIUM_NE);

                        sumAmountNO += decimal.Parse(data.AMOUNT_NO);
                        sumSumNO += decimal.Parse(data.SUMM_NO);
                        sumPremiumNO += decimal.Parse(data.PREMIUM_NO);

                        sumAmountTotal += amountTotal;
                        sumSumTotal += summTotal;
                        sumPremiumTotal += premiunTotal;
                    }
                    lastRowSum.AMOUNT_CE = sumAmountCE.ToString("n0");
                    lastRowSum.SUMM_CE = sumSumCE.ToString("n0");
                    lastRowSum.PREMIUM_CE = sumPremiumCE.ToString("n0");

                    lastRowSum.AMOUNT_EA = sumAmountEA.ToString("n0");
                    lastRowSum.SUMM_EA = sumSumEA.ToString("n0");
                    lastRowSum.PREMIUM_EA = sumPremiumEA.ToString("n0");

                    lastRowSum.AMOUNT_NE = sumAmountNE.ToString("n0");
                    lastRowSum.SUMM_NE = sumSumNE.ToString("n0");
                    lastRowSum.PREMIUM_NE = sumPremiumNE.ToString("n0");

                    lastRowSum.AMOUNT_NO = sumAmountNO.ToString("n0");
                    lastRowSum.SUMM_NO = sumSumNO.ToString("n0");
                    lastRowSum.PREMIUM_NO = sumPremiumNO.ToString("n0");

                    lastRowSum.AmountTotal = sumAmountTotal.ToString("n0");
                    lastRowSum.SummTotal = sumSumTotal.ToString("n0");
                    lastRowSum.PremiumTotal = sumPremiumTotal.ToString("n0");
                    datas = datas.Concat(new PolicyOfficeModel[] { lastRowSum }).ToArray();

                }


            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }


        public PolicyOfficeOfMonthModel[] GetPolicyOfficeOfMonth(RequestPolicyOfMonthModel sPolicy) => GetPolicyOfficeOfMonth(sPolicy, (Repository)null);

        private PolicyOfficeOfMonthModel[] GetPolicyOfficeOfMonth(RequestPolicyOfMonthModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyOfficeOfMonthModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = string.Empty;
                var txtEndPolDT = string.Empty;
                if (sPolicy.Month.Equals("00"))
                {
                    txtStartPolDT = "01/01/" + sPolicy.Year;
                    txtEndPolDT = "31/12/" + sPolicy.Year;
                }
                else
                {
                    txtStartPolDT = "01/" + sPolicy.Month + "/" + sPolicy.Year;
                    txtEndPolDT = manage.GetLastDayOfMonth(int.Parse(sPolicy.Year), int.Parse(sPolicy.Month));
                }

                datas = repository.GetPolicyOfficeOfMonth(sPolicy.Section, txtStartPolDT, txtEndPolDT);
                PolicyOfficeOfMonthModel lstRowSum = new PolicyOfficeOfMonthModel();
                lstRowSum.SECTION_NAME = "รวม";
                lstRowSum.SECTION_CODE = "ALL";
                decimal sumAMOUNT_POLICY, sumSUMM_POLICY, sumPREMIUM_POLICY;
                sumAMOUNT_POLICY = sumSUMM_POLICY = sumPREMIUM_POLICY = 0;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lstRowSum.Id = (datas.Length + 1).ToString();
                        }

                        data.AMOUNT_POLICY = decimal.Parse(data.AMOUNT_POLICY).ToString("n0");
                        data.PREMIUM_POLICY = decimal.Parse(data.PREMIUM_POLICY).ToString("n0");
                        data.SUMM_POLICY = decimal.Parse(data.SUMM_POLICY).ToString("n0");

                        sumAMOUNT_POLICY += decimal.Parse(data.AMOUNT_POLICY);
                        sumPREMIUM_POLICY += decimal.Parse(data.PREMIUM_POLICY);
                        sumSUMM_POLICY += decimal.Parse(data.SUMM_POLICY);

                    }

                    lstRowSum.AMOUNT_POLICY = sumAMOUNT_POLICY.ToString("n0");
                    lstRowSum.PREMIUM_POLICY = sumPREMIUM_POLICY.ToString("n0");
                    lstRowSum.SUMM_POLICY = sumSUMM_POLICY.ToString("n0");

                    datas = datas.Concat(new PolicyOfficeOfMonthModel[] { lstRowSum }).ToArray();
                }

            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }


        public DropDownListModel[] GetRegion() => GetRegion((Repository)null);

        private DropDownListModel[] GetRegion(Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            List<DropDownListModel> datas = new List<DropDownListModel>();
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                datas.Add(new DropDownListModel
                {
                    id = "ALL",
                    text = "ส่วนทั้งหมด"
                });
                DataTable region = repository.GetRegion();
                foreach (DataRow row in region.Rows)
                {
                    datas.Add(new DropDownListModel
                    {
                        id = row.ItemArray[0].ToString(),
                        text = row.ItemArray[1].ToString()
                    });
                }

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas.ToArray();
        }

        public DropDownListModel[] GetRegionSub(string secCode) => GetRegionSub(secCode, (Repository)null);

        private DropDownListModel[] GetRegionSub(string secCode, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            List<DropDownListModel> datas = new List<DropDownListModel>();
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                datas.Add(new DropDownListModel
                {
                    id = "ALL",
                    text = "ภาคทั้งหมด"
                });
                DataTable subRegion = repository.GetRegionSub(secCode);
                foreach (DataRow row in subRegion.Rows)
                {
                    datas.Add(new DropDownListModel
                    {
                        id = row.ItemArray[0].ToString(),
                        text = row.ItemArray[1].ToString()
                    });
                }

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas.ToArray();
        }

        public DropDownListModel[] GetBkkRegionDdl(string Dep) => GetBkkRegionDdl(Dep, (Repository)null);

        private DropDownListModel[] GetBkkRegionDdl(string Dep, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            List<DropDownListModel> datas = new List<DropDownListModel>();
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                datas.Add(new DropDownListModel
                {
                    id = "ALL",
                    text = "สาขาทั้งหมด"
                });
                DataTable subRegion = repository.GetBkkRegionDdl(Dep, "");
                foreach (DataRow row in subRegion.Rows)
                {
                    datas.Add(new DropDownListModel
                    {
                        id = row.ItemArray[1].ToString(),
                        text = row.ItemArray[12].ToString()
                    });
                }

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas.ToArray();
        }

        public PolicyOfficeOfRegionModel[] GetPolicyOfficeOfRegion(RequestPolicyOfRegionModel sPolicy) => GetPolicyOfficeOfRegion(sPolicy, (Repository)null);

        private PolicyOfficeOfRegionModel[] GetPolicyOfficeOfRegion(RequestPolicyOfRegionModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyOfficeOfRegionModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = string.Empty;
                var txtEndPolDT = string.Empty;
                if (sPolicy.Month.Equals("00"))
                {
                    txtStartPolDT = "01/01/" + sPolicy.Year;
                    txtEndPolDT = "31/12/" + sPolicy.Year;
                }
                else
                {
                    txtStartPolDT = "01/" + sPolicy.Month + "/" + sPolicy.Year;
                    txtEndPolDT = manage.GetLastDayOfMonth(int.Parse(sPolicy.Year), int.Parse(sPolicy.Month));
                }

                datas = repository.GetPolicyOfficeOfRegion(sPolicy.Section, sPolicy.Region, txtStartPolDT, txtEndPolDT);
                PolicyOfficeOfRegionModel lstRowSum = new PolicyOfficeOfRegionModel();
                lstRowSum.REGION_NAME = "รวม";
                lstRowSum.REGION_CODE = "ALL";
                decimal sumAMOUNT_POLICY, sumSUMM_POLICY, sumPREMIUM_POLICY;
                sumAMOUNT_POLICY = sumSUMM_POLICY = sumPREMIUM_POLICY = 0;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lstRowSum.Id = (datas.Length + 1).ToString();
                        }

                        data.AMOUNT_POLICY = decimal.Parse(data.AMOUNT_POLICY).ToString("n0");
                        data.PREMIUM_POLICY = decimal.Parse(data.PREMIUM_POLICY).ToString("n0");
                        data.SUMM_POLICY = decimal.Parse(data.SUMM_POLICY).ToString("n0");

                        sumAMOUNT_POLICY += decimal.Parse(data.AMOUNT_POLICY);
                        sumPREMIUM_POLICY += decimal.Parse(data.PREMIUM_POLICY);
                        sumSUMM_POLICY += decimal.Parse(data.SUMM_POLICY);

                    }

                    lstRowSum.AMOUNT_POLICY = sumAMOUNT_POLICY.ToString("n0");
                    lstRowSum.PREMIUM_POLICY = sumPREMIUM_POLICY.ToString("n0");
                    lstRowSum.SUMM_POLICY = sumSUMM_POLICY.ToString("n0");

                    datas = datas.Concat(new PolicyOfficeOfRegionModel[] { lstRowSum }).ToArray();
                }

            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }

        public PolicyOfficeOfDepModel[] GetPolicyOfficeOfDep(RequestPolicyOfDepModel sPolicy) => GetPolicyOfficeOfDep(sPolicy, (Repository)null);

        private PolicyOfficeOfDepModel[] GetPolicyOfficeOfDep(RequestPolicyOfDepModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyOfficeOfDepModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = string.Empty;
                var txtEndPolDT = string.Empty;
                if (sPolicy.Month.Equals("00"))
                {
                    txtStartPolDT = "01/01/" + sPolicy.Year;
                    txtEndPolDT = "31/12/" + sPolicy.Year;
                }
                else
                {
                    txtStartPolDT = "01/" + sPolicy.Month + "/" + sPolicy.Year;
                    txtEndPolDT = manage.GetLastDayOfMonth(int.Parse(sPolicy.Year), int.Parse(sPolicy.Month));
                }

                datas = repository.GetPolicyOfficeOfDep(sPolicy.Section, sPolicy.Region, sPolicy.Office, txtStartPolDT, txtEndPolDT);
                PolicyOfficeOfDepModel lstRowSum = new PolicyOfficeOfDepModel();
                lstRowSum.DESCRIPTION = "รวม";
                lstRowSum.OFFICE = "ALL";
                decimal sumAMOUNT_POLICY, sumSUMM_POLICY, sumPREMIUM_POLICY, sumAmountDayClean, sumAmountClean, sumCleanPercent, sumAmountDayUnClean, sumAmountUnClean, sumUnCleanPercent, sumTotalAmountDay, sumTotalAmount ,sumSUMAmountDayUnClean, sumSUMAmountDayClean;
                sumAMOUNT_POLICY = sumSUMM_POLICY = sumPREMIUM_POLICY = sumAmountDayClean = sumAmountClean = sumCleanPercent = sumAmountDayUnClean = sumAmountUnClean = sumUnCleanPercent = sumTotalAmountDay = sumTotalAmount = sumSUMAmountDayUnClean = sumSUMAmountDayClean = 0;
                int i = 1;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        data.Id = i.ToString();
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lstRowSum.Id = (datas.Length + 1).ToString();
                        }

                        data.AMOUNT_POLICY = decimal.Parse(data.AMOUNT_POLICY).ToString("n0");
                        data.PREMIUM_POLICY = decimal.Parse(data.PREMIUM_POLICY).ToString("n0");
                        data.SUMM_POLICY = decimal.Parse(data.SUMM_POLICY).ToString("n0");
                        data.CLEAN_PERCENT = Convert.ToInt32(decimal.Parse(data.CLEAN_PERCENT)).ToString();
                        data.UNCLEAN_PERCENT = Convert.ToInt32(decimal.Parse(data.UNCLEAN_PERCENT)).ToString();

                        sumAMOUNT_POLICY += decimal.Parse(data.AMOUNT_POLICY);
                        sumPREMIUM_POLICY += decimal.Parse(data.PREMIUM_POLICY);
                        sumSUMM_POLICY += decimal.Parse(data.SUMM_POLICY);
                        sumAmountDayClean += decimal.Parse(data.AMOUNT_DAY_CLEAN);
                        sumAmountClean += decimal.Parse(data.AMOUNT_CLEAN);
                        sumCleanPercent += decimal.Parse(data.CLEAN_PERCENT);
                        sumAmountDayUnClean += decimal.Parse(data.AMOUNT_DAY_UNCLEAN);
                        sumAmountUnClean += decimal.Parse(data.AMOUNT_UNCLEAN);
                        sumUnCleanPercent += decimal.Parse(data.UNCLEAN_PERCENT);
                        sumTotalAmountDay += decimal.Parse(data.TOTAL_AMOUNT_DAY);
                        sumTotalAmount += decimal.Parse(data.TOTAL_AMOUNT);
                        sumSUMAmountDayUnClean += decimal.Parse(data.SUM_AMOUNT_DAY_UNCLEAN);
                        sumSUMAmountDayClean += decimal.Parse(data.SUM_AMOUNT_DAY_CLEAN);
                        i++;
                    }

                    lstRowSum.AMOUNT_POLICY = sumAMOUNT_POLICY.ToString("n0");
                    lstRowSum.PREMIUM_POLICY = sumPREMIUM_POLICY.ToString("n0");
                    lstRowSum.SUMM_POLICY = sumSUMM_POLICY.ToString("n0");

                    lstRowSum.AMOUNT_DAY_CLEAN = sumAmountClean == 0 ? sumAmountClean.ToString() : Math.Round(sumSUMAmountDayClean / sumAmountClean).ToString();
                    lstRowSum.AMOUNT_CLEAN = sumAmountClean.ToString();
                    lstRowSum.CLEAN_PERCENT = sumTotalAmount == 0 ? sumTotalAmount.ToString() : Convert.ToInt32((sumAmountClean / sumTotalAmount *100)).ToString() ;
                    lstRowSum.AMOUNT_DAY_UNCLEAN = sumAmountUnClean == 0 ? sumAmountUnClean.ToString() : Math.Round(sumSUMAmountDayUnClean / sumAmountUnClean).ToString();
                    lstRowSum.AMOUNT_UNCLEAN = sumAmountUnClean.ToString();
                    lstRowSum.UNCLEAN_PERCENT = sumTotalAmount == 0 ? sumTotalAmount.ToString() : Convert.ToInt32((sumAmountUnClean / sumTotalAmount * 100)).ToString();
                    lstRowSum.TOTAL_AMOUNT_DAY = sumAmountClean == 0 && sumAmountUnClean == 0 ? (sumAmountClean + sumAmountUnClean).ToString() : Math.Round((sumSUMAmountDayClean + sumSUMAmountDayUnClean) / (sumAmountClean + sumAmountUnClean)).ToString();
                    lstRowSum.TOTAL_AMOUNT = sumTotalAmount.ToString();
                    lstRowSum.SUM_AMOUNT_DAY_CLEAN = sumSUMAmountDayClean.ToString();
                    lstRowSum.SUM_AMOUNT_DAY_UNCLEAN = sumSUMAmountDayUnClean.ToString();

                    datas = datas.Concat(new PolicyOfficeOfDepModel[] { lstRowSum }).ToArray();
                }

            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }

        public PolicyOfficeOfDateModel[] GetPolicyOfficeOfDate(RequestPolicyOfDepModel sPolicy) => GetPolicyOfficeOfDate(sPolicy, (Repository)null);

        private PolicyOfficeOfDateModel[] GetPolicyOfficeOfDate(RequestPolicyOfDepModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyOfficeOfDateModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = string.Empty;
                var txtEndPolDT = string.Empty;
                if (sPolicy.Month.Equals("00"))
                {
                    txtStartPolDT = "01/01/" + sPolicy.Year;
                    txtEndPolDT = "31/12/" + sPolicy.Year;
                }
                else
                {
                    txtStartPolDT = "01/" + sPolicy.Month + "/" + sPolicy.Year;
                    txtEndPolDT = manage.GetLastDayOfMonth(int.Parse(sPolicy.Year), int.Parse(sPolicy.Month));
                }

                datas = repository.GetPolicyOfficeOfDate(sPolicy.Region, sPolicy.Office, txtStartPolDT, txtEndPolDT);
                PolicyOfficeOfDateModel lstRowSum = new PolicyOfficeOfDateModel();
                lstRowSum.INSTALL_DT = "รวม";
                decimal sumAMOUNT_POLICY, sumSUMM_POLICY, sumPREMIUM_POLICY, sumAmountDayClean, sumSUMAmountDayClean, sumSUMAmountDayUnClean, sumAmountClean, sumCleanPercent, sumAmountDayUnClean, sumAmountUnClean, sumUnCleanPercent, sumTotalAmountDay, sumTotalAmount, sumAmountSendPolicy, sumAmountNotSendPolicy;
                sumAMOUNT_POLICY = sumSUMM_POLICY = sumPREMIUM_POLICY = sumAmountDayClean = sumSUMAmountDayClean = sumSUMAmountDayUnClean = sumAmountClean = sumCleanPercent = sumAmountDayUnClean = sumAmountUnClean = sumUnCleanPercent = sumTotalAmountDay = sumTotalAmount = sumAmountSendPolicy = sumAmountNotSendPolicy = 0;
                int i = 1;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        data.Id = i.ToString();
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lstRowSum.Id = (datas.Length + 1).ToString();
                        }
                        data.INSTALL_DT = manage.GetDateFomatTHN(data.INSTALL_DT);
                        data.AMOUNT_POLICY = decimal.Parse(data.AMOUNT_POLICY).ToString("n0");
                        data.PREMIUM_POLICY = decimal.Parse(data.PREMIUM_POLICY).ToString("n0");
                        data.SUMM_POLICY = decimal.Parse(data.SUMM_POLICY).ToString("n0");
                        data.CLEAN_PERCENT = Math.Round(decimal.Parse(data.CLEAN_PERCENT)).ToString(); 
                        data.UNCLEAN_PERCENT = Math.Round(decimal.Parse(data.UNCLEAN_PERCENT)).ToString();

                        sumAMOUNT_POLICY += decimal.Parse(data.AMOUNT_POLICY);
                        sumPREMIUM_POLICY += decimal.Parse(data.PREMIUM_POLICY);
                        sumSUMM_POLICY += decimal.Parse(data.SUMM_POLICY);
                        sumAmountSendPolicy += decimal.Parse(data.AMOUNT_SEND_POLICY);
                        sumAmountNotSendPolicy += decimal.Parse(data.AMOUNT_NOTSEND_POLICY);
                        sumAmountDayClean += decimal.Parse(data.AMOUNT_DAY_CLEAN);
                        sumAmountClean += decimal.Parse(data.AMOUNT_CLEAN);
                        sumCleanPercent += decimal.Parse(data.CLEAN_PERCENT);
                        sumAmountDayUnClean += decimal.Parse(data.AMOUNT_DAY_UNCLEAN);
                        sumAmountUnClean += decimal.Parse(data.AMOUNT_UNCLEAN);
                        sumUnCleanPercent += decimal.Parse(data.UNCLEAN_PERCENT);
                        sumTotalAmountDay += decimal.Parse(data.TOTAL_AMOUNT_DAY);
                        sumTotalAmount += decimal.Parse(data.TOTAL_AMOUNT);
                        sumSUMAmountDayClean += decimal.Parse(data.SUM_AMOUNT_DAY_CLEAN);
                        sumSUMAmountDayUnClean += decimal.Parse(data.SUM_AMOUNT_DAY_UNCLEAN);
                        i++;
                    }

                    lstRowSum.AMOUNT_POLICY = sumAMOUNT_POLICY.ToString("n0");
                    lstRowSum.PREMIUM_POLICY = sumPREMIUM_POLICY.ToString("n0");
                    lstRowSum.SUMM_POLICY = sumSUMM_POLICY.ToString("n0");
                    lstRowSum.AMOUNT_SEND_POLICY = sumAmountSendPolicy.ToString("n0");
                    lstRowSum.AMOUNT_NOTSEND_POLICY = sumAmountNotSendPolicy.ToString("n0");
                    lstRowSum.AMOUNT_DAY_CLEAN = sumAmountClean == 0 ? sumAmountClean.ToString() : Math.Round((sumSUMAmountDayClean / sumAmountClean)).ToString();
                    lstRowSum.AMOUNT_CLEAN = sumAmountClean.ToString();
                    lstRowSum.CLEAN_PERCENT = sumAmountSendPolicy == 0 ? sumAmountSendPolicy.ToString() : Math.Round(sumAmountClean / sumAmountSendPolicy *100 ).ToString();
                    lstRowSum.AMOUNT_DAY_UNCLEAN = sumAmountUnClean == 0 ? sumAmountUnClean.ToString() : Math.Round(sumSUMAmountDayUnClean / sumAmountUnClean).ToString();
                    lstRowSum.AMOUNT_UNCLEAN = sumAmountUnClean.ToString();
                    lstRowSum.UNCLEAN_PERCENT = sumAmountSendPolicy == 0 ? sumAmountSendPolicy.ToString() : Math.Round(sumAmountUnClean / sumAmountSendPolicy * 100).ToString();
                    lstRowSum.TOTAL_AMOUNT_DAY = sumAmountClean == 0 && sumAmountUnClean == 0 ? (sumAmountClean + sumAmountUnClean).ToString() : Math.Round((sumSUMAmountDayClean + sumSUMAmountDayUnClean) / (sumAmountClean + sumAmountUnClean)).ToString("n0");
                    lstRowSum.TOTAL_AMOUNT = sumTotalAmount.ToString();

                    datas = datas.Concat(new PolicyOfficeOfDateModel[] { lstRowSum }).ToArray();
                }

            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }

        public PolicyListModel[] GetPolicyList(RequestPolicyListModel sPolicy) => GetPolicyList(sPolicy, (Repository)null);

        private PolicyListModel[] GetPolicyList(RequestPolicyListModel sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyListModel[] datas = null;
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                var txtStartPolDT = string.Empty;
                var txtEndPolDT = string.Empty;
                if (sPolicy.Month.Equals("00"))
                {
                    txtStartPolDT = "01/01/" + sPolicy.Year;
                    txtEndPolDT = "31/12/" + sPolicy.Year;
                }
                else
                {
                    txtStartPolDT = "01/" + sPolicy.Month + "/" + sPolicy.Year;
                    txtEndPolDT = manage.GetLastDayOfMonth(int.Parse(sPolicy.Year), int.Parse(sPolicy.Month));
                }
                sPolicy.StartInstallDT = manage.ConvertStrDateToTH(sPolicy.StartInstallDT);
                sPolicy.EndInstallDT = manage.ConvertStrDateToTH(sPolicy.EndInstallDT);
                sPolicy.BancType = "";
                sPolicy.PlanType = "";
                sPolicy.Send = "";
                sPolicy.Case = "";
                datas = repository.GetPolicyList(txtStartPolDT, txtEndPolDT, sPolicy.StartInstallDT, sPolicy.EndInstallDT, sPolicy.Type, sPolicy.Branch, sPolicy.BancType, sPolicy.PlanType, sPolicy.Send, sPolicy.Case);
                PolicyListModel lstRowSum = new PolicyListModel();
                lstRowSum.POLNO = "จำนวนข้อมูลทั้งหมด" + datas.Count().ToString();
                lstRowSum.IBBL_TRN_DATE = "รวม (1)-(2)";
                lstRowSum.CSP_UPD_DATE_TIME = "รวม (3)-(4)";
                lstRowSum.APPSYS_DATE = "รวม (2)-(5)";
                lstRowSum.INSTALL_DT = "รวม (10)-(11)";
                lstRowSum.SEND_DATE = "รวม(11)-(12)";
                decimal sumKpiBank, sumKPI_CSP, sumKPI_APPSYS, sumKPI_INSTALL, sumKPI_SEND, sumKPI_TOTAL, sumAMOUNT_CLEAN, sumAMOUNT_TOTAL, SumAmountDayClean, sumAMOUNT_UNCLEAN, SumAmountDayUnClean, sumAmountDayTotal;
                sumKpiBank = sumKPI_CSP = sumKPI_APPSYS = sumKPI_INSTALL = sumKPI_SEND = sumKPI_TOTAL = sumAMOUNT_CLEAN = sumAMOUNT_TOTAL = SumAmountDayClean = sumAMOUNT_UNCLEAN = SumAmountDayUnClean = sumAmountDayTotal = 0;
                int i = 1;
                if (datas != null)
                {
                    foreach (var data in datas)
                    {
                        data.Id = i.ToString();
                        if (int.Parse(data.Id) == datas.Length)
                        {
                            lstRowSum.Id = (datas.Length + 1).ToString();
                        }
                        sumKpiBank += decimal.Parse(data.KPI_BANK);
                        sumKPI_CSP += decimal.Parse(data.KPI_CSP);
                        sumKPI_APPSYS += decimal.Parse(data.KPI_APPSYS);
                        sumKPI_INSTALL += decimal.Parse(data.KPI_INSTALL);
                        sumKPI_SEND += decimal.Parse(data.KPI_SEND);
                        sumKPI_TOTAL += decimal.Parse(data.KPI_TOTAL);
                        sumAMOUNT_CLEAN += decimal.Parse(data.AMOUNT_CLEAN);
                        sumAMOUNT_UNCLEAN += decimal.Parse(data.AMOUNT_UNCLEAN);
                        int AmountTotal = 0;
                        if (data.SEND_DATE != "")
                        {
                            AmountTotal = 1;
                        }
                        sumAMOUNT_TOTAL += AmountTotal;
                        SumAmountDayClean += decimal.Parse(data.AMOUNT_DAY_CLEAN);
                        SumAmountDayUnClean += decimal.Parse(data.AMOUNT_DAY_UNCLEAN);
                        sumAmountDayTotal += decimal.Parse(data.KPI_TOTAL);
                        i++;
                    }
                    lstRowSum.KPI_BANK = sumKpiBank.ToString("n0");
                    lstRowSum.KPI_CSP = sumKPI_CSP.ToString("n0");
                    lstRowSum.KPI_APPSYS = sumKPI_APPSYS.ToString("n0");
                    lstRowSum.KPI_INSTALL = sumKPI_INSTALL.ToString("n0");
                    lstRowSum.KPI_SEND = sumKPI_SEND.ToString("n0");
                    lstRowSum.KPI_TOTAL = sumKPI_TOTAL.ToString("n0");
                    lstRowSum.AMOUNT_CLEAN = sumAMOUNT_CLEAN.ToString("n0");
                    lstRowSum.AMOUNT_UNCLEAN = sumAMOUNT_UNCLEAN.ToString("n0");
                    lstRowSum.AMOUNT_CLEAN_PERCENT = (((double)sumAMOUNT_CLEAN / (double)sumAMOUNT_TOTAL) * 100).ToString("n0") + "%";
                    lstRowSum.AMOUNT_UNCLEAN_PERCENT = (((double)sumAMOUNT_UNCLEAN / (double)sumAMOUNT_TOTAL) * 100).ToString("n0") + "%";
                    lstRowSum.CLEAN_CASE_AVG = (((double)SumAmountDayClean / (double)sumAMOUNT_CLEAN)).ToString("n0");
                    lstRowSum.UNCLEAN_CASE_AVG = (((double)SumAmountDayUnClean / (double)sumAMOUNT_UNCLEAN)).ToString("n0");
                    lstRowSum.SUM_AMOUNT_TOTAL = sumAMOUNT_TOTAL.ToString("n0");
                    if (sumAMOUNT_TOTAL > 0)
                    {
                        lstRowSum.SUM_PERCENT_TOTAL = "100%";
                        lstRowSum.SUM_TOTAL_AVG = ((decimal)sumAmountDayTotal / (decimal)sumAMOUNT_TOTAL).ToString("n0");
                    }
                    datas = datas.Concat(new PolicyListModel[] { lstRowSum }).ToArray();
                }

            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas;
        }

        public DropDownListModel[] GetBkkRegion(string Dep) => GetBkkRegion(Dep, (Repository)null);

        private DropDownListModel[] GetBkkRegion(string Dep, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            List<DropDownListModel> datas = new List<DropDownListModel>();
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here
                //datas.Add(new DropDownListModel
                //{
                //    id = "00",
                //    text = "แสดงข้อมูลทุกสาขา"
                //});
                DataTable subRegion = repository.GetBkkRegion(Dep);
                foreach (DataRow row in subRegion.Rows)
                {
                    datas.Add(new DropDownListModel
                    {
                        id = row.ItemArray[0].ToString(),
                        text = row.ItemArray[1].ToString()
                    });
                }

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return datas.ToArray();
        }

        public PolicyDetailModel GetPolicyDetail(RequestPolicyDetail sPolicy) => GetPolicyDetail(sPolicy, (Repository)null);

        private PolicyDetailModel GetPolicyDetail(RequestPolicyDetail sPolicy, Repository repository)
        {
            bool internalConnection = false;
            //create return type model
            PolicyDetailModel data = new PolicyDetailModel();
            if (repository is null)
            {
                repository = new Repository(connectionString);
                repository.OpenConnection();
                internalConnection = true;
            }

            try
            {
                //Business logic here

                data.CR_POLICY_REC = repository.GetPolicyDetail(sPolicy.Policy);
                data.CR_POLICY_REC.BIRTH_DT = data.CR_POLICY_REC.BIRTH_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.ISU_DT = data.CR_POLICY_REC.ISU_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.MAT_DT = data.CR_POLICY_REC.MAT_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.NXTDUE_DT = data.CR_POLICY_REC.NXTDUE_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.STR_DT = data.CR_POLICY_REC.STR_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.INSTALL_DT = data.CR_POLICY_REC.INSTALL_DT.Split(' ')[0].ToString();
                data.CR_POLICY_REC.INFSUMM = decimal.Parse(data.CR_POLICY_REC.INFSUMM).ToString("n0");
                data.CR_POLICY_REC.SUMM = decimal.Parse(data.CR_POLICY_REC.SUMM).ToString("n0");

                data.CR_PREMIUM_REC = repository.GetPolByPremium(sPolicy.Policy);
                decimal sumSUMM, sumPREMIUM, sumXTR_PREMIUM, sumSumPremiumPlusXtr;
                sumSUMM = sumPREMIUM = sumXTR_PREMIUM = sumSumPremiumPlusXtr = 0;
                int i = 1;
                CR_PREMIUM_REC lstPremiumSum = new CR_PREMIUM_REC();
                if (data.CR_PREMIUM_REC.Count() > 0)
                {
                    foreach (var dp in data.CR_PREMIUM_REC)
                    {
                        dp.Id = i.ToString();
                        if (int.Parse(dp.Id) == data.CR_PREMIUM_REC.Length)
                        {
                            lstPremiumSum.PLAN_NAME = "รวม";
                        }
                        dp.SUMM = decimal.Parse(dp.SUMM).ToString("n0");
                        dp.PREMIUM = decimal.Parse(dp.PREMIUM).ToString("n0");
                        dp.XTR_PREMIUM = decimal.Parse(dp.XTR_PREMIUM).ToString("n0");
                        dp.SumPremiumPlusXtr = (decimal.Parse(dp.PREMIUM) + decimal.Parse(dp.XTR_PREMIUM)).ToString("n0");
                        dp.MAT_DT = dp.MAT_DT.Split(' ')[0].ToString();
                        dp.LASTPAY_DT = dp.LASTPAY_DT.Split(' ')[0].ToString();

                        sumSUMM += decimal.Parse(dp.SUMM);
                        sumPREMIUM += decimal.Parse(dp.PREMIUM);
                        sumXTR_PREMIUM += decimal.Parse(dp.XTR_PREMIUM);
                        sumSumPremiumPlusXtr += decimal.Parse(dp.SumPremiumPlusXtr);
                        i++;
                    }
                    lstPremiumSum.Id = i.ToString();
                    lstPremiumSum.SUMM = sumSUMM.ToString("n0");
                    lstPremiumSum.PREMIUM = sumPREMIUM.ToString("n0");
                    lstPremiumSum.XTR_PREMIUM = sumXTR_PREMIUM.ToString("n0");
                    lstPremiumSum.SumPremiumPlusXtr = sumSumPremiumPlusXtr.ToString("n0");
                }
                data.CR_PREMIUM_REC = data.CR_PREMIUM_REC.Concat(new CR_PREMIUM_REC[] { lstPremiumSum }).ToArray();

                if (sPolicy.Office.Equals("D"))
                {
                    data.Saleco = repository.GetSalecoByPolCaseD(sPolicy.Policy);
                }
                else
                {
                    data.Agent = repository.GetSalecoByPol(sPolicy.Policy);
                }
            }
            catch (Exception ex)
            {
                throw new NotImplementedException();

            }
            finally
            {
                if (internalConnection)
                {
                    repository.CloseConnection();
                }
            }
            return data;
        }
    }
}
