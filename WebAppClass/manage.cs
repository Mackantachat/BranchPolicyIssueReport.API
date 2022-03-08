
using System;
using System.Data;
using System.Configuration;
//using System.Linq;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
//using System.Xml.Linq;

/// <summary>
/// Summary description for manage
/// </summary>
/// 
namespace WebApp2 {

    public static class manage {
        public static string GetFirstDate() {
            string stYear = "";
            stYear = ((DateTime.Now.Year) + 543).ToString();
            return "01/01/" + stYear;
        }
        public static string GetDateNow() {
            string stDay, stMonth, stYear;
            stDay = "";
            stMonth = "";
            stYear = "";
            stDay = DateTime.Now.Day.ToString();
            stMonth = DateTime.Now.Month.ToString();
            stYear = ((DateTime.Now.Year) + 543).ToString();
            return stDay + "/" + stMonth + "/" + stYear;
        }
        public static string GetDateFomatTH(this DateTime myDate) {
            if ((myDate == null) || (myDate.ToShortDateString() == "")) {
                return "";
            }
            else {
                string stDay, stMonth, stYear;
                stDay = "";
                stMonth = "";
                stYear = "";
                stDay = myDate.Day.ToString();
                stMonth = myDate.Month.ToString();
                stYear = ((myDate.Year) + 543).ToString();
                return stDay + "/" + stMonth + "/" + stYear;
            }

        }
        public static string GetDateFomatTHN(this string myDate) {
            if ((myDate == null) || (myDate == "")) {
                return "";
            }
            else {
                DateTime stDate;
                string stDay, stMonth, stYear;
                stDate = DateTime.Parse(myDate);
                stDay = "";
                stMonth = "";
                stYear = "";
                stDay = stDate.Day.ToString();
                stMonth = stDate.Month.ToString();
                stYear = ((stDate.Year) + 543).ToString();
                return stDay + "/" + stMonth + "/" + stYear;
            }

        }
        public static string GetDateTimeIsFomatDate(this string myDate)
        {
            if ((myDate == null) || (myDate == ""))
            {
                return "";
            }
            else
            {
                DateTime stDate;
                string stDay, stMonth, stYear;
                stDate = DateTime.Parse(myDate);
                stDay = "";
                stMonth = "";
                stYear = "";
                stDay = stDate.Day.ToString();
                stMonth = stDate.Month.ToString();
                stYear = stDate.Year.ToString();
                return stDay + "/" + stMonth + "/" + stYear;
            }

        }
        public static string GetDateFomatEN(this String myDate) {

            if ((myDate == null) || (myDate == "")) {
                return "";
            }
            else {
                //DateTime stDate;
                //string stDay, stMonth, stYear;
                //stDate = DateTime.Parse(myDate);

                //stDay = "";
                //stMonth = "";
                //stYear = "";
                //stDay = stDate.Day.ToString();
                //stMonth = stDate.Month.ToString();
                //stYear = stDate.Year.ToString();
                ////stYear = stDate.Year.ToString("0000", System.Globalization.DateTimeFormatInfo.InvariantInfo);

                //return stDay + "/" + stMonth + "/" + (int.Parse(stYear) - 543).ToString();
                ////return stDate.ToString("dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                ////return stDay + "/" + stMonth + "/" + stYear;
                ////return stDate.ToString("dd/MM/yyyy", System.Globalization.DateTimeFormatInfo.InvariantInfo);
                string stDay, stMonth, stYear;
                int istDay, istMonth, istYear;
                stDay = myDate.Split("/".ToCharArray())[0];
                stMonth = myDate.Split("/".ToCharArray())[1];
                stYear = myDate.Split("/".ToCharArray())[2];
                istDay = int.Parse(stDay);
                istMonth = int.Parse(stMonth);
                istYear = int.Parse(stYear) - 543;
                return stDay + "/" + stMonth + "/" + istYear.ToString();
            }
        }
        public static string GetCovertStrDate(this string myDate) {
            string stDay, stMonth, stYear, stDate;
            stDate = "";
            stDay = "";
            stMonth = "";
            stYear = "";
            if (myDate.Length == 6) {
                stYear = myDate.Substring(0, 2);
                stMonth = myDate.Substring(2, 2);
                stDay = myDate.Substring(4, 2);
                stDate = stDay + "/" + stMonth + "/25" + stYear;
            }
            else if (myDate.Length == 8) {
                stYear = myDate.Substring(0, 4);
                stMonth = myDate.Substring(4, 2);
                stDay = myDate.Substring(6, 2);
                stDate = stDay + "/" + stMonth + "/" + stYear;
            }
            else {
                stDate = "";
            }
            return stDate;
        }
        public static string GetStrDate(this string myDate) {
            string[] buf = System.Text.RegularExpressions.Regex.Split(myDate, "/");
            string stDay, stMonth, stYear;
            stDay = "";
            stMonth = "";
            stYear = "";
            stDay = buf[0];
            stMonth = buf[1];
            stYear = buf[2];
            stYear = stYear.Substring(2, 2);
            if (stDay.Length == 1) {
                stDay = "0" + stDay;
            }
            else {
                stDay = stDay;
            }
            if (stMonth.Length == 1) {
                stMonth = "0" + stMonth;
            }
            else {
                stMonth = stMonth;
            }
            return stYear + stMonth + stDay;
        }
        public static string GetDateEN(this string myDate) {
            string[] buf = System.Text.RegularExpressions.Regex.Split(myDate, "/");
            int yearEN = int.Parse(buf[2]) - 543;
            return buf[0] + "/" + buf[1] + "/" + yearEN.ToString();
        }
        public static string GetThaiMonth(this string myMonth) {
            string NameMonth;
            switch (myMonth) {
                case "00":
                    NameMonth = "ทุกเดือน";
                    break;
                case "01":
                    NameMonth = "มกราคม";
                    break;
                case "02":
                    NameMonth = "กุมภาพันธ์";
                    break;
                case "03":
                    NameMonth = "มีนาคม";
                    break;
                case "04":
                    NameMonth = "เมษายน";
                    break;
                case "05":
                    NameMonth = "พฤษภาคม";
                    break;
                case "06":
                    NameMonth = "มิถุนายน";
                    break;
                case "07":
                    NameMonth = "กรกฎาคม";
                    break;
                case "08":
                    NameMonth = "สิงหาคม";
                    break;
                case "09":
                    NameMonth = "กันยายน";
                    break;
                case "10":
                    NameMonth = "ตุลาคม";
                    break;
                case "11":
                    NameMonth = "พฤศจิกายน";
                    break;
                case "12":
                    NameMonth = "ธันวาคม";
                    break;
                default:
                    NameMonth = "Error";
                    break;
            }
            return NameMonth;
        }

        public static DataTable TableOfMonthShow() {
            DataTable myDataTable = new DataTable();
            DataColumn myDataColumn;

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "MONTH";
            myDataTable.Columns.Add(myDataColumn);

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "MONTH_NAME";
            myDataTable.Columns.Add(myDataColumn);

            for (int i = 0; i <= 12; i++) {
                DataRow dr;
                dr = myDataTable.NewRow();
                string month;
                if (i.ToString().Length == 1) {
                    month = "0" + i.ToString();
                }
                else {
                    month = i.ToString();
                }
                dr["MONTH"] = month;
                dr["MONTH_NAME"] = GetThaiMonth(month);
                myDataTable.Rows.Add(dr);
            }
            return myDataTable;
        }
        public static DataTable TableOfMonthShow1() {
            DataTable myDataTable = new DataTable();
            DataColumn myDataColumn;

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "MONTH";
            myDataTable.Columns.Add(myDataColumn);

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "MONTH_NAME";
            myDataTable.Columns.Add(myDataColumn);

            for (int i = 1; i <= 12; i++) {
                DataRow dr;
                dr = myDataTable.NewRow();
                string month;
                if (i.ToString().Length == 1) {
                    month = "0" + i.ToString();
                }
                else {
                    month = i.ToString();
                }
                dr["MONTH"] = month;
                dr["MONTH_NAME"] = GetThaiMonth(month);
                myDataTable.Rows.Add(dr);
            }
            return myDataTable;
        }
        public static DataTable TableOfYearShow() {
            DataTable myDataTable = new DataTable();
            DataColumn myDataColumn;

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "YEAR";
            myDataTable.Columns.Add(myDataColumn);
            int stYear = (DateTime.Now.Year) + 543;
            for (int i = (stYear - 10); i <= stYear; i++) {
                DataRow dr;
                dr = myDataTable.NewRow();
                dr["YEAR"] = i.ToString();
                myDataTable.Rows.Add(dr);
            }
            return myDataTable;

        }
        public static string GetDepartment(this string myType) {
            string DepName;
            switch (myType) {
                case "A":
                    DepName = "ส่วนกรมธรรม์สามัญ 1";
                    break;
                case "B":
                    DepName = "ส่วนกรมธรรม์สามัญ 2";
                    break;
                case "C":
                    DepName = "สะสมเงินเดือน";
                    break;
                case "D":
                    DepName = "Bancassurance";
                    break;
                case "E":
                    DepName = "Tele Market";
                    break;
                case "SO":
                    DepName = "ภาคใต้";
                    break;
                case "NE":
                    DepName = "ภาคตะวันออกเฉียงเหนือ";
                    break;
                case "CE":
                    DepName = "ภาคนครหลวง และภาคกลาง";
                    break;
                case "EA":
                    DepName = "ภาคตะวันออก และภาคกลาง";
                    break;
                case "NO":
                    DepName = "ภาคเหนือ";
                    break;

                default:
                    DepName = "Error";
                    break;
            }
            return DepName;

        }
        public static string GetLastDayOfMonth(this int myYear, int myMonth) {
            string LDay;
            string NDay;
            myYear = (myYear - 543);
            LDay = DateTime.DaysInMonth(myYear, myMonth).ToString();
            NDay = LDay + "/" + myMonth.ToString() + "/" + (myYear + 543).ToString();
            return NDay;
        }
        public static DataTable TableOfMonth(this int myYear, int myMonth) {
            myYear = myYear - 543;
            int totalDay = DateTime.DaysInMonth(myYear, myMonth);

            DataTable myDataTable = new DataTable();
            DataColumn myDataColumn;

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "APP_DATE";
            myDataTable.Columns.Add(myDataColumn);

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "APP_DAY";
            myDataTable.Columns.Add(myDataColumn);
            int j;
            for (int i = 0; i < totalDay; i++) {
                DataRow dr;
                dr = myDataTable.NewRow();
                j = i + 1;
                dr["APP_DATE"] = j.ToString() + "/" + myMonth.ToString() + "/" + (myYear + 543).ToString();
                dr["APP_DAY"] = j.ToString();
                myDataTable.Rows.Add(dr);
            }
            return myDataTable;
        }
        public static DataTable TableType(this string myDate, string myDay) {
            DataTable myDataTable = new DataTable();
            DataColumn myDataColumn;
            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "KEY_IN";
            myDataTable.Columns.Add(myDataColumn);

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "APP_DATE";
            myDataTable.Columns.Add(myDataColumn);

            myDataColumn = new DataColumn();
            myDataColumn.DataType = Type.GetType("System.String");
            myDataColumn.ColumnName = "APP_DAY";
            myDataTable.Columns.Add(myDataColumn);

            for (int i = 1; i <= 2; i++) {
                DataRow dr;
                dr = myDataTable.NewRow();
                if (i == 1) {
                    dr["KEY_IN"] = "A";
                }
                else {
                    dr["KEY_IN"] = "B";
                }
                dr["APP_DATE"] = myDate;
                dr["APP_DAY"] = myDay;
                myDataTable.Rows.Add(dr);
            }
            return myDataTable;
        }
        public static string GetMonth2Digit(this string myMonth) {
            string StrMonth = "";
            int LenMonth = 0;
            LenMonth = myMonth.Length;
            if (LenMonth > 1) {
                StrMonth = myMonth;
            }
            else {
                StrMonth = "0" + myMonth;
            }
            return StrMonth;
        }

        public static string ConvertStrYearTHToEN(this string year)
        {
            if (string.IsNullOrEmpty(year))
            {
                return "";
            }
            int yearInt = Convert.ToInt32(year) - 543;
            return yearInt.ToString();
        }

        public static string ConvertStrDateToTH(this string date)
        {
            if (string.IsNullOrEmpty(date))
            {
                return "";
            }
            var sDate = date.Split('/');
            int yearInt = Convert.ToInt32(sDate[2]) + 543;
            return sDate[0] + "/" + sDate[1] + "/" + yearInt.ToString();
        }
    }
}