using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BudgetActualReport.Models
{
    public class AgeingAnalysis
    {
        public int Cid { get; set; }
        public string sessionId { get; set; }
        public int Userid { get; set; }
        public List<Ageing> Ageing { get; set; }

        public string ReportDate { get; set; }
        public string Selection { get; set; }
        public string Accounts { get; set; }
        public string Currency { get; set; }
        public string SalesMan { get; set; }
        public string Months { get; set; }
    }

    public class Months
    {
        public string Month { get; set; }
        public string MonthYear { get; set; }
    }
    public class AccountMaster
    {
        public int iMasterId { get; set; }
    }
    public class Ageing
    {
        public string AccountName { get; set; }
        public string VoucherName { get; set; }
        public string SalesMan { get; set; }
        public string LPONo { get; set; }
        public decimal InvoiceAmt { get; set; }
        public decimal BalanceAmt { get; set; }
        public string Date { get; set; }
        public int DelayDays { get; set; }
        public string Month { get; set; }
    }
}