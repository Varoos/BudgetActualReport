using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace BudgetActualReport.Models
{
    public class BudgetVsActual
    {
        public string Category { get; set; }
        public decimal Budget { get; set; }
        public decimal NonPo { get; set; }
        public decimal PO { get; set; }
        public decimal Forcast { get; set; }
        public decimal Save { get; set; }
    }
    public class Analysis
    {
        public decimal Initial_Order_Value { get; set; }
        public decimal Variation { get; set; }
        public decimal Total_Sales_Value { get; set; }
        public decimal ActualCost { get; set; }
        public decimal Forcasted { get; set; }
        public decimal TotalCost { get; set; }
        public decimal InvoicedTillDate { get; set; }
        public decimal Pending { get; set; }
        public decimal  Total { get; set; }
        public decimal Received { get; set; }
        public decimal Retension { get; set; }
        public decimal Outstanding { get; set; }
    }
    public class Budget_Actual_Analysis
    {
        public Analysis analysis { get; set; }
        public List<BudgetVsActual> budgetvsactuallist { get; set; }
        public string Projects { get; set; }
        public string ReportDate { get; set; }
        public int CompanyId { get; set; }
        public string SessionId { get; set; }
        public int UserId { get; set; }
    }
}