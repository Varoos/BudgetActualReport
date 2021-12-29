using ClosedXML.Excel;
using BudgetActualReport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace BudgetActualReport.Controllers
{
    public class AgeingController : Controller
    {
        string errors1 = "";
        string Message = "";
        // GET: Ageing
        public ActionResult Index(int CompanyId)
        {
            ViewBag.CompId = CompanyId;
            var _customers = GetCustomers(CompanyId);
            ViewBag.Customers = _customers;
            var _salesmans = GetSalesMans(CompanyId);
            ViewBag.SalesMans = _salesmans;
            var _currencies = GetCurrencies(CompanyId);
            ViewBag.Currencies = _currencies;
            return View();
        }

        public IEnumerable<SelectListItem> GetSalesMans(int cid)
        {
            string retrievequery = string.Format(@"select distinct s.iMasterId,sName from muCore_Account_Details ad join mCore_Salesman s on ad.Salesman=s.iMasterId where s.iMasterId<>0 order by s.iMasterId");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                containers.Add(new SelectListItem()
                {
                    Text = ds.Tables[0].Rows[i]["sName"].ToString(),
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                });
            }
            return new SelectList(containers.ToArray(), "Value", "Text");
        }

        public IEnumerable<SelectListItem> GetCustomers(int cid)
        {
            string retrievequery = string.Format(@"select iMasterId,sName from mCore_Account  where iMasterId<>0 and iStatus<>5 and bGroup=1");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                containers.Add(new SelectListItem()
                {
                    Text = ds.Tables[0].Rows[i]["sName"].ToString(),
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                });
            }

            return new SelectList(containers.ToArray(), "Value", "Text");
        }

        public IEnumerable<SelectListItem> GetCurrencies(int cid)
        {
            string retrievequery = string.Format(@"select iCurrencyId,sName from mCore_Currency  where iCurrencyId<>0");
            List<SelectListItem> containers = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                containers.Add(new SelectListItem()
                {
                    Text = ds.Tables[0].Rows[i]["sName"].ToString(),
                    Value = ds.Tables[0].Rows[i]["iCurrencyId"].ToString(),
                });
            }
            var containeritem = new SelectListItem()
            {
                Text = "-- Please Select Currency --"
            };
            containers.Insert(0, containeritem);
            return new SelectList(containers.ToArray(), "Value", "Text");
        }

        public ActionResult GetData(int CompanyId, int SelectValue)
        {
            var _customers = GetSelectionCustomers(CompanyId, SelectValue);

            return Json(new { Customers = _customers }, JsonRequestBehavior.AllowGet);
        }
        public IEnumerable<SelectListItem> GetSelectionCustomers(int cid, int SelectValue)
        {
            string retrievequery = "";
            if (SelectValue == 1)
            {
                retrievequery = @"select iMasterId,sName from mCore_Account  where iMasterId<>0 and iStatus<>5 and bGroup=1";
            }
            else
            {
                retrievequery = @"select iMasterId,sName from mCore_Account  where iMasterId<>0 and iStatus<>5 and bGroup=0";
            }
            List<SelectListItem> stores = new List<SelectListItem>();
            DataSet ds = DBClass.GetData(retrievequery, cid, ref errors1);

            for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
            {
                stores.Add(new SelectListItem()
                {
                    Value = ds.Tables[0].Rows[i]["iMasterId"].ToString(),
                    Text = ds.Tables[0].Rows[i]["sName"].ToString()
                });
            }
            return new SelectList(stores.AsEnumerable(), "Value", "Text");
        }

        public ActionResult AgeingReport(int CompanyId, string Accounts, string SalesMans, string Months, string ReportDate, int Currency, int SelectValue, bool showPrint = true)
        {
            TempData["CompanyId"] = CompanyId;
            TempData["Accounts"] = Accounts;
            TempData["SalesMans"] = SalesMans;
            TempData["Months"] = Months;
            TempData["ReportDate"] = ReportDate;
            TempData["Currency"] = Currency;
            TempData["SelectValue"] = SelectValue;
            var MonthsList = Months.Split(',').ToList();
            var nosmCount = MonthsList.Count();
            if (MonthsList[0] == "")
            {
                nosmCount = 0;
            }

            DateTime reportDt = Convert.ToDateTime(ReportDate);

            int date = reportDt.Day;
            int month = reportDt.Month;
            int year = reportDt.Year;
            if (date > 28)
            {
                DateTime ConvertedDt = new DateTime(year, month, 28);
                string stringDt = ConvertedDt.ToString("yyyy-MM-dd");
                ReportDate = stringDt;
            }

            ViewBag.NoOfSelectedMonthsCount = nosmCount;
            ViewBag.NoOfSelectedMonths = MonthsList;
            TempData["NoOfSelectedMonths"] = MonthsList;
            TempData["NoOfSelectedMonthsCount"] = nosmCount;

            #region NoOfMonths
            string Monthquery = string.Format(@"DECLARE @Today DATETIME,@Date DATETIME, @nMonths TINYINT
                                        SET @Date='{0}'
                                        SET @nMonths = 11
                                        SET @Today = DATEADD(month, (-1) * @nMonths, @Date)

                                        ;WITH q AS
                                        (
	                                        SELECT  @Today AS datum
	                                        UNION ALL
	                                        SELECT  DATEADD(month, 1, datum) 
	                                        FROM q WHERE datum  < @Date
                                        )
                                        SELECT  UPPER(SUBSTRING(DATENAME(MONTH, datum), 1, 3)) [Month],SUBSTRING(DATENAME(MONTH, datum), 1, 3)+' -' + CAST(YEAR(datum) AS VARCHAR(4)) [MonthYear]
                                        FROM q", ReportDate);
            DataSet dsrt = DBClass.GetData(Monthquery, CompanyId, ref errors1);
            var rtList = dsrt.Tables[0].AsEnumerable().Select(r => new Months { Month = r.Field<string>("Month"), MonthYear = r.Field<string>("MonthYear") });
            var MonthNames = rtList.Select(_ => _.Month).Reverse().ToList();
            var MonthYear = rtList.Select(_ => _.MonthYear).Reverse().ToList();
            ViewBag.DynamicMonths = MonthNames;
            ViewBag.DynamicMonthYrs = MonthYear;
            TempData["showprint"] = showPrint;
            TempData["DynamicMonths"] = MonthNames;
            TempData["DynamicMonthYrs"] = MonthYear;
            TempData.Keep();
            #endregion

            #region QueryData
            string filters = "";
            string AccIds = "";
            if (SelectValue == 1)
            {
                #region AcctQuery
                //string accquery = string.Format(@"select iMasterId from vmCore_Account where iParentId in(" + Accounts + ") and bGroup=0");
                string accquery = string.Format(@"select p.iMasterId from vmCore_Account p join fCore_GetAccountHierarchy(" + Accounts + ",0) f on p.iMasterId = f.iMasterId where iStatus<> 5 and p.iParentId<> 0 and bGroup=0");
                DataSet dsacc = DBClass.GetData(accquery, CompanyId, ref errors1);
                if (dsacc.Tables[0] != null)
                {
                    if (dsacc.Tables[0].Rows.Count > 0)
                    {
                        if (dsacc.Tables.Count != 0)
                        {
                            for (int i = 0; i < dsacc.Tables[0].Rows.Count; i++)
                            {
                                AccIds = AccIds + dsacc.Tables[0].Rows[i]["iMasterId"].ToString() + ",";
                            }
                        }
                    }
                }
                #endregion
                AccIds = AccIds.TrimEnd(',');
            }
            else
            {
                AccIds = Accounts;
            }

            if (Accounts != "")
            {
                filters = " where AccId in(" + string.Join(",", AccIds) + ")";
            }
            //if (SalesMans != "")
            //{
            //    var sm = $"'{SalesMans.Replace(",", "', '")}'";
            //    filters += string.IsNullOrEmpty(filters?.Trim())
            //        ? $"where SalesMan in(" + sm + ")"
            //        : $" and SalesMan  in(" + sm + ")";
            //}

            //string Strsql = string.Format(@"select * from [vu_Core_MonthAgingWithoutPDC] where AccId = 1075 order by iVoucherType, sVoucherNo");// + filters + "");
            string Strsql = string.Format($@"exec Core_MonthAgingWithoutPDC_SP @TagID = '{string.Join(", ", AccIds)}'");
            DataSet ds = DBClass.GetDataSet(Strsql, CompanyId, ref errors1);

            int table = ds.Tables.Count;
            List<Ageing> reportlist = new List<Ageing>();

            for (int i = 0; i < table; i++)
            {
                foreach (DataRow dr in ds.Tables[i].Rows)
                {
                    reportlist.Add(new Ageing
                    {
                        AccountName = dr["AccName"].ToString(),
                        VoucherName = dr["sVoucherNo"].ToString(),
                        SalesMan = dr["SalesMan"].ToString(),
                        LPONo = dr["LPO_No"].ToString(),
                        InvoiceAmt = Convert.ToDecimal(Convert.ToDecimal(dr["Invoice Amount"]).ToString("#,##0.00")),
                        BalanceAmt = Convert.ToDecimal(Convert.ToDecimal(dr["Balance Amount"]).ToString("#,##0.00")),
                        Date = dr["ConvertedBillDate"].ToString(),
                        DelayDays = Convert.ToInt32(dr["DelayDays"]),
                        Month = dr["Month"].ToString(),
                    });
                }
            }
            #endregion
            AgeingAnalysis _viewmodel = new AgeingAnalysis();
           
            _viewmodel.Ageing = reportlist;
            return View(_viewmodel);
        }

        [HttpPost]
        public FileResult ExcelGenerate()
        {
            var aCode = 65;

            #region TempData
            int CompanyId = Convert.ToInt32(TempData["CompanyId"]);
            string Accounts = Convert.ToString(TempData["Accounts"]);
            string SalesMans = Convert.ToString(TempData["SalesMans"]);
            string Months = Convert.ToString(TempData["Months"]);
            string ReportDate = Convert.ToString(TempData["ReportDate"]);
            int Currency = Convert.ToInt32(TempData["Currency"]);
            int SelectValue = Convert.ToInt32(TempData["SelectValue"]);

            List<string> MonthNames = (List<string>)TempData["DynamicMonths"];
            List<string> MonthYear = (List<string>)TempData["DynamicMonthYrs"];
            List<string> MonthsList = (List<string>)TempData["NoOfSelectedMonths"];
            int NoOfSelectedMonthsCount = Convert.ToInt32(TempData["NoOfSelectedMonthsCount"]);
            TempData.Keep();
            #endregion

            System.Data.DataTable data = new System.Data.DataTable("Ageing Report");
            #region DataColumns
            data.Columns.Add("PARTICULARS", typeof(string));
            data.Columns.Add("SALESMAN", typeof(string));
            data.Columns.Add("LPONO", typeof(string));
            data.Columns.Add("INVOICE AMOUNT", typeof(string));
            data.Columns.Add("BALANCE AMOUNT", typeof(string));
            data.Columns.Add("DATE", typeof(string));
            data.Columns.Add("DELAY DAYS", typeof(string));
            if (NoOfSelectedMonthsCount > 0)
            {
                foreach (string _nom in MonthNames)
                {
                    foreach (var sm in MonthsList)
                    {
                        if (_nom.ToLower() == sm)
                        {
                            data.Columns.Add(_nom, typeof(string));
                        }
                    }
                }
            }
            else
            {
                foreach (var _mn in MonthNames)
                {
                    data.Columns.Add(_mn, typeof(string));
                }
            }
            if (MonthsList.Count > 11){
                data.Columns.Add("More Than One Year", typeof(string));
            }
            data.Columns.Add("Total", typeof(string));
            #endregion

            string Strsql = "";
            DataSet ds = new DataSet();

            #region QueryData
            string filters = "";
            string AccIds = "";
            if (SelectValue == 1)
            {
                #region AcctQuery
                //string accquery = string.Format(@"select iMasterId from vmCore_Account where iParentId in(" + Accounts + ") and bGroup=0");
                string accquery = string.Format(@"select p.iMasterId from vmCore_Account p join fCore_GetAccountHierarchy(" + Accounts + ",0) f on p.iMasterId = f.iMasterId where iStatus<> 5 and p.iParentId<> 0 and bGroup=0");
                DataSet dsacc = DBClass.GetData(accquery, CompanyId, ref errors1);
                if (dsacc.Tables[0] != null)
                {
                    if (dsacc.Tables[0].Rows.Count > 0)
                    {
                        if (dsacc.Tables.Count != 0)
                        {
                            for (int i = 0; i < dsacc.Tables[0].Rows.Count; i++)
                            {
                                AccIds = AccIds + dsacc.Tables[0].Rows[i]["iMasterId"].ToString() + ",";
                            }
                        }
                    }
                }
                #endregion
                AccIds = AccIds.TrimEnd(',');
            }
            else
            {
                AccIds = Accounts;
            }

            if (Accounts != "")
            {
                filters = " where AccId in(" + string.Join(",", AccIds) + ")";
            }

            //if (SalesMans != "")
            //{
            //    var sm = $"'{SalesMans.Replace(",", "', '")}'";
            //    filters += string.IsNullOrEmpty(filters?.Trim())
            //        ? $"where SalesMan in(" + sm + ")"
            //        : $" and SalesMan  in(" + sm + ")";
            //}

            //Strsql = string.Format(@"select * from [vu_Core_MonthAgingWithoutPDC] " + filters + "");
             Strsql = string.Format($@"exec Core_MonthAgingWithoutPDC_SP @TagID = '{string.Join(", ", AccIds)}'");
            ds = DBClass.GetDataSet(Strsql, CompanyId, ref errors1);

            int table = ds.Tables.Count;
            List<Ageing> reportlist = new List<Ageing>();

            for (int i = 0; i < table; i++)
            {
                foreach (DataRow dr in ds.Tables[i].Rows)
                {
                    reportlist.Add(new Ageing
                    {
                        AccountName = dr["AccName"].ToString(),
                        VoucherName = dr["sVoucherNo"].ToString(),
                        SalesMan = dr["SalesMan"].ToString(),
                        LPONo = dr["LPO_No"].ToString(),
                        InvoiceAmt = Convert.ToDecimal(Convert.ToDecimal(dr["Invoice Amount"]).ToString("#,##0.00")),
                        BalanceAmt = Convert.ToDecimal(Convert.ToDecimal(dr["Balance Amount"]).ToString("#,##0.00")),
                        Date = dr["ConvertedBillDate"].ToString(),
                        DelayDays = Convert.ToInt32(dr["DelayDays"]),
                        Month = dr["Month"].ToString(),
                    });
                }
            }

            #endregion

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Ageing Report");
                var dataTable = data;

                var ListCount = 0;
                if (NoOfSelectedMonthsCount > 0)
                {
                    ListCount = MonthsList.Count();
                }
                else
                {
                    ListCount = 12;
                }

                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 10 + ListCount));//ws.Range(string.Format("A{0}:{1}{0}", 1, Char.ConvertFromUtf32(aCode + dataTable.Columns.Count + 1)));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "                                                                                                               AGEING REPORT     ";

                int cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(4, cell++).Value = "PARTICULARS";
                    ws.Cell(4, cell++).Value = "SALESMAN";
                    ws.Cell(4, cell++).Value = "LPONO";
                    ws.Cell(4, cell++).Value = "INVOICE AMOUNT";
                    ws.Cell(4, cell++).Value = "BALANCE AMOUNT";
                    ws.Cell(4, cell++).Value = "DATE";
                    ws.Cell(4, cell++).Value = "DELAY DAYS";
                    if (NoOfSelectedMonthsCount > 0)
                    {
                        foreach (string _nom in MonthNames)
                        {
                            foreach (var sm in MonthsList)
                            {
                                if (_nom.ToLower() == sm)
                                {
                                    ws.Cell(4, cell++).Value = _nom;
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (var _mn in MonthNames)
                        {
                            ws.Cell(4, cell++).Value = _mn;
                        }
                    }
                    if (MonthsList.Count > 11)
                    {
                        ws.Cell(4, cell++).Value = "More Than One Year";
                    }
                    ws.Cell(4, cell++).Value = "Total";
                    #endregion
                }
                var TableRange = ws.Range(ws.Cell(4, 2), ws.Cell(4, 10 + ListCount));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int r = 5;
                int c = 2;

                #region TableLoop
                var CustomersGroup = reportlist.GroupBy(_ => _.AccountName);
                decimal GrandTotalZeroSum = 0;
                decimal GrandTotalRowSum = 0;
                decimal GrandInvoiceAmtTotal = 0;
                decimal GrandBalanceAmtTotal = 0;
                foreach (var _customer in CustomersGroup)
                {
                    string TotalRowSum = "";
                    decimal TotalZeroSum = 0;
                    decimal FinalTotalRowSum = 0;
                   
                    c = 2;
                    ws.Cell(r, c).Value = _customer.Key;
                    ws.Cell(r, c).Style.Font.Bold = true;
                    ws.Range(ws.Cell(r, c), ws.Cell(r, 10 + ListCount)).Style.Fill.BackgroundColor = XLColor.FromHtml("#90EE90");
                    r++;

                    foreach (var _cust in _customer)
                    {
                        var count = 0;
                        c = 2;

                        ws.Cell(r, c++).Value = _cust.VoucherName;
                        ws.Cell(r, c++).Value = _cust.SalesMan;
                        ws.Cell(r, c++).Value = _cust.LPONo;
                        ws.Cell(r, c++).Value = _cust.InvoiceAmt.ToString("N", new CultureInfo("en-US"));
                        ws.Cell(r, c++).Value = _cust.BalanceAmt.ToString("N", new CultureInfo("en-US"));
                        ws.Cell(r, c++).Value = "'" + _cust.Date;
                        ws.Cell(r, c++).Value = _cust.DelayDays;
                        if (NoOfSelectedMonthsCount > 0)
                        {
                            TotalRowSum = _cust.BalanceAmt.ToString("N", new CultureInfo("en-US"));
                            FinalTotalRowSum = Convert.ToDecimal(FinalTotalRowSum) + Convert.ToDecimal(_cust.BalanceAmt.ToString("N", new CultureInfo("en-US")));
                            GrandTotalRowSum = Convert.ToDecimal(GrandTotalRowSum) + FinalTotalRowSum;
                            foreach (var _nom in MonthYear)
                            {
                                foreach (var sm in MonthsList)
                                {
                                    if (sm == _nom.ToLower().Remove(_nom.Length - 6))
                                    {
                                        if (_cust.Month == _nom)
                                        {
                                            ws.Cell(r, c++).Value = _cust.BalanceAmt.ToString("N", new CultureInfo("en-US"));
                                            count = 1;
                                        }
                                        else
                                        {
                                            ws.Cell(r, c++).Value = "";
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            foreach (var _nom in MonthYear)
                            {
                                if (_cust.Month == _nom)
                                {
                                    ws.Cell(r, c++).Value = _cust.BalanceAmt.ToString("N", new CultureInfo("en-US"));
                                    count = 1;
                                }
                                else
                                {
                                    ws.Cell(r, c++).Value = "";
                                }
                            }
                        }
                        if (MonthsList.Count > 11)
                        {
                            if (count == 0)
                            {
                                TotalZeroSum = Convert.ToDecimal(TotalZeroSum) + Convert.ToDecimal(_cust.BalanceAmt.ToString("N", new CultureInfo("en-US")));
                                ws.Cell(r, c++).Value = _cust.BalanceAmt.ToString("N", new CultureInfo("en-US"));
                            }
                            else
                            {
                                ws.Cell(r, c++).Value = "";
                            }
                        }
                        
                        ws.Cell(r, c++).Value = TotalRowSum;
                        r++;
                    }

                    var InvoiceAmtTotal = _customer.Sum(_ => _.InvoiceAmt);
                    var BalanceAmtTotal = _customer.Sum(_ => _.BalanceAmt);
                    GrandInvoiceAmtTotal = Convert.ToDecimal(GrandInvoiceAmtTotal) + InvoiceAmtTotal;
                    GrandBalanceAmtTotal = Convert.ToDecimal(GrandBalanceAmtTotal) + BalanceAmtTotal;
                    c = 2;
                    ws.Cell(r, c++).Value = "Sub Total";
                    ws.Cell(r, c++).Value = "";
                    ws.Cell(r, c++).Value = "";
                    ws.Cell(r, c++).Value = InvoiceAmtTotal.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = BalanceAmtTotal.ToString("N", new CultureInfo("en-US"));
                    ws.Cell(r, c++).Value = "";
                    ws.Cell(r, c++).Value = "";
                    if (NoOfSelectedMonthsCount > 0)
                    {
                        foreach (string _nom in MonthYear)
                        {
                            foreach (var sm in MonthsList)
                            {
                                if (sm == _nom.ToLower().Remove(_nom.Length - 6))
                                {
                                    var CustMonthWise = _customer.Where(_ => _.Month.ToLower().Remove(_nom.Length - 6).Contains(_nom.ToLower().Remove(_nom.Length - 6))).Sum(_ => _.BalanceAmt);
                                    if (Math.Abs(CustMonthWise) > 0)
                                    {
                                        ws.Cell(r, c++).Value = CustMonthWise.ToString("#,##0.00");
                                    }
                                    else
                                    {
                                        ws.Cell(r, c++).Value = "0.00";
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (string _nom in MonthYear)
                        {
                            var CustMonthWise = _customer.Where(_ => _.Month == _nom).Sum(_ => _.BalanceAmt);
                            ws.Cell(r, c++).Value = CustMonthWise.ToString("N", new CultureInfo("en-US"));
                        }
                    }
                    if (MonthsList.Count > 11)
                    {
                        GrandTotalZeroSum = Convert.ToDecimal(GrandTotalZeroSum) + TotalZeroSum;
                        ws.Cell(r, c++).Value = TotalZeroSum.ToString("N", new CultureInfo("en-US"));
                    }
                    
                    ws.Cell(r, c++).Value =FinalTotalRowSum;
                    ws.Range("B" + r + ":Z" + r + "").Style.Font.Bold = true;
                    r++;
                }

                //Grand Total Row
                c = 2;
                ws.Cell(r, c++).Value = "Grand Total";
                ws.Cell(r, c++).Value = "";
                ws.Cell(r, c++).Value = "";
                ws.Cell(r, c++).Value = GrandInvoiceAmtTotal.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, c++).Value = GrandBalanceAmtTotal.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, c++).Value = "";
                ws.Cell(r, c++).Value = "";
                if (NoOfSelectedMonthsCount > 0)
                {
                    foreach (string _nom in MonthYear)
                    {
                        foreach (var sm in MonthsList)
                        {
                            if (sm == _nom.ToLower().Remove(_nom.Length - 6))
                            {
                                ws.Cell(r, c++).Value = "0.00";
                            }
                        }
                    }
                }
               
                if (MonthsList.Count > 11)
                {
                    ws.Cell(r, c++).Value = GrandTotalZeroSum.ToString("N", new CultureInfo("en-US"));
                }

                ws.Cell(r, c++).Value = GrandTotalRowSum.ToString("N", new CultureInfo("en-US"));
                ws.Range("B" + r + ":Z" + r + "").Style.Font.Bold = true;
                r++;


                #endregion

                TableRange = ws.Range(ws.Cell(4, 2), ws.Cell(r - 1, 10 + ListCount));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                ws.Range(ws.Cell(4, 2), ws.Cell(r, 10 + ListCount)).Style.NumberFormat.Format = "0.00";

                ws.Columns("A:BZ").AdjustToContents();

                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "AgeingReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
        public static int ColumnNameToIndex(string name)
{
  var upperCaseName = name.ToUpper();
  var number = 0;
  var pow = 1;
  for (var i = upperCaseName.Length - 1; i >= 0; i--)
  {
    number += (upperCaseName[i] - 'A' + 1) * pow;
    pow *= 26;
  }

  return number;
}
    }
}