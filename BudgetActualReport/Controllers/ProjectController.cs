using BudgetActualReport.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Net;
using System.IO;
using ClosedXML.Excel;
using System.Globalization;

namespace BudgetActualReport.Controllers
{
    public class ProjectController : Controller
    {
        string errors1 = "";
        public ActionResult ProjectIndex(int CompanyId)//72
        {
            ViewBag.CompId = CompanyId;
            var _projects = GetProjects(CompanyId);
            ViewBag.Projects = _projects;
            return View();
        }

        public IEnumerable<SelectListItem> GetProjects(int cid)
        {
            string retrievequery = string.Format(@"exec pCore_CommonSp @Operation=getProjects");
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

        public ActionResult BudgetActualReport(int CompanyId, int Project, string ReportDate,string ProjectName)
        {
            TempData["CompanyId"] = CompanyId;
            TempData["Project"] = Project;
            TempData["ReportDate"] = ReportDate;
            TempData["ProjectName"] = ProjectName;
            DateTime reportDt = Convert.ToDateTime(ReportDate);
            TempData.Keep();
            #region QueryData


            string Strsql = string.Format($@"exec pCore_CommonSp @Operation = getReportData,@p1 = {Project},@p3='{ReportDate}'");
            DataSet ds = DBClass.GetDataSet(Strsql, CompanyId, ref errors1);

            int table = ds.Tables.Count;
            Budget_Actual_Analysis reportObj = new Budget_Actual_Analysis();
            List<BudgetVsActual> listobj = new List<BudgetVsActual>();
            if (ds.Tables[0].Rows.Count>0) {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    listobj.Add(new BudgetVsActual
                    {
                        Category = dr["Category"].ToString(),
                        Budget = Convert.ToDecimal(dr["Budget"].ToString()),
                        NonPo = Convert.ToDecimal(dr["NonPO"].ToString()),
                        PO = Convert.ToDecimal(dr["PO"].ToString()),
                        Forcast = 0,
                        Save = 0,
                    });
                }
                reportObj.budgetvsactuallist = listobj;
                Analysis objA = new Analysis();
                objA.Initial_Order_Value = Convert.ToDecimal(ds.Tables[0].Rows[0]["OrderValue"].ToString());
                objA.Variation = Convert.ToDecimal(ds.Tables[0].Rows[0]["Variation"].ToString());
                objA.ActualCost = Convert.ToDecimal(ds.Tables[0].Rows[0]["ActualCost"].ToString());
                objA.InvoicedTillDate = Convert.ToDecimal(ds.Tables[0].Rows[0]["InvoiceTillDate"].ToString());
                objA.Pending = Convert.ToDecimal(ds.Tables[0].Rows[0]["Pending"].ToString());
                objA.Received = Convert.ToDecimal(ds.Tables[0].Rows[0]["Received"].ToString());
                objA.Retension = Convert.ToDecimal(ds.Tables[0].Rows[0]["Retension"].ToString());
                objA.Outstanding = Convert.ToDecimal(ds.Tables[0].Rows[0]["Outstanding"].ToString());
                objA.Total_Sales_Value = Convert.ToDecimal(ds.Tables[0].Rows[0]["TotalSalesValue"].ToString());
                objA.Total = Convert.ToDecimal(ds.Tables[0].Rows[0]["Total"].ToString());
                objA.TotalCost = Convert.ToDecimal(ds.Tables[0].Rows[0]["ActualCost"].ToString());
                reportObj.analysis = objA;
                reportObj.Projects = ProjectName;
                reportObj.ReportDate = ReportDate;
                reportObj.CompanyId = CompanyId;
            }
            #endregion
            return View(reportObj);
        }

        [HttpPost]
        public FileResult ExcelGenerate()
        {

            #region TempData
            int CompanyId = Convert.ToInt32(TempData["CompanyId"]);
            string ReportDate = Convert.ToString(TempData["ReportDate"]);
            int Currency = Convert.ToInt32(TempData["Currency"]);
            int Project = Convert.ToInt32(TempData["Project"]);
            string ProjectName = Convert.ToString(TempData["ProjectName"]);
            DateTime reportDt = Convert.ToDateTime(ReportDate);
            TempData.Keep();
            #endregion

            System.Data.DataTable data = new System.Data.DataTable("BUDGET vs ACTUAL REPORT");
            #region DataColumns
            data.Columns.Add("Sn", typeof(string));
            data.Columns.Add("Category", typeof(string));
            data.Columns.Add("Budget", typeof(string));
            data.Columns.Add("Non PO's", typeof(string));
            data.Columns.Add("PO's", typeof(string));
            data.Columns.Add("Forecast", typeof(string));
            data.Columns.Add("Save/(Loss)", typeof(string));
            
            #endregion

            #region QueryData
            string Strsql = string.Format($@"exec pCore_CommonSp @Operation = getReportData,@p1 = {Project},@p3='{ReportDate}'");
            DataSet ds = DBClass.GetDataSet(Strsql, CompanyId, ref errors1);
            
            int table = ds.Tables.Count;
            Budget_Actual_Analysis reportObj = new Budget_Actual_Analysis();
            List<BudgetVsActual> listobj = new List<BudgetVsActual>();
            Analysis objA = new Analysis();
            if (ds.Tables[0].Rows.Count > 0)
            {
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    listobj.Add(new BudgetVsActual
                    {
                        Category = dr["Category"].ToString(),
                        Budget = Convert.ToDecimal(dr["Budget"].ToString()),
                        NonPo = Convert.ToDecimal(dr["NonPO"].ToString()),
                        PO = Convert.ToDecimal(dr["PO"].ToString()),
                        Forcast = 0,
                        Save = 0,
                    });
                }
                reportObj.budgetvsactuallist = listobj;
                
                objA.Initial_Order_Value = Convert.ToDecimal(ds.Tables[0].Rows[0]["OrderValue"].ToString());
                objA.Variation = Convert.ToDecimal(ds.Tables[0].Rows[0]["Variation"].ToString());
                objA.ActualCost = Convert.ToDecimal(ds.Tables[0].Rows[0]["ActualCost"].ToString());
                objA.InvoicedTillDate = Convert.ToDecimal(ds.Tables[0].Rows[0]["InvoiceTillDate"].ToString());
                objA.Pending = Convert.ToDecimal(ds.Tables[0].Rows[0]["Pending"].ToString());
                objA.Received = Convert.ToDecimal(ds.Tables[0].Rows[0]["Received"].ToString());
                objA.Retension = Convert.ToDecimal(ds.Tables[0].Rows[0]["Retension"].ToString());
                objA.Outstanding = Convert.ToDecimal(ds.Tables[0].Rows[0]["Outstanding"].ToString());
                objA.Total_Sales_Value = Convert.ToDecimal(ds.Tables[0].Rows[0]["TotalSalesValue"].ToString());
                objA.Total = Convert.ToDecimal(ds.Tables[0].Rows[0]["Total"].ToString());
                objA.TotalCost = Convert.ToDecimal(ds.Tables[0].Rows[0]["ActualCost"].ToString());
                reportObj.analysis = objA;
                reportObj.Projects = ProjectName;
                reportObj.ReportDate = ReportDate;
                reportObj.CompanyId = CompanyId;
            }
            #endregion

            using (XLWorkbook wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("BUDGET vs ACTUAL REPORT");
                var dataTable = data;

                

                var wsReportNameHeaderRange = ws.Range(ws.Cell(1, 2), ws.Cell(1, 8));//ws.Range(string.Format("A{0}:{1}{0}", 1, Char.ConvertFromUtf32(aCode + dataTable.Columns.Count + 1)));
                wsReportNameHeaderRange.Style.Font.Bold = true;
                wsReportNameHeaderRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Left;
                wsReportNameHeaderRange.Style.Fill.BackgroundColor = XLColor.Yellow;
                wsReportNameHeaderRange.Merge();
                wsReportNameHeaderRange.Value = "                                                                             BUDGET vs ACTUAL REPORT          ";

                int cell = 2;
                for (int i = 1; i < data.Columns.Count; i++)
                {
                    cell = 2;
                    #region Headers
                    ws.Cell(4, cell++).Value = "Sn";
                    ws.Cell(4, cell++).Value = "Category";
                    ws.Cell(4, cell++).Value = "Budget";
                    ws.Cell(4, cell++).Value = "Non PO's";
                    ws.Cell(4, cell++).Value = "PO's";
                    ws.Cell(4, cell++).Value = "Forecast";
                    ws.Cell(4, cell++).Value = "Save/(Loss)";
                    
                    #endregion
                }
                var TableRange = ws.Range(ws.Cell(4, 2), ws.Cell(4, 8));
                TableRange.Style.Font.FontColor = XLColor.White;
                TableRange.Style.Fill.BackgroundColor = XLColor.FromArgb(0, 115, 170);
                TableRange.Style.Font.Bold = true;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                int r = 5;
                int c = 2;

                #region TableLoop
                var list = reportObj.budgetvsactuallist;
                decimal TotalBudget = 0;
                decimal TotalNonPO = 0;
                decimal TotalPO = 0;
                int count = 1;
                foreach (var obj in list)
                {

                    TotalBudget = Convert.ToDecimal(TotalBudget) + obj.Budget;
                    TotalNonPO = Convert.ToDecimal(TotalNonPO) + obj.NonPo;
                    TotalPO = Convert.ToDecimal(TotalPO) + obj.PO;
                    c = 2;
                    ws.Range(ws.Cell(r, c), ws.Cell(r, 8)).Style.Fill.BackgroundColor = XLColor.FromHtml("#90EE90");
                    ws.Cell(r, c++).Value = count++;
                    ws.Cell(r, c++).Value = obj.Category;
                    ws.Cell(r, c++).Value = obj.Budget;
                    ws.Cell(r, c++).Value = obj.NonPo;
                    ws.Cell(r, c++).Value = obj.PO;
                    ws.Cell(r, c++).Value = 0;
                    ws.Cell(r, c++).Value = 0;
                    r++;

                }


                //Grand Total Row
                c = 2;
                ws.Cell(r, c++).Value = "";
                ws.Cell(r, c++).Value = "Grand Total";
                ws.Cell(r, c++).Value = TotalBudget.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, c++).Value = TotalNonPO.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, c++).Value = TotalPO.ToString("N", new CultureInfo("en-US"));
                ws.Cell(r, c++).Value = "";
                ws.Cell(r, c++).Value = "";
                
                ws.Range("B" + r + ":Z" + r + "").Style.Font.Bold = true;
                r++;


                #endregion

                TableRange = ws.Range(ws.Cell(4, 2), ws.Cell(r - 1, 8));
                TableRange.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                TableRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

                //ws.Range(ws.Cell(4, 3), ws.Cell(r, 8)).Style.NumberFormat.Format = "0.00";

                //ws.Columns("A:BZ").AdjustToContents();
                r = r + 3;
                c = 3;
                var TableRange1 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange1.Style.Font.FontColor = XLColor.Black;
                TableRange1.Style.Fill.BackgroundColor = XLColor.White;
                TableRange1.Style.Font.Bold = false;
                TableRange1.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange1.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange1.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "INITIAL ORDER VALUE : ";
                ws.Cell(r, c++).Value = objA.Initial_Order_Value.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange7 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange7.Style.Font.FontColor = XLColor.Black;
                TableRange7.Style.Fill.BackgroundColor = XLColor.White;
                TableRange7.Style.Font.Bold = false;
                TableRange7.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange7.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange7.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "ACTUAL  COST : ";
                ws.Cell(r, c++).Value = objA.ActualCost.ToString("N", new CultureInfo("en-US"));

                



                r++;
                c = 3;
                var TableRange2 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange2.Style.Font.FontColor = XLColor.Black;
                TableRange2.Style.Fill.BackgroundColor = XLColor.White;
                TableRange2.Style.Font.Bold = false;
                TableRange2.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange2.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange2.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "VARIATIONS : ";
                ws.Cell(r, c++).Value = objA.Variation.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange8 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange8.Style.Font.FontColor = XLColor.Black;
                TableRange8.Style.Fill.BackgroundColor = XLColor.White;
                TableRange8.Style.Font.Bold = false;
                TableRange8.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange8.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange8.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "FORECASTED : ";
                ws.Cell(r, c++).Value = objA.Forcasted.ToString("N", new CultureInfo("en-US"));

                r++;
                c = 3;
                var TableRange3 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange3.Style.Font.FontColor = XLColor.Black;
                TableRange3.Style.Fill.BackgroundColor = XLColor.White;
                TableRange3.Style.Font.Bold = true;
                TableRange3.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange3.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange3.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "TOTAL SALE VALUE : ";
                ws.Cell(r, c++).Value = objA.Total_Sales_Value.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange9 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange9.Style.Font.FontColor = XLColor.Black;
                TableRange9.Style.Fill.BackgroundColor = XLColor.White;
                TableRange9.Style.Font.Bold = true;
                TableRange9.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange9.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange9.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "TOTAL COST : ";
                ws.Cell(r, c++).Value = objA.TotalCost.ToString("N", new CultureInfo("en-US"));


                r= r+3;
                c = 3;
                var TableRange4 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange4.Style.Font.FontColor = XLColor.Black;
                TableRange4.Style.Fill.BackgroundColor = XLColor.White;
                TableRange4.Style.Font.Bold = false;
                TableRange4.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange4.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange4.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "INVOICED TILL DATE : ";
                ws.Cell(r, c++).Value = objA.InvoicedTillDate.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange0 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange0.Style.Font.FontColor = XLColor.Black;
                TableRange0.Style.Fill.BackgroundColor = XLColor.White;
                TableRange0.Style.Font.Bold = false;
                TableRange0.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange0.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange0.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "RECEIVED : ";
                ws.Cell(r, c++).Value = objA.Received.ToString("N", new CultureInfo("en-US"));


                r++;
                c = 3;
                var TableRange5 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange5.Style.Font.FontColor = XLColor.Black;
                TableRange5.Style.Fill.BackgroundColor = XLColor.White;
                TableRange5.Style.Font.Bold = false;
                TableRange5.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange5.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange5.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "PENDING : ";
                ws.Cell(r, c++).Value = objA.Pending.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange11 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange11.Style.Font.FontColor = XLColor.Black;
                TableRange11.Style.Fill.BackgroundColor = XLColor.White;
                TableRange11.Style.Font.Bold = false;
                TableRange11.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange11.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange11.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "RETENTION : ";
                ws.Cell(r, c++).Value = objA.Retension.ToString("N", new CultureInfo("en-US"));


                r++;
                c = 3;
                var TableRange6 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange6.Style.Font.FontColor = XLColor.Black;
                TableRange6.Style.Fill.BackgroundColor = XLColor.White;
                TableRange6.Style.Font.Bold = true;
                TableRange6.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange6.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange6.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "TOTAL : ";
                ws.Cell(r, c++).Value = objA.Total.ToString("N", new CultureInfo("en-US"));

                c = 8;
                var TableRange12 = ws.Range(ws.Cell(r, c), ws.Cell(r, c+1));
                TableRange12.Style.Font.FontColor = XLColor.Black;
                TableRange12.Style.Fill.BackgroundColor = XLColor.White;
                TableRange12.Style.Font.Bold = false;
                TableRange12.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                TableRange12.Style.Border.InsideBorder = XLBorderStyleValues.Thin;
                TableRange12.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Cell(r, c++).Value = "OUTSTANDING : ";
                ws.Cell(r, c++).Value = objA.Outstanding.ToString("N", new CultureInfo("en-US"));
                ws.Range(ws.Cell(4, 3), ws.Cell(r, 13)).Style.NumberFormat.Format = "0.00";
                ws.Columns("A:BZ").AdjustToContents();
                using (MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "BudgetvsActualReport" + "_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx");
                }
            }
        }
    }
}