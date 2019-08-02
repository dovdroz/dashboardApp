using System.Web;
using System.Web.Mvc;
using Microsoft.Office.Interop.Excel;
using DashApp.Models;
using System.Collections.Generic;

namespace DashApp.Controllers
{
    public class HomeController : Controller
    {
        static List<GetVal> val = new List<GetVal>();
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        //[AcceptVerbs(HttpVerbs.Get)]
        public ActionResult Import(HttpPostedFileBase excelfile, GetVal valu)
        {
            bool dateBool = false;


            //check for valid dates
            if ((valu.StartDate != null) && (valu.EndDate != null))
            {
                if (valu.StartDate > valu.EndDate)
                {
                    ViewBag.Error = "Please choose a start date prior the end date<br>";
                    return View("Index");
                }
                ViewBag.StartDate = valu.StartDate;
                ViewBag.EndDate = valu.EndDate;
                dateBool = true;
            }
            else if (valu.StartDate != null && valu.EndDate == null)
            {
                ViewBag.StartDate = valu.StartDate;
                dateBool = true;
            }
            else if (valu.EndDate != null && valu.StartDate == null)
            {
                ViewBag.EndDate = valu.EndDate;
                dateBool = true;
            }

            if (excelfile ==null || excelfile.ContentLength == 0)
            {
                ViewBag.Error = "Please select an excel file<br>";
                return View("Index");
            }

            else
            {
                if (excelfile.FileName.EndsWith("xls") || excelfile.FileName.EndsWith("xlsx"))
                {
                    string path = Server.MapPath("~/Content/" + excelfile.FileName);


                    //Save to Content folder
                    if (System.IO.File.Exists(path))
                    {
                        System.IO.File.Delete(path);
                    }
                    excelfile.SaveAs(path);

                    //Reading the data in the excel file
                    Application application = new Application();
                    Workbook workbook= application.Workbooks.Open(path);
                    Worksheet worksheet = workbook.ActiveSheet;
                    Range range = worksheet.UsedRange;
                    List<ImportExcel> listData = new List<ImportExcel>();

                    //Function I created as this function was getting overcrowded
                    readData(range, listData, valu, dateBool);

                    //this function closes the excel file and stops excel processes
                    closeExcel(application, workbook);

                    return View("Success");
                }
                else
                {
                    ViewBag.Error = "File type is incorrect<br>";
                    return View("Index");
                }
            }
        }

        public void readData(Range range, List<ImportExcel> listData, GetVal valu, bool dateBool)
        {
            int createdCnt = 0;
            int resolvedCnt = 0;
            int slaCnt = 0;
            int rejCnt = 0;
            bool resolvedBool = false;


            if (dateBool == true)
            {
                for (int row = 5; row < range.Rows.Count; row++)
                {
                    ImportExcel p = new ImportExcel();
                    string status = ((Range)range.Cells[row, 5]).Text;


                    //The four lines below allow for the application to convert the excel file value to fit DateTime's values
                    //Must need to do the same thing to find the date when the value was resolved
                    string created = ((Range)range.Cells[row, 11]).Value2.ToString();
                    double createdDate = double.Parse(created);
                    System.DateTime createdDateTime = System.DateTime.FromOADate(createdDate);


                    if (((Range)range.Cells[row, 14]).Value2 != null)
                    {
                        //we first need to check if status is resolved first before we can find resolved date
                        //this is because resolved date can be a null value and then function ToString() will not work
                        string findDate = ((Range)range.Cells[row, 14]).Value2.ToString();
                        double resolvedDate = double.Parse(findDate);
                        System.DateTime resolvedDateTime = System.DateTime.FromOADate(resolvedDate);
                        resolvedBool = true;

                        if ((valu.StartDate != null) && (valu.EndDate == null))
                        {
                            if (resolvedDateTime.CompareTo(valu.StartDate) >= 0)
                            {
                                resolvedTickets(p, resolvedDateTime, range, row);
                                resolvedCnt += 1;
                            }
                        }
                        else if ((valu.StartDate == null) && (valu.EndDate != null))
                        {
                            if (resolvedDateTime.CompareTo(valu.EndDate) <= 0)
                            {
                                resolvedTickets(p, resolvedDateTime, range, row);
                                resolvedCnt += 1;
                            }
                        }
                        else
                        {
                            if ((resolvedDateTime.CompareTo(valu.StartDate) >= 0) && (resolvedDateTime.CompareTo(valu.EndDate) <= 0))
                            {
                                resolvedTickets(p, resolvedDateTime, range, row);
                                resolvedCnt += 1;
                            }
                        }
                    }

                    if (resolvedBool == false) {

                        if ((valu.StartDate != null) && (valu.EndDate == null))
                        {
                            if (createdDateTime.CompareTo(valu.StartDate) >= 0)
                            {
                                createdTickets(p, createdDateTime, range, row);

                                if (status == "Rejected")
                                {
                                    rejCnt += 1;
                                }

                                if ((status == "To Do") || (status == "In Progress") || (status == "Open"))
                                {
                                    if (createdDateTime.AddDays(14) <= System.DateTime.Now)
                                    {
                                        slaCnt += 1;
                                    }
                                    else
                                    {
                                        createdCnt += 1;
                                    }
                                }
                            }
                        }

                        else if ((valu.StartDate == null) && (valu.EndDate != null))
                        {
                            if (createdDateTime.CompareTo(valu.EndDate) <= 0)
                            {
                                createdTickets(p, createdDateTime, range, row);

                                if (status == "Rejected")
                                {
                                    rejCnt += 1;
                                }

                                if ((status == "To Do") || (status == "In Progress") || (status == "Open"))
                                {
                                    if (createdDateTime.AddDays(14) <= System.DateTime.Now)
                                    {
                                        slaCnt += 1;
                                    }
                                    else
                                    {
                                        createdCnt += 1;
                                    }
                                }
                            }
                        }

                        else
                        {
                            if ((createdDateTime.CompareTo(valu.StartDate) >= 0) && (createdDateTime.CompareTo(valu.EndDate) <= 0))
                            {
                                createdTickets(p, createdDateTime, range, row);

                                if (status == "Rejected")
                                {
                                    rejCnt += 1;
                                }

                                if ((status == "To Do") || (status == "In Progress") || (status == "Open"))
                                {
                                    if (createdDateTime.AddDays(14) <= System.DateTime.Now)
                                    {
                                        slaCnt += 1;
                                    }
                                    else
                                    {
                                        createdCnt += 1;
                                    }
                                }
                            }
                        }

                        resolvedBool = false;

                    }

                    listData.Add(p);
                }
            }
            else
            {
                for (int row = 5; row < range.Rows.Count; row++)
                {
                    ImportExcel p = new ImportExcel();
                    string status = ((Range)range.Cells[row, 5]).Text;

                    /* p.Key = ((Range)range.Cells[row, 2]).Text;
                    p.KeyLink = ((Range)range.Cells[row, 2]).Hyperlinks[1].Address;
                    p.Status = ((Range)range.Cells[row, 5]).Text;
                    p.Summary = ((Range)range.Cells[row, 3]).Text;*/

                    //The four lines below allow for the application to convert the excel file value to fit DateTime's values
                    //Must need to do the same thing to find the date when the value was resolved
                    string created = ((Range)range.Cells[row, 11]).Value2.ToString();
                    double createdDate = double.Parse(created);
                    System.DateTime createdDateTime = System.DateTime.FromOADate(createdDate);
                    createdTickets(p, createdDateTime, range, row);
                    //p.Created = createdDateTime;

                    if (p.Status == "Resolved")
                    {
                        //we first need to check if status is resolved first before we can find resolved date
                        //this is because resolved date can be a null value and then function ToString() will not work
                        string findDate = ((Range)range.Cells[row, 14]).Value2.ToString();
                        double resolvedDate = double.Parse(findDate);
                        System.DateTime resolvedDateTime = System.DateTime.FromOADate(resolvedDate);
                        p.Resolved = resolvedDateTime;

                        resolvedCnt += 1;
                    }

                    if (p.Status == "Rejected")
                    {
                        rejCnt += 1;
                    }

                    if ((p.Status == "To Do") || (p.Status == "In Progress") || (p.Status == "Open"))
                    {
                        if (createdDateTime.AddDays(14) <= System.DateTime.Now)
                        {
                            slaCnt += 1;
                        }
                        else
                        {
                            createdCnt += 1;
                        }
                    }

                    

                    listData.Add(p);
                }
            }


            //brings values to Views file (specifically Success.cshtml)
            ViewBag.CreatedCount = createdCnt;
            ViewBag.ResolvedCount = resolvedCnt;
            ViewBag.SLACount = slaCnt;
            ViewBag.RejectedCount = rejCnt;
            ViewBag.ListData = listData;
        }


        public void createdTickets(ImportExcel p, System.DateTime createdDateTime, Range range, int row)
        {
            p.Created = createdDateTime;
            p.Key = ((Range)range.Cells[row, 2]).Text;
            p.KeyLink = ((Range)range.Cells[row, 2]).Hyperlinks[1].Address;
            p.Status = ((Range)range.Cells[row, 5]).Text;
            p.Summary = ((Range)range.Cells[row, 3]).Text;
        }


        public void resolvedTickets(ImportExcel p, System.DateTime resolvedDateTime, Range range, int row)
        {
            p.Resolved = resolvedDateTime;
            p.Key = ((Range)range.Cells[row, 2]).Text;
            p.KeyLink = ((Range)range.Cells[row, 2]).Hyperlinks[1].Address;
            p.Status = ((Range)range.Cells[row, 5]).Text;
            p.Summary = ((Range)range.Cells[row, 3]).Text;
        }


        public void closeExcel(Application application, Workbook workbook)
        {
            //closing excel file
            workbook.Save();
            workbook.Close(true);
            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);

            //This code kills any Microsoft Excel background processes
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
        }
    }
}
