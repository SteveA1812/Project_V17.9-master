using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using Project_V17.Data;
using Project_V17.Models;
using System.Web;
namespace Project_V17.Pages.Admin
{
    public class ExportModel : PageModel
    {
        private readonly Project_V17.Data.ApplicationDbContext _context;

        public ExportModel(Project_V17.Data.ApplicationDbContext context)
        {
            _context = context;
        }

        public IList<FSApp> FSApp { get; set; }
        public IEnumerable<object> Applications { get; private set; }

        public async Task OnGetAsync()
        {
            FSApp = await _context.FSApp.ToListAsync();
        }
        [HttpGet]
        public IActionResult Export()
        {

            List<FSApp> Applications = _context.FSApp.Select(x => new FSApp
            {
                StaffFirstName = x.StaffFirstName,
                StaffSurname = x.StaffSurname,
                Department = x.Department,
                Function = x.Function,
                CourseName = x.CourseName,
                Level = x.Level,
                Provider = x.Provider,
                Details = x.Details,
                StartYear = x.StartYear,
                Mode = x.Mode,
                Duration = x.Duration,
                Cost = x.Cost,
                Q1 = x.Q1,
                Q2 = x.Q2,
                Q3 = x.Q3,
                Q4 = x.Q4,
                Q5 = x.Q5

            }).ToList();


            var stream = new MemoryStream();

            using (var package = new ExcelPackage(stream))
            {
                var workSheet = package.Workbook.Worksheets.Add("Sheet1");
                workSheet.TabColor = System.Drawing.Color.Black;
                workSheet.DefaultRowHeight = 12;
                workSheet.Cells.LoadFromCollection(Applications, true);
                package.Save();
            }
            stream.Position = 0;
            string excelName = "FurtherStudyApplications.xlsx";
            return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

        public void ExportToExcel()
        {
            List<FSApp> Applications = _context.FSApp.Select(x => new FSApp
            {

                StaffFirstName = x.StaffFirstName,
                StaffSurname = x.StaffSurname,
                Department = x.Department,
                Function = x.Function,
                CourseName = x.CourseName,
                Level = x.Level,
                Provider = x.Provider,
                Details = x.Details,
                StartYear = x.StartYear,
                Mode = x.Mode,
                Duration = x.Duration,
                Cost = x.Cost,
                Q1 = x.Q1,
                Q2 = x.Q2,
                Q3 = x.Q3,
                Q4 = x.Q4,
                Q5 = x.Q5

            }).ToList();

            ExcelPackage pck = new ExcelPackage();
            ExcelWorksheet ws = pck.Workbook.Worksheets.Add("Report");

            ws.Cells["A1"].Value = "Further Study";
            ws.Cells["B1"].Value = "Applications";

            ws.Cells["A2"].Value = "Report";
            ws.Cells["B2"].Value = "Report1";

            ws.Cells["A3"].Value = "Date";
            ws.Cells["B3"].Value = string.Format("{0:dd MMMM yyyy} at {0:H: mm tt}", DateTimeOffset.Now);

            ws.Cells["A6"].Value = "StaffFirstName";
            ws.Cells["B6"].Value = "StaffSurname";
            ws.Cells["C6"].Value = "Department";
            ws.Cells["D6"].Value = "Function";
            ws.Cells["E6"].Value = "CourseName";
            ws.Cells["F6"].Value = "Level";
            ws.Cells["G6"].Value = "Provider";
            ws.Cells["H6"].Value = "Details";
            ws.Cells["I6"].Value = "StartYear";
            ws.Cells["J6"].Value = "Mode";
            ws.Cells["K6"].Value = "Duration";
            ws.Cells["L6"].Value = "Cost";
            ws.Cells["M6"].Value = "Q1";
            ws.Cells["N6"].Value = "Q2";
            ws.Cells["O6"].Value = "Q3";
            ws.Cells["P6"].Value = "Q4";
            ws.Cells["Q6"].Value = "Q5";

            int rowStart = 7;
            foreach (var item in Applications)
            {
                if (item.Cost < 3500)
                {
                    ws.Row(rowStart).Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    ws.Row(rowStart).Style.Fill.BackgroundColor.SetColor(ColorTranslator.FromHtml(string.Format("pink")));

                }

                ws.Cells[string.Format("A{0}", rowStart)].Value = item.StaffFirstName;
                ws.Cells[string.Format("B{0}", rowStart)].Value = item.StaffSurname;
                ws.Cells[string.Format("C{0}", rowStart)].Value = item.Department;
                ws.Cells[string.Format("D{0}", rowStart)].Value = item.Function;
                ws.Cells[string.Format("E{0}", rowStart)].Value = item.CourseName;
                ws.Cells[string.Format("F{0}", rowStart)].Value = item.Level;
                ws.Cells[string.Format("G{0}", rowStart)].Value = item.Provider;
                ws.Cells[string.Format("H{0}", rowStart)].Value = item.Details;
                ws.Cells[string.Format("I{0}", rowStart)].Value = item.StartYear;
                ws.Cells[string.Format("J{0}", rowStart)].Value = item.Mode;
                ws.Cells[string.Format("K{0}", rowStart)].Value = item.Duration;
                ws.Cells[string.Format("L{0}", rowStart)].Value = item.Cost;
                ws.Cells[string.Format("M{0}", rowStart)].Value = item.Q1;
                ws.Cells[string.Format("N{0}", rowStart)].Value = item.Q2;
                ws.Cells[string.Format("O{0}", rowStart)].Value = item.Q3;
                ws.Cells[string.Format("P{0}", rowStart)].Value = item.Q4;
                ws.Cells[string.Format("Q{0}", rowStart)].Value = item.Q5;
                rowStart++;
            }

            ws.Cells["A:AZ"].AutoFitColumns();
            Response.Clear();
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            Response.Headers.Add("content-disposition", "attachment: filename=" + "ExcelReport.xlsx");
            Response.Body.WriteAsync(pck.GetAsByteArray());


        }
    }
}