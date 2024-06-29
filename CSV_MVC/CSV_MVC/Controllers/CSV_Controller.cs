using CSV_MVC.Models;
using CsvHelper;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using System.Formats.Asn1;
using System.Globalization;

namespace CSV_MVC.Controllers
{
    public class CSV_Controller : Controller
    {
        public ActionResult Index()
        {
            return View(new DataViewModel());
        }


        [HttpPost("upload")]
        public ActionResult Upload(IFormFile file)
        {
            DataViewModel model = new();
            if (file == null || file.Length == 0)
            {
                ViewBag.ErrorMessage = "No file uploaded or its empty.";
                return View("Index",model);
            }

            var allowedExtension =  ".csv";
            var fileExtension = Path.GetExtension(file.FileName).ToLower();

            if (!allowedExtension.Equals(fileExtension))
            {
                ViewBag.ErrorMessage = "Invalid file extension. Only CSV files are allowed.";
                return View("Index", model);
            }

            model.Headers = new List<string>();
            model.Rows = new List<List<string>>();
            using (var reader = new StreamReader(file.OpenReadStream()))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                model.Headers.AddRange(reader.ReadLine().Split(','));
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');
                    model.Rows.Add(new List<string>(values));
                }
            }

            return View("Index", model);
        }

        public ActionResult DownloadAsXLSX(DataViewModel model)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");

                for (int i = 0; i < model.Headers.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = model.Headers[i];
                }

                for (int i = 0; i < model.Rows.Count; i++)
                {
                    for (int j = 0; j < model.Rows[i].Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = model.Rows[i][j];
                    }
                }

                var stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                string excelName = $"Data-{DateTime.Now:yyyyMMddHHmmssfff}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
            }
        }
    }
}

