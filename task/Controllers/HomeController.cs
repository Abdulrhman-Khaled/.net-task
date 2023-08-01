using Microsoft.AspNetCore.Mvc;
using System.Data;
using Microsoft.AspNetCore.Http;
using System.IO;
using IHostingEnvironment = Microsoft.AspNetCore.Hosting.IHostingEnvironment;
using ClosedXML.Excel;

namespace task.Controllers;

public class HomeController : Controller
{
    private IHostingEnvironment Environment;
    public HomeController(IHostingEnvironment _environment)
    {
        Environment = _environment;
    }
    public IActionResult Index()
    {
        return View();
    }

    [HttpPost]
    public IActionResult Index(IFormFile postedFile)
    {
        if (postedFile != null)
        {
            string path = Path.Combine(this.Environment.WebRootPath, "Uploads");
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string fileName = Path.GetFileName(postedFile.FileName);
            string filePath = Path.Combine(path, fileName);
            using (FileStream stream = new(filePath, FileMode.Create))
            {
                postedFile.CopyTo(stream);
            }
            string csvData = System.IO.File.ReadAllText(filePath);
            DataTable dt = new();
            bool firstRow = true;
            foreach (string row in csvData.Split('\n'))
            {
                if (!string.IsNullOrEmpty(row))
                {
                    if (!string.IsNullOrEmpty(row))
                    {
                        if (firstRow)
                        {
                            foreach (string cell in row.Split(','))
                            {
                                dt.Columns.Add(cell.Trim());
                            }
                            firstRow = false;
                        }
                        else
                        {
                            dt.Rows.Add();
                            int i = 0;
                            foreach (string cell in row.Split(','))
                            {
                                dt.Rows[^1][i] = cell.Trim();
                                i++;
                            }
                        }
                    }
                }
            }

            return View(dt);
        }

        return View();
    }

    [HttpGet]
public IActionResult Download(string fileName)
{
    string path = Path.Combine(this.Environment.WebRootPath, "Uploads", fileName);

    using (XLWorkbook workbook = new())
    {
        var worksheet = workbook.Worksheets.Add("Sheet1");
        worksheet.Cell(1, 1).Value = "Data from CSV file";
        worksheet.Cell(1, 1).Style.Font.Bold = true;

        string csvData = System.IO.File.ReadAllText(path);
        bool firstRow = true;
        int rowNumber = 2;
        foreach (string row in csvData.Split('\n'))
        {
            if (!string.IsNullOrEmpty(row))
            {
                if (firstRow)
                {
                    foreach (string cell in row.Split(','))
                    {
                        worksheet.Cell(rowNumber, cell.Split(',')[0]).Value = cell.Split(',')[0];
                        worksheet.Cell(rowNumber, cell.Split(',')[0]).Style.Font.Bold = true;
                    }
                    firstRow = false;
                    rowNumber++;
                }
                else
                {
                    int columnNumber = 1;
                    foreach (string cell in row.Split(','))
                    {
                        worksheet.Cell(rowNumber, columnNumber).Value = cell.Trim();
                        columnNumber++;
                    }
                    rowNumber++;
                }
            }
        }

        var range = worksheet.Range(1, 1, rowNumber - 1, worksheet.ColumnsUsed().Count());
        range.Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
        range.Style.Border.InsideBorder = XLBorderStyleValues.Thin;

        using (MemoryStream stream = new())
        {
            workbook.SaveAs(stream);
            var content = stream.ToArray();
            return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "data.xlsx");
        }
    }
}
}
