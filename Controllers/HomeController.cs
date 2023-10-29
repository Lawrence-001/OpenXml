using Microsoft.AspNetCore.Mvc;
using openxml.Models;
using System.Data;
using System.Diagnostics;
using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace openxml.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly AppDbContext _dbContext;

        public HomeController(ILogger<HomeController> logger, AppDbContext dbContext)
        {
            _logger = logger;
            _dbContext = dbContext;
        }

        public IActionResult Index()
        {
            return View();
        }

        //public IActionResult ReadData()
        //{
        //    string filepath = @"C:\Users\lmutuku\Desktop\Learn\Openxml.xlsx";

        //    // Open the document as read-only.
        //    using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
        //    {
        //        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        //        WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
        //        SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
        //        string text;
        //        foreach (Row r in sheetData.Elements<Row>())
        //        {
        //            foreach (Cell c in r.Elements<Cell>())
        //            {
        //                text = c.CellValue.Text;

        //                _dbContext.SaveChanges();
        //                //Console.Write(text + " ");
        //            }
        //        }
        //    }

        //    return View();
        //}

        //public IActionResult ReadData()
        //{
        //    string filepath = @"C:\Users\lmutuku\Desktop\Learn\Openxml.xlsx";

        //    // Open the document as read-only.
        //    using (var spreadsheetDocument = SpreadsheetDocument.Open(filepath, false))
        //    {
        //        var workbookPart = spreadsheetDocument.WorkbookPart;
        //        var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
        //        var worksheet = worksheetPart.Worksheet;

        //        var sharedStringTablePart = workbookPart.SharedStringTablePart;
        //        var sharedStringTable = sharedStringTablePart.SharedStringTable;

        //        var rows = worksheet.GetFirstChild<SheetData>().Elements<Row>();

        //        foreach (var row in rows.Skip(1)) // Skip the header row
        //        {
        //            var cellValues = row.Elements<Cell>().Select(c => c.InnerText).ToList();

        //            if (cellValues.Count >= 4)
        //            {
        //                Student std = new Student
        //                {
        //                    //Id = int.Parse(cellValues[0]),
        //                    FirstName = cellValues[1],
        //                    LastName = cellValues[2],
        //                    Age = int.Parse(cellValues[3]),
        //                    City = cellValues[4]

        //                };
        //                _dbContext.Students.Add(std);
        //                _dbContext.SaveChanges();
        //            }
        //        }
        //    }
        //    return View();
        //}

        //public IActionResult ReadData(Student student)
        //{
        //    List<Student> students = new List<Student>();
        //    var data = GetData();
        //    student.FirstName = data.Columns[0].ColumnName;
        //    student.LastName = data.Columns[1].ColumnName; //+ " " + data.Columns[1].ColumnName;

        //    //_dbContext.Students.AddRange(data);
        //    _dbContext.SaveChanges();
        //    return View();
        //}
        public IActionResult ReadData()
        {
            var data = GetData(); // Assuming GetData() returns a DataTable or some data source

            foreach (DataRow row in data.Rows)
            {
                var student = new Student
                {
                    FirstName = row[data.Columns[0].ColumnName].ToString(),
                    LastName = row[data.Columns[1].ColumnName].ToString(),
                    Age = Convert.ToInt32(row[data.Columns[2].ColumnName]),
                City = row[data.Columns[3].ColumnName].ToString()
                };

                _dbContext.Students.Add(student);
            }

            _dbContext.SaveChanges();
            return View();
        }


        public static DataTable GetData()
        {
            string filePath = @"C:\Users\lmutuku\Desktop\Learn\Openxml.xlsx";
            var table = new DataTable();
            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = spreadSheetDocument.WorkbookPart;
                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    table.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                //this will also include your header row...
                foreach (Row row in rows)
                {
                    DataRow tempRow = table.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                    }

                    table.Rows.Add(tempRow);
                }
            }

            table.Rows.RemoveAt(0);
            
            return table;
        }


        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}