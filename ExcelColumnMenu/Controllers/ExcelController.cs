using ExcelColumnMenu.Controllers;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelCoumnMenus.Controllers
{
    [Route("/[controller]/[action]")]
    public class ExcelController : Controller
    {

        private IDataValidation GetColumnValidation(ISheet refSheet, CellRangeAddressList adress, string[] options)
        {
            IDataValidationHelper helper = new XSSFDataValidationHelper(refSheet as XSSFSheet);
            var constrint = helper.CreateExplicitListConstraint(options);
            IDataValidation dataValidation = helper.CreateValidation(constrint, adress);
            dataValidation.ShowErrorBox = true;
            dataValidation.CreateErrorBox("Select", "please select one of the option from list");
            dataValidation.SuppressDropDownArrow = true;
            return dataValidation;
        }
        private IDataValidation GetColumnFormulaListValidation(ISheet refSheet, CellRangeAddressList adress, string refAddress)
        {
            IDataValidationHelper helper = new XSSFDataValidationHelper(refSheet as XSSFSheet);
            var constrint = helper.CreateFormulaListConstraint(refAddress);
            IDataValidation dataValidation = helper.CreateValidation(constrint, adress);
            dataValidation.ShowErrorBox = true;
            dataValidation.CreateErrorBox("Select", "please select one of the option from list");
            dataValidation.SuppressDropDownArrow = true;
            return dataValidation;
        }
        private string AddHiddenReferenceSheet(IWorkbook workbook, string[] names)
        {

            string sheetName = "Courses";
            string columnName = "CourseName";
            ISheet excelSheet = workbook.CreateSheet(sheetName);
            workbook.SetSheetHidden(1, SheetState.VeryHidden);
            IRow row = excelSheet.CreateRow(0);
            row.CreateCell(0).SetCellValue(columnName);

            for (int i = 0; i < names.Length; i++)
            {
                IRow newrow = excelSheet.CreateRow(i + 1);
                newrow.CreateCell(0).SetCellValue(names[i]);
            }
            return sheetName;
        }
        [HttpPost]
        public async Task<IActionResult> ExportExcel([FromForm] ExcelColumnDropdownOptions menuOptions,[FromServices] IWebHostEnvironment hosting)
        {
          

            string webRootPath = hosting.WebRootPath;
            string fileName = Guid.NewGuid().ToString() + ".xlsx";
            string fullFileName = Path.Combine(webRootPath, fileName);
            var memoryStream = new MemoryStream();
           
            using (var fs = new FileStream(fullFileName, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Courses Details");

                IRow row = excelSheet.CreateRow(0);
                row.CreateCell(0).SetCellValue("S.No");
                row.CreateCell(1).SetCellValue("Course Name");
                row.CreateCell(2).SetCellValue("Institute Name");
                row.CreateCell(3).SetCellValue("Course Started");
                row.CreateCell(4).SetCellValue("Course Period");
                row.CreateCell(5).SetCellValue("Price");
              
                IRow row1 = excelSheet.CreateRow(1);
                row1.CreateCell(0).SetCellValue(1);
                row1.CreateCell(1).SetCellValue("NODE.JS");
                row1.CreateCell(2).SetCellValue("Naresh Technologies");
                row1.CreateCell(3).SetCellValue("Yes");
                row1.CreateCell(4).SetCellValue("2 Weeks");
                row1.CreateCell(5).SetCellValue("10000");

                string sheetName = AddHiddenReferenceSheet(workbook,menuOptions.Courses);

                excelSheet.AddValidationData(GetColumnFormulaListValidation(excelSheet, new CellRangeAddressList(1, 1000, 1, 1), $"={sheetName}!$A$1:$A$1000"));
                excelSheet.AddValidationData(GetColumnValidation(excelSheet, new CellRangeAddressList(1, 1000, 2, 2), menuOptions.CourseInstitution));
                excelSheet.AddValidationData(GetColumnValidation(excelSheet, new CellRangeAddressList(1, 1000, 3, 3), new string[] { "Yes", "No" }));
                excelSheet.AddValidationData(GetColumnValidation(excelSheet, new CellRangeAddressList(1, 1000, 4, 4), menuOptions.CoursePeriods));



                workbook.Write(fs);
            }
            using (var fileStream = new FileStream(fullFileName, FileMode.Open))
            {
                await fileStream.CopyToAsync(memoryStream);
            }
            memoryStream.Position = 0;
            return File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", menuOptions.SheetName);
        }
    }
}
