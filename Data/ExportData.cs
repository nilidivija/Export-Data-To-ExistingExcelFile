using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Xml;
using System.IO;
using System.Text;
using System.Diagnostics;
using ExportTryOut.Data;

namespace ExportTryOut.Data
{
    public class ExportData
    {

        public void ExportExcelDoc(string fileName, string sheetname)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, true))
            {

                WorkbookPart workbookPart = document.WorkbookPart;
                IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetname);
               
                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheets.First().Id);
                worksheetPart.Worksheet = new Worksheet(new SheetData());


                List<Employee> employees = Employees.EmployeesList;

                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Constructing header
                Row row = new Row();

                row.Append(
                    ConstructCell("Id", CellValues.String),
                    ConstructCell("Name", CellValues.String),
                    ConstructCell("Birth Date", CellValues.String),
                    ConstructCell("Salary", CellValues.String)
                    );

                // Insert the header row to the Sheet Data
                sheetData.AppendChild(row);

                // Inserting each employee
                foreach (var emp in employees)
                {
                    row = new Row();

                    row.Append(
                        ConstructCell(emp.Id.ToString(), CellValues.Number),
                        ConstructCell(emp.Name, CellValues.String),
                        ConstructCell(emp.DOB.ToString("yyyy/MM/dd"), CellValues.String),
                        ConstructCell(emp.Salary.ToString(), CellValues.Number)
                        );

                    sheetData.AppendChild(row);
                }

                worksheetPart.Worksheet.Save();
            }


        }
      

        static Cell ConstructCell(string value, CellValues dataType)
        {
            return new Cell()
            {
                CellValue = new CellValue(value),
                DataType = new EnumValue<CellValues>(dataType)
            };
        }




    }
}
