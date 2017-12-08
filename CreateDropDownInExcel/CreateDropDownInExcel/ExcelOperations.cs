using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace CreateDropDownInExcel
{
    class ExcelOperations
    {
        public static void CreatingExcelAndDrowownInExcel()
        {
            var filepath = @"C:\Test.xlsx";
            OpenXMLWindowsApp app = new OpenXMLWindowsApp();
            //app.UpdateSheet(filepath);
            SpreadsheetDocument myWorkbook = SpreadsheetDocument.Create(filepath, SpreadsheetDocumentType.Workbook);
            //SpreadsheetDocument myWorkbook = SpreadsheetDocument.Open(filepath,true);
            WorkbookPart workbookpart = myWorkbook.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            WorksheetPart worksheetPart2 = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart2.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = myWorkbook.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
            Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            SheetDimension sheetDimension1 = new SheetDimension() { Reference = "A1" };
            Sheet sheet = new Sheet()
            {
                Id = myWorkbook.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "DropDownContainingSheet"
            };

            Sheet sheet1 = new Sheet()
            {
                Id = myWorkbook.WorkbookPart.GetIdOfPart(worksheetPart2),
                SheetId = 2,
                Name = "DropDownDataContainingSheet"

            };

            sheets.Append(sheet);
            sheets.Append(sheet1);
            SheetData sheetData = new SheetData();
            SheetData sheetData1 = new SheetData();
            int Counter1 = 1;
            foreach (var value in DataInSheet.GetDataOfSheet1())
            {

                Row contentRow = CreateRowValues(Counter1, value);
                Counter1++;
                sheetData.AppendChild(contentRow);
            }

            worksheet1.Append(sheetData);
            int Counter2 = 1;
            foreach (var value in DataInSheet.GetDataOfSheet2())
            {

                Row contentRow = CreateRowValues(Counter2, value);
                Counter2++;
                sheetData1.AppendChild(contentRow);
            }
            worksheet2.Append(sheetData1);


            DataValidation dataValidation = new DataValidation
            {
                Type = DataValidationValues.List,
                AllowBlank = true,
                SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" },
                Formula1 = new Formula1("'DropDownDataContainingSheet'!$A$1:$A$8")

            };

            DataValidations dataValidations = worksheet1.GetFirstChild<DataValidations>();
            if (dataValidations != null)
            {
                dataValidations.Count = dataValidations.Count + 1;
                dataValidations.Append(dataValidation);
            }
            else
            {
                DataValidations newdataValidations = new DataValidations();
                newdataValidations.Append(dataValidation);
                newdataValidations.Count = 1;
                worksheet1.Append(newdataValidations);
            }


            worksheetPart.Worksheet = worksheet1; ;
            worksheetPart2.Worksheet = worksheet2;
            workbookpart.Workbook.Save();
            myWorkbook.Close();

        }
        static string[] headerColumns = new string[] { "A", "B", "C", "D" };
        public static Row CreateRowValues(int index, DataInSheet objToInsert)
        {
            Row row = new Row();
            row.RowIndex = (UInt32)index;
            int i = 0;
            foreach (var property in objToInsert.GetType().GetProperties())
            {
                Cell cell = new Cell();
                cell.CellReference = headerColumns[i].ToString() + index;
                if (property.PropertyType.ToString().Equals("System.string", StringComparison.InvariantCultureIgnoreCase))
                {

                    var result = property.GetValue(objToInsert, null);
                    if (result == null)
                    {
                        result = "";
                    }
                    cell.DataType = CellValues.String;
                    InlineString inlineString = new InlineString();
                    Text text = new Text();
                    text.Text = result.ToString();
                    inlineString.AppendChild(text);
                    cell.AppendChild(inlineString);
                }
                if (property.PropertyType.ToString().Equals("System.int32", StringComparison.InvariantCultureIgnoreCase))
                {
                    var result = property.GetValue(objToInsert, null);
                    if (result == null)
                    {
                        result = 0;
                    }
                    CellValue cellValue = new CellValue();
                    cellValue.Text = result.ToString();
                    cell.AppendChild(cellValue);
                }
                if (property.PropertyType.ToString().Equals("System.boolean", StringComparison.InvariantCultureIgnoreCase))
                {
                    var result = property.GetValue(objToInsert, null);
                    if (result == null)
                    {
                        result = "False";
                    }
                    cell.DataType = CellValues.InlineString;
                    InlineString inlineString = new InlineString();
                    Text text = new Text();
                    text.Text = result.ToString();
                    inlineString.AppendChild(text);
                    cell.AppendChild(inlineString);
                }

                row.AppendChild(cell);
                i = i + 1;
            }
            return row;
        }
    }


    public class OpenXMLWindowsApp
    {
        public void UpdateSheet(string filepath)
        {
            UpdateCell(filepath, "20", 2, "B");
            //UpdateCell("Chart.xlsx", "80", 3, "B");
            //UpdateCell("Chart.xlsx", "80", 2, "C");
            //UpdateCell("Chart.xlsx", "20", 3, "C");

            //ProcessStartInfo startInfo = new ProcessStartInfo("Chart.xlsx");
            //startInfo.WindowStyle = ProcessWindowStyle.Normal;
            //Process.Start(startInfo);
        }

        public static void UpdateCell(string docName, string text,
            uint rowIndex, string columnName)
        {


            //Worksheet worksheet1 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            //worksheet1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            //worksheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            //worksheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            //Worksheet worksheet2 = new Worksheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            //worksheet2.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            //worksheet2.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            //worksheet2.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

             Worksheet worksheet1 = new Worksheet();
              
            ExcelOperations ex = new ExcelOperations();
            // Open the document for editing.
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(docName, true))
            {
                WorksheetPart worksheetPart = GetWorksheetPartByName(spreadSheet, "DropDownContainingSheet");

                if (worksheetPart != null)
                {
                    worksheet1 = worksheetPart.Worksheet;
                    SheetData sheetData = new SheetData();
                    SheetData sheetData1 = new SheetData();
                    int Counter1 = 1;

                    foreach (var value in DataInSheet.GetDataOfSheet1())
                    {

                        Row contentRow = ExcelOperations.CreateRowValues(Counter1, value);
                        Counter1++;
                        sheetData.AppendChild(contentRow);
                    }

                    worksheet1.Append(sheetData);
                    int Counter2 = 1;
                    //foreach (var value in DataInSheet.GetDataOfSheet2())
                    //{

                    //    Row contentRow = ExcelOprations.CreateRowValues(Counter2, value);
                    //    Counter2++;
                    //    sheetData1.AppendChild(contentRow);
                    //}
                    //worksheet2.Append(sheetData1);

                    
                    DataValidation dataValidation = new DataValidation
                    {
                        Type = DataValidationValues.List,
                        AllowBlank = true,
                        SequenceOfReferences = new ListValue<StringValue>() { InnerText = "A1" },
                        Formula1 = new Formula1("'DropDownDataContainingSheet'!$B$1:$B$5")

                    };

                    DataValidations dataValidations = worksheet1.GetFirstChild<DataValidations>();
                    if (dataValidations != null)
                    {
                        dataValidations.Count = dataValidations.Count + 1;
                        dataValidations.Append(dataValidation);
                    }
                    else
                    {
                        DataValidations newdataValidations = new DataValidations();
                        newdataValidations.Append(dataValidation);
                        newdataValidations.Count = 1;
                        worksheet1.Append(newdataValidations);
                    }
                    


                    Cell cell = GetCell(worksheetPart.Worksheet,
                                             columnName, rowIndex);

                    cell.CellValue = new CellValue(text);
                    cell.DataType =
                        new EnumValue<CellValues>(CellValues.Number);

                    // Save the worksheet.
                    worksheetPart.Worksheet.Append(worksheet1);
                   // worksheetPart.Worksheet=(worksheet2);

                    worksheetPart.Worksheet.Save();



                }
            }

        }

        private static WorksheetPart   GetWorksheetPartByName(SpreadsheetDocument document,string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private static Cell GetCell(Worksheet worksheet,
                  string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();
        }


        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }
    }

}
