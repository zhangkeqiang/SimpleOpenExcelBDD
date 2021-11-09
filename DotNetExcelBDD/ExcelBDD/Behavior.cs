using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;

namespace ExcelBDD
{
    public static class Behavior
    {
        public static IEnumerable<object[]> GetDataTableByArray(String filePath, String sheetName, int headerRow)
        {
            List<object[]> exampleList = new List<object[]>();
            List<string> Headers = new List<string>();
            //open the excel using openxml sdk  
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                String sheetIdValue = null;
                foreach (Sheet eachsheet in sheets)
                {
                    Console.WriteLine(eachsheet.Name);
                    if (eachsheet.Name == sheetName)
                    {
                        sheetIdValue = eachsheet.Id.Value;
                        break;
                    }
                }

                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheetIdValue) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                int counter = 0;
                int columnCount = 0;
                foreach (Row row in rows)
                {
                    counter = counter + 1;
                    //Read the headerRow row as header
                    if (counter == headerRow)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = GetCellValue(doc, cell);
                            columnCount++;
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                            // dt.Columns.Add(colunmName);
                        }
                    }
                    else if (counter > headerRow)
                    {
                        // dt.Rows.Add();
                        // Dictionary<string, string> dic = new System.Collections.Generic.Dictionary<string, string>();
                        object[] values = new object[columnCount];
                        int i = 0;
                        Console.Write(counter);
                        Console.Write(": ");
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            String cellValue = GetCellValue(doc, cell);
                            Console.Write(cellValue + " | ");
                            // dt.Rows[dt.Rows.Count - 1][i] = cellValue;
                            values[i] = (object)cellValue;
                            i++;
                        }
                        Console.WriteLine();
                        exampleList.Add(values);
                    }
                }

            }
            return (exampleList as IEnumerable<object[]>);
        }

        public static IEnumerable<object[]> GetDataTable(String filePath, String sheetName, int headerRow)
        {
            List<object[]> exampleList = new List<object[]>();
            List<string> Headers = new List<string>();
            // DataTable dt = new DataTable();
            //open the excel using openxml sdk  
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                Sheets sheets = doc.WorkbookPart.Workbook.Sheets;
                String sheetIdValue = null;
                foreach (Sheet eachsheet in sheets)
                {
                    Console.WriteLine(eachsheet.Name);
                    if (eachsheet.Name == sheetName)
                    {
                        sheetIdValue = eachsheet.Id.Value;
                        break;
                    }
                }

                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheetIdValue) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                int counter = 0;
                int columnCount = 0;
                foreach (Row row in rows)
                {
                    counter = counter + 1;
                    //Read the headerRow row as header
                    if (counter == headerRow)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = GetCellValue(doc, cell);
                            columnCount++;
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                        }
                    }
                    else if (counter > headerRow)
                    {
                        Dictionary<string, string> dic = new System.Collections.Generic.Dictionary<string, string>();
                        int i = 0;
                        Console.Write(counter);
                        Console.Write(": ");
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            String cellValue = GetCellValue(doc, cell);
                            Console.Write(cellValue + " | ");
                            dic.Add(Headers[i], cellValue);
                            i++;
                        }
                        Console.WriteLine();
                        exampleList.Add(new object[] { dic });
                    }
                }
            }

            return (exampleList as IEnumerable<object[]>);
        }
        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            try
            {
                string value = cell.CellValue.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                }
                return value;
            }
            catch (NullReferenceException)
            {
                // Console.WriteLine(cell.ToString() + " is null");
                return "";
            }
        }

        public static String GetValue(Dictionary<string, string> dic, string parameterName)
        {
            try
            {
                return dic[parameterName];
            }
            catch (KeyNotFoundException e)
            {
                return "";
            }
        }
    }
}
