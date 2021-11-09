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
        public static List<Dictionary<string, string>> GetDataTable(String filePath, String sheetName, int headerRow)
        {
            List<Dictionary<string, string>> exampleList = new List<Dictionary<string, string>>();
            List<string> Headers = new List<string>();
            DataTable dt = new DataTable();
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

                    // foreach (OpenXmlAttribute attr in asheet.GetAttributes())
                    // {
                    //     Console.WriteLine("{0}: {1}", attr.LocalName, attr.Value);
                    // }
                }

                Worksheet worksheet = (doc.WorkbookPart.GetPartById(sheetIdValue) as WorksheetPart).Worksheet;
                IEnumerable<Row> rows = worksheet.GetFirstChild<SheetData>().Descendants<Row>();
                int counter = 0;
                foreach (Row row in rows)
                {
                    counter = counter + 1;
                    //Read the headerRow row as header
                    if (counter == headerRow)
                    {
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            var colunmName = GetCellValue(doc, cell);
                            Console.WriteLine(colunmName);
                            Headers.Add(colunmName);
                            dt.Columns.Add(colunmName);
                        }
                    }
                    else if (counter > headerRow)
                    {
                        dt.Rows.Add();
                        Dictionary<string, string> dic = new System.Collections.Generic.Dictionary<string, string>();
                        int i = 0;
                        Console.Write(counter);
                        Console.Write(": ");
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            String cellValue = GetCellValue(doc, cell);
                            Console.Write(cellValue + " | ");
                            dt.Rows[dt.Rows.Count - 1][i] = cellValue;
                            dic.Add(Headers[i], cellValue);
                            i++;
                        }
                        Console.WriteLine();
                        exampleList.Add(dic);
                    }
                }

            }

            return exampleList;
        }

        private static string GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            string value = "";
            try
            {
                value = cell.CellValue.InnerText;
                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                {
                    return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                }
            }
            catch (NullReferenceException)
            {
                // Console.WriteLine(cell.ToString() + " is null");
            }
            return value;
        }
    }
}
