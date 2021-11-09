using System;
using System.Collections.Generic;
// using DocumentFormat.OpenXml;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelBDD
{
    public static class Behavior
    {
        public static bool StartsWithUpper(this string str)
        {
            if (string.IsNullOrWhiteSpace(str))
                return false;

            char ch = str[0];
            return char.IsUpper(ch);
        }

        public static List<Dictionary<string, string>> GetExampleList(String filePath, String sheetName)
        {
            // try
            // {
            //     //specify the file name where its actually exist   
            //     string filepath = "D:\\TPMS\\Uploaded_Boq\\test.xlsx";

            //     //open the excel using openxml sdk  
            //     using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filepath, false))
            //     {

            //         //create the object for workbook part  
            //         WorkbookPart wbPart = doc.WorkbookPart;

            //         //statement to get the count of the worksheet  
            //         int worksheetcount = doc.WorkbookPart.Workbook.Sheets.Count();

            //         //statement to get the sheet object  
            //         Sheet mysheet = (Sheet)doc.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);

            //         //statement to get the worksheet object by using the sheet id  
            //         Worksheet Worksheet = ((WorksheetPart)wbPart.GetPartById(mysheet.Id)).Worksheet;

            //         //Note: worksheet has 8 children and the first child[1] = sheetviewdimension,....child[4]=sheetdata  
            //         int wkschildno = 4;


            //         //statement to get the sheetdata which contains the rows and cell in table  
            //         SheetData Rows = (SheetData)Worksheet.ChildElements.GetItem(wkschildno);


            //         //getting the row as per the specified index of getitem method  
            //         Row currentrow = (Row)Rows.ChildElements.GetItem(1);

            //         //getting the cell as per the specified index of getitem method  
            //         Cell currentcell = (Cell)currentrow.ChildElements.GetItem(1);

            //         //statement to take the integer value  
            //         string currentcellvalue = currentcell.InnerText;

            //     }
            // }
            // catch (Exception Ex)
            // {

            // }
            List<Dictionary<string, string>> exampleList = new List<Dictionary<string, string>>();
            return exampleList;
        }
    }
}
