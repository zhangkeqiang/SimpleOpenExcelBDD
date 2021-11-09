using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
// using UtilityLibraries;
using ExcelBDD;
namespace ExcelBDDTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestGetDataTable()
        {
            String currentPath = Directory.GetCurrentDirectory();
            System.Console.WriteLine(currentPath);
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            System.Console.WriteLine(filePath);
            List<Dictionary<string, string>> exampleList = ExcelBDD.Behavior.GetDataTable(filePath, "DataTable1", 2);
            Assert.IsNotNull(exampleList);
            Assert.AreEqual(7, exampleList.Count);
        }
    }
}

