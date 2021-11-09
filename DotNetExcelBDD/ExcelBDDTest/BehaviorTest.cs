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
            IEnumerable<object[]> exampleList = ExcelBDD.Behavior.GetDataTable(filePath, "DataTable1", 2);
            Assert.IsNotNull(exampleList);
            int count = 0;
            foreach (var item in exampleList)
            {
                count++;
            }
            Assert.AreEqual(7, count);
        }


        [DataTestMethod]
        [DynamicData(nameof(GetData), DynamicDataSourceType.Method)]
        public void TestDataTable1(String value1, String value2, String value3, String value4, String value5, String value6, String value7, String value8)
        {
            Console.WriteLine("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}", value1,  value2,  value3,  value4,  value5,  value6,  value7,  value8);
        }

        public static IEnumerable<object[]> GetData()
        {
            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            Console.WriteLine(filePath);
            return ExcelBDD.Behavior.GetDataTable(filePath, "DataTable1", 2);
        }
    }
}

