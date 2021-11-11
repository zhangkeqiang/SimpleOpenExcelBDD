using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
// using UtilityLibraries;
using ExcelBDD;
namespace ExcelBDDTest
{
    [TestClass]
    public class DataTableTest
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
            int count = 0;
            foreach (var item in exampleList)
            {
                Dictionary<string, string> dic = item;
                Console.Write(dic.ToString());
                Console.WriteLine(dic["Header01"]);
                count++;
            }
            Assert.AreEqual(7, count);
        }

        [TestMethod]
        public void TestGetDataTableByArray()
        {
            String currentPath = Directory.GetCurrentDirectory();
            System.Console.WriteLine(currentPath);
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            System.Console.WriteLine(filePath);
            IEnumerable<object[]> exampleList = ExcelBDD.Behavior.GetDataTableByArray(filePath, "DataTable1", 2);
            Assert.IsNotNull(exampleList);
            int count = 0;
            foreach (var item in exampleList)
            {
                count++;
            }
            Assert.AreEqual(7, count);
        }


        [DataTestMethod]
        [DynamicData(nameof(GetDataByObjectArray), DynamicDataSourceType.Method)]
        public void TestDataTableByObjectArray(String value1, String value2, String value3, String value4, String value5, String value6, String value7, String value8)
        {
            Console.WriteLine("{0}|{1}|{2}|{3}|{4}|{5}|{6}|{7}", value1, value2, value3, value4, value5, value6, value7, value8);
        }

        public static IEnumerable<object[]> GetDataByObjectArray()
        {
            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            Console.WriteLine(filePath);
            return ExcelBDD.Behavior.GetDataTableByArray(filePath, "DataTable1", 2);
        }

        [DataTestMethod]
        [DynamicData(nameof(GetDataByDictionary), DynamicDataSourceType.Method)]
        public void TestDataTable(Dictionary<string, string> paramDic)
        {
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header01"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header02"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header03"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header04"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header05"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header06"));
            Console.Write("{0}|", Behavior.GetValue(paramDic, "Header07"));
            Console.WriteLine("{0}", Behavior.GetValue(paramDic, "Header08"));
            foreach (KeyValuePair<string, string> item in paramDic)
            {
                Console.WriteLine("Dictionary: {0} - {1}", item.Key, item.Value);
            }
        }

        public static IEnumerable<object[]> GetDataByDictionary()
        {
            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            Console.WriteLine(filePath);
            return ExcelBDD.Behavior.ConvertToIEnumerable(ExcelBDD.Behavior.GetDataTable(filePath, "DataTable1", 2));
        }
    }
}

