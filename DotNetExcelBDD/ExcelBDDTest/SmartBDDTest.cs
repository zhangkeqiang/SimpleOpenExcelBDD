using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.IO;
using ExcelBDD;
namespace ExcelBDDTest
{
    [TestClass]
    public class SmartBDDTest
    {
        [TestMethod]
        public void TestGetExampleList()
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
                Dictionary<string, string> dic = (Dictionary<string, string>)item[0];
                Console.Write(dic.ToString());
                Console.WriteLine(dic["Header01"]);
                count++;
            }
            Assert.AreEqual(7, count);
        }

      
    }
}

