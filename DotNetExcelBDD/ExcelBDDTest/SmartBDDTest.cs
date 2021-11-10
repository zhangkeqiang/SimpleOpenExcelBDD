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
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\ExcelBDD.xlsx";
            System.Console.WriteLine(filePath);
            IEnumerable<object[]> exampleList = ExcelBDD.Behavior.GetExampleList(filePath, "Sheet3");
            Assert.IsNotNull(exampleList);
            int count = 0;
            foreach (var item in exampleList)
            {
                Dictionary<string, string> dic = (Dictionary<string, string>)item[0];
                Console.WriteLine(item[0]);
                Console.WriteLine(dic["header"]);
                count++;
            }
            Assert.AreEqual(4, count);
        }

      
    }
}

