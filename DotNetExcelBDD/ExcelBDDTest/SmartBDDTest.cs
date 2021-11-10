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
            List<Dictionary<string, string>> exampleList = ExcelBDD.Behavior.GetExampleList(filePath, "Sheet3");
            Assert.IsNotNull(exampleList);
            int count = 0;
            foreach (Dictionary<string, string> dic in exampleList)
            {
                foreach (KeyValuePair<string, string> sample in dic)
                {
                    Console.WriteLine("Show Dictionary: {0} - {1}",sample.Key, sample.Value);
                }
                count++;
            }
            Assert.AreEqual(4, count);
            Assert.AreEqual("Scenario3", exampleList[2]["header"]);
            Assert.AreEqual("V1.3", exampleList[2]["ParamName1"]);
        }
    }
}

