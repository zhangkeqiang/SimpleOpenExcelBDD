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
                    Console.WriteLine("Show Dictionary: {0} - {1}", sample.Key, sample.Value);
                }
                count++;
            }
            Assert.AreEqual(4, count);
            Assert.AreEqual("Scenario3", exampleList[2]["header"]);
            Assert.AreEqual("V1.3", exampleList[2]["ParamName1"]);
        }

        [DataTestMethod]
        [DynamicData(nameof(GetSmartBDDList), DynamicDataSourceType.Method)]
        public void TestSmartBDD(Dictionary<string, string> paramDic)
        {
            foreach (KeyValuePair<string, string> item in paramDic)
            {
                Console.WriteLine("Dictionary: {0} - {1}", item.Key, item.Value);
            }

            Assert.AreEqual("Scenario1",paramDic["Header1Name"]);
            Assert.AreEqual("V1.2",paramDic["ParamName1InSet2Value"]);
            Assert.AreEqual("V2.2",paramDic["ParamName2InSet2Value"]);
            Assert.AreEqual("3",paramDic["MaxBlankThreshold"]);
            Assert.AreEqual("",paramDic["ParamName3Value"]);
        }

        [TestMethod]
        public void TestGetSmartBDDList()
        {
            foreach (var obj in GetSmartBDDList())
            {
                foreach (KeyValuePair<string, string> item in ((Dictionary<string, string>)obj[0]))
                {
                    Console.WriteLine("Dictionary: {0} - {1}", item.Key, item.Value);
                }
            }
        }
        public static IEnumerable<object[]> GetSmartBDDList()
        {
            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\ExcelBDD.xlsx";
            Console.WriteLine(filePath);
            return ExcelBDD.Behavior.ConvertToIEnumerable(ExcelBDD.Behavior.GetExampleList(filePath, "SmartBDD"));
        }
    }
}

