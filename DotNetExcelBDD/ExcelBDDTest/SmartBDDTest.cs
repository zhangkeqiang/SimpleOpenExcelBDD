using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
namespace ExcelBDD.Tests
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


        public static IEnumerable<object[]> GetSmartBDDList()
        {
            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\ExcelBDD.xlsx";
            Console.WriteLine(filePath);
            return ExcelBDD.Behavior.GetExampleEnumerable(filePath, "SmartBDD", "", "Expected");
        }

        [TestMethod]
        public void TestGetSmartBDDList()
        {
            int i = 1;
            foreach (var obj in GetSmartBDDList())
            {
                Console.WriteLine("============ {0} ============", i++);
                foreach (KeyValuePair<string, string> item in ((Dictionary<string, string>)obj[0]))
                {
                    Console.WriteLine("Dictionary: {0} - {1}", item.Key, item.Value);
                }
            }
        }

        [DataTestMethod]
        [DynamicData(nameof(GetSmartBDDList), DynamicDataSourceType.Method)]
        public void TestSmartBDD(Dictionary<string, string> paramDic)
        {
            foreach (KeyValuePair<string, string> item in paramDic)
            {
                Console.WriteLine("Dictionary: {0} - {1}", item.Key, item.Value);
            }

            Assert.AreEqual("Scenario1", paramDic["Header1Name"]);
            Assert.AreEqual("V1.2", paramDic["ParamName1InSet2Value"]);
            Assert.AreEqual("V2.2", paramDic["ParamName2InSet2Value"]);
            Assert.AreEqual("3", paramDic["MaxBlankThreshold"]);
            Assert.AreEqual("", paramDic["ParamName3Value"]);

            String currentPath = Directory.GetCurrentDirectory();
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\" + paramDic["ExcelFileName"];
            List<Dictionary<string, string>> list = ExcelBDD.Behavior.GetExampleList(filePath, paramDic["SheetName"], paramDic["HeaderMatcher"], paramDic["HeaderUnmatcher"]);
            Assert.IsNotNull(list);
            Assert.AreEqual(paramDic["TestDataSetCount"], list.Count.ToString());
            Assert.AreEqual(paramDic["FirstGridValue"], list[0]["ParamName1"]);
            Assert.AreEqual(paramDic["ParamName1InSet2Value"], list[1]["ParamName1"]);
            Assert.AreEqual("V1.3", list[2]["ParamName1"]);
            Assert.AreEqual("V1.4", list[3]["ParamName1"]);

            Assert.AreEqual("V2.1", list[0]["ParamName2"]);
            Assert.AreEqual(paramDic["ParamName2InSet2Value"], list[1]["ParamName2"]);

            Assert.AreEqual("", list[0]["ParamName3"]);
            Assert.AreEqual("", list[1]["ParamName3"]);
            Assert.AreEqual("", list[2]["ParamName3"]);
            Assert.AreEqual("", list[3]["ParamName3"]);

            Assert.AreEqual("2021/4/30", list[0]["ParamName4"]);
            Assert.AreEqual("0", list[1]["ParamName4"]);
            Assert.AreEqual("1", list[2]["ParamName4"]);
            Assert.AreEqual(paramDic["LastGridValue"], list[3]["ParamName4"]);
        }


        [TestMethod]
        public void CheckBasic()
        {
            Assert.AreNotEqual("", null);
            Assert.IsTrue("Abcdefed".IndexOf("") >= 0);
        }
    }
}

