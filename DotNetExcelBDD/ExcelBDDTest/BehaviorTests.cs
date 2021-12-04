using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelBDD.Tests
{
    [TestClass()]
    public class BehaviorTests
    {
        [TestMethod()]
        public void GetRowListTest()
        {
            //Assert.Fail();
        }

        [TestMethod()]
        public void GetDataTableEnumerableTest()
        {

        }

        [TestMethod()]
        public void GetDataTableByArrayTest()
        {
            String currentPath = Directory.GetCurrentDirectory();
            System.Console.WriteLine(currentPath);
            String filePath = currentPath.Substring(0, currentPath.IndexOf("DotNetExcelBDD")) + "BDDExcel\\DataTableBDD.xlsx";
            System.Console.WriteLine(filePath);
            IEnumerable<object[]> exampleList = ExcelBDD.Behavior.GetDataTableByArray(filePath, "DataTable2", 2);
            Assert.IsNotNull(exampleList);
            int count = 0;
            foreach (var item in exampleList)
            {
                System.Console.WriteLine(item[0]);
                count++;
            }
            Assert.AreEqual(6, count);
        }
    }
}