using System;
using System.Collections.Generic;

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
            List<Dictionary<string, string>> exampleList = new List<Dictionary<string, string>>();
            return exampleList;
        }
    }
}
