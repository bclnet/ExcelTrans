using System;

namespace ExcelTrans
{
    public static class TestExtensions
    {
        public static string ToLocalString(this string value) => value.Replace(@"
", Environment.NewLine);
    }
}
