using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using NPOI.SS.UserModel;

namespace FireflySoft.Excel.UnitTest
{
    [TestClass]
    public class NPOIOperatorTest
    {
        [TestMethod]
        public void TestGetWorkBook()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "data.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var book = oper.GetWorkBook();

            Assert.IsNotNull(book);
        }
    }
}
