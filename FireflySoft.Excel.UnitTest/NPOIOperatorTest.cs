using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using NPOI.SS.UserModel;
using System.Data;

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

        [TestMethod]
        public void TestGetSheetByName()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "data.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet1 = oper.GetSheet("Sheet1");
            Assert.IsNotNull(sheet1);

            var sheet100 = oper.GetSheet("Sheet100");
            Assert.IsNull(sheet100);
        }

        [TestMethod]
        public void TestGetSheetByIndex()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "data.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet1 = oper.GetSheet(0);
            Assert.IsNotNull(sheet1);

            var sheet100 = oper.GetSheet(100);
            Assert.IsNull(sheet100);
        }

        [TestMethod]
        public void TestCreateSheet()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "nothisfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var firstSheet = oper.CreateSheet("FirstSheet");
            Assert.IsNotNull(firstSheet);

            var secondSheet = oper.CreateSheet("SecondSheet");
            Assert.IsNotNull(secondSheet);

            Assert.AreEqual(2, oper.GetSheetCount());
        }

        [TestMethod]
        public void TestGetSheetCount()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheetCount = oper.GetSheetCount();
            Assert.AreEqual(1, oper.GetSheetCount());
        }

        [TestMethod]
        public void TestWriteSheetWithSheetName()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var writeLineNumber = oper.WriteSheet("Sheet3", dt, true);
            Assert.AreEqual(4, writeLineNumber);

            var readDT = oper.ReadSheet("Sheet3", true);
            Assert.AreEqual(3, readDT.Rows.Count);
        }

        [TestMethod]
        public void TestWriteSheetWithSheetInstance()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var writeLineNumber = oper.WriteSheet(sheet, dt, true);
            Assert.AreEqual(4, writeLineNumber);

            var readDT = oper.ReadSheet("Sheet3", true);
            Assert.AreEqual(3, readDT.Rows.Count);
        }

        [TestMethod]
        public void TestReadSheetAndUseCellType()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "data.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet(0);
            var readDT = oper.ReadSheet(sheet, true, 0, true);

            Assert.AreEqual(2, readDT.Rows.Count);
            Assert.AreEqual(1, readDT.Rows[0][1]);
        }

        private DataTable GetTestData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("学号", typeof(string));
            dt.Columns.Add("姓名", typeof(string));
            dt.Columns.Add("性别", typeof(int));
            dt.Columns.Add("班级", typeof(string));

            var birthdayColumn = dt.Columns.Add("出生日期", typeof(DateTime));
            birthdayColumn.ExtendedProperties.Add("DataType", "Date");

            var row1 = dt.NewRow();
            row1[0] = "030601";
            row1[1] = "张三";
            row1[2] = 1;
            row1[3] = "一班";
            row1[4] = new DateTime(2000, 1, 1);
            dt.Rows.Add(row1);

            var row2 = dt.NewRow();
            row2[0] = "030602";
            row2[1] = "李四";
            row2[2] = 0;
            row2[3] = "一班";
            row2[4] = new DateTime(2000, 8, 8);
            dt.Rows.Add(row2);

            var row3 = dt.NewRow();
            row3[0] = "030603";
            row3[1] = "王五";
            row3[2] = 1;
            row3[3] = "二班";
            row3[4] = new DateTime(1999, 6, 6);
            dt.Rows.Add(row3);

            return dt;
        }
    }
}
