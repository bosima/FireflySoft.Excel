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
        public void TestWriteSheetWithSheetNameAndNoTitle()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var writeLineNumber = oper.WriteSheet("Sheet3", dt, false);
            Assert.AreEqual(3, writeLineNumber);

            var readDT = oper.ReadSheet("Sheet3", false);
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
        public void TestWriteSheetWithSheetInstanceAndNoTitle()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var writeLineNumber = oper.WriteSheet(sheet, dt, false);
            Assert.AreEqual(3, writeLineNumber);

            var readDT = oper.ReadSheet("Sheet3", false);
            Assert.AreEqual(3, readDT.Rows.Count);
        }

        [TestMethod]
        public void TestWriteSheetWithSheetInstanceAndRowNumber()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var writeLineNumber = oper.WriteSheet(sheet, dt, true, 2);
            Assert.AreEqual(6, writeLineNumber);

            var readDT = oper.ReadSheet(sheet, true, 2);
            Assert.AreEqual(3, readDT.Rows.Count);

            Assert.AreEqual("李四", readDT.Rows[1]["姓名"].ToString());
        }

        [TestMethod]
        public void TestWriteSheetWithSheetInstanceAndRowNumberAndNoTitle()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var writeLineNumber = oper.WriteSheet(sheet, dt, false, 2);
            Assert.AreEqual(5, writeLineNumber);

            var readDT = oper.ReadSheet(sheet, false, 2);
            Assert.AreEqual(3, readDT.Rows.Count);

            Assert.AreEqual("李四", readDT.Rows[1]["列2"].ToString());
        }

        [TestMethod]
        public void TestWriteTitle()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            oper.WriteTitle(sheet, dt.Columns);

            var cols = oper.ReadTitle(sheet);
            Assert.AreEqual(5, cols.Length);
            Assert.AreEqual("姓名", cols[1].ColumnName);
        }

        [TestMethod]
        public void TestWriteTitleWithRow()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var row = sheet.CreateRow(0);
            oper.WriteTitle(row, dt.Columns);

            var cols = oper.ReadTitle(sheet);
            Assert.AreEqual(5, cols.Length);
            Assert.AreEqual("姓名", cols[1].ColumnName);
        }

        [TestMethod]
        public void TestWriteContent()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var writeLineNumber = oper.WriteContent(sheet, dt, 2);
            Assert.AreEqual(5, writeLineNumber);

            var readDT = oper.ReadSheet(sheet, false, 2);
            Assert.AreEqual(3, readDT.Rows.Count);

            Assert.AreEqual("李四", readDT.Rows[1]["列2"].ToString());
        }

        [TestMethod]
        public void TestWriteRow()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var row = sheet.CreateRow(1);
            oper.WriteRow(row, dt.Columns, dt.Rows[1]);

            var readDT = oper.ReadSheet(sheet, false, 1);
            Assert.AreEqual(1, readDT.Rows.Count);
            Assert.AreEqual("李四", readDT.Rows[0]["列2"].ToString());
        }

        [TestMethod]
        public void TestWriteCellWithSheet()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            oper.WriteCell(sheet, 1, 1, dt.Rows[1]["姓名"], CellDataType.Text);
            oper.WriteCell(sheet, 1, 2, dt.Rows[1]["性别"], CellDataType.Int);

            var readDT = oper.ReadSheet(sheet, false, 1, true);
            Assert.AreEqual(1, readDT.Rows.Count);
            Assert.AreEqual("李四", readDT.Rows[0]["列1"].ToString());
            Assert.AreEqual(0, readDT.Rows[0]["列2"]);
        }

        [TestMethod]
        public void TestWriteCellWithRow()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");
            var row = sheet.CreateRow(0);
            oper.WriteCell(row, 1, dt.Rows[1]["姓名"], CellDataType.Text);
            oper.WriteCell(row, 2, dt.Rows[1]["性别"], CellDataType.Int);

            var readDT = oper.ReadSheet(sheet, false, 0, true);
            Assert.AreEqual(1, readDT.Rows.Count);
            Assert.AreEqual("李四", readDT.Rows[0]["列1"].ToString());
            Assert.AreEqual(0, readDT.Rows[0]["列2"]);
        }

        [TestMethod]
        public void TestWriteCell()
        {
            var dt = GetTestData();

            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "onesheetfile.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var sheet = oper.GetSheet("Sheet3");

            var row = sheet.CreateRow(0);

            var cell1 = row.CreateCell(0);
            oper.WriteCell(cell1, dt.Rows[1]["姓名"], CellDataType.Text);

            var cell2 = row.CreateCell(1);
            oper.WriteCell(cell2, dt.Rows[1]["性别"], CellDataType.Int);

            var readDT = oper.ReadSheet(sheet, false, 0, true);
            Assert.AreEqual(1, readDT.Rows.Count);
            Assert.AreEqual("李四", readDT.Rows[0]["列1"].ToString());
            Assert.AreEqual(0, readDT.Rows[0]["列2"]);
        }

        [TestMethod]
        public void TestReadSheetWithName()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var readDT = oper.ReadSheet("Sheet1", true);
            Assert.AreEqual(6, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[3]["姓名"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[3]["出生日期"]);
        }

        [TestMethod]
        public void TestReadSheetWithNameAndNoTitle()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);

            var readDT = oper.ReadSheet("Sheet1", false);
            Assert.AreEqual(7, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[4]["列2"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[4]["列5"]);
        }

        [TestMethod]
        public void TestReadSheetWithSheet()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var readDT = oper.ReadSheet(sheet, true);
            Assert.AreEqual(6, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[3]["姓名"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[3]["出生日期"]);
        }

        [TestMethod]
        public void TestReadSheetWithSheetAndNoTitle()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var readDT = oper.ReadSheet(sheet, false);
            Assert.AreEqual(7, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[4]["列2"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[4]["列5"]);
        }

        [TestMethod]
        public void TestReadSheetWithFirstLine()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var readDT = oper.ReadSheet(sheet, true, 0);
            Assert.AreEqual(6, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[3]["姓名"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[3]["出生日期"]);
        }

        [TestMethod]
        public void TestReadSheetWithSencondLine()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var readDT = oper.ReadSheet(sheet, false, 1);
            Assert.AreEqual(6, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[3]["列2"].ToString());
            Assert.AreEqual("14-3月-1988", readDT.Rows[3]["列5"]);
        }

        [TestMethod]
        public void TestReadSheetWithCellType()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var readDT = oper.ReadSheet(sheet, true, 0, true);
            Assert.AreEqual(6, readDT.Rows.Count);
            Assert.AreEqual("猴六", readDT.Rows[3]["姓名"].ToString());
            Assert.AreEqual(new DateTime(1988, 3, 14), readDT.Rows[3]["出生日期"]);
        }

        [TestMethod]
        public void TestReadTitleWithSheet()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var cols = oper.ReadTitle(sheet);

            Assert.AreEqual(5, cols.Length);
            Assert.AreEqual("出生日期", cols[4].ColumnName);
        }

        [TestMethod]
        public void TestReadTitleWithRow()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var firstRow = sheet.GetRow(0);

            var cols = oper.ReadTitle(firstRow);
            Assert.AreEqual(5, cols.Length);
            Assert.AreEqual("出生日期", cols[4].ColumnName);
        }

        [TestMethod]
        public void TestReadContent()
        {
            string dataFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "file", "forread.xlsx");
            NPOIOperator oper = new NPOIOperator(dataFilePath, true);
            var sheet = oper.GetSheet("Sheet1");
            var dataTable = GetTestDataSchema();
            var readCount = oper.ReadContent(sheet, 1, 6, dataTable);

            Assert.AreEqual(6, readCount);
            Assert.AreEqual("猴六", dataTable.Rows[3]["姓名"].ToString());
            Assert.AreEqual(new DateTime(1988, 3, 14), dataTable.Rows[3]["出生日期"]);
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
            DataTable dt = GetTestDataSchema();

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

        private DataTable GetTestDataSchema()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("学号", typeof(string));
            dt.Columns.Add("姓名", typeof(string));
            dt.Columns.Add("性别", typeof(int));
            dt.Columns.Add("班级", typeof(string));

            var birthdayColumn = dt.Columns.Add("出生日期", typeof(DateTime));
            birthdayColumn.ExtendedProperties.Add("DataType", "Date");
            return dt;
        }
    }
}
