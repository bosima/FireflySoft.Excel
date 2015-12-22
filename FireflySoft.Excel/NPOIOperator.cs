using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace FireflySoft.Excel
{
    public class NPOIOperator
    {
        /// <summary>
        /// Excel数据文件物理路径
        /// </summary>
        private string dataFilePath = null;

        /// <summary>
        /// Excel模板文件物理路径，根据模板生成时需要
        /// </summary>
        private string templateFilePath = null;

        /// <summary>
        /// NPOI Excel工作簿操作接口
        /// </summary>
        private IWorkbook workbook = null;

        /// <summary>
        /// 内容样式字典
        /// </summary>
        private Dictionary<string, ICellStyle> cellStyleDictionary;

        /// <summary>
        /// 标题样式
        /// </summary>
        private ICellStyle titleStyle;

        /// <summary>
        /// 写入时是否使用内置样式
        /// </summary>
        private bool isUseBuiltInStyle;

        /// <summary>
        /// 初始化构造函数
        /// </summary>
        /// <param name="dataFilePath">Excel数据文件物理路径</param>
        public NPOIOperator(string dataFilePath, bool isUseBuiltInStyle)
            : this(string.Empty, dataFilePath, isUseBuiltInStyle)
        {
        }

        /// <summary>
        /// 初始化构造函数
        /// </summary>
        /// <param name="templateFilePath">Excel模板文件物理路径</param>
        /// <param name="dataFilePath">Excel数据文件物理路径</param>
        public NPOIOperator(string templateFilePath, string dataFilePath, bool isUseBuiltInStyle)
        {
            this.isUseBuiltInStyle = isUseBuiltInStyle;
            this.dataFilePath = dataFilePath;
            this.templateFilePath = templateFilePath;

            CreateWorkbook();

            InitTitleStyle();
            InitContentStyle();
        }

        /// <summary>
        /// 初始化内容样式
        /// </summary>
        private void InitContentStyle()
        {
            cellStyleDictionary = new Dictionary<string, ICellStyle>();

            // 内容字体
            IFont contentFont = workbook.CreateFont();
            contentFont.FontHeightInPoints = 9;
            contentFont.FontName = "宋体";

            IDataFormat format = workbook.CreateDataFormat();

            // 日期样式
            ICellStyle dateStyle = workbook.CreateCellStyle();
            dateStyle.SetFont(contentFont);
            dateStyle.WrapText = false;
            dateStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            dateStyle.DataFormat = format.GetFormat("yyyy-MM-dd");
            cellStyleDictionary.Add("date", dateStyle);

            // 日期时间样式
            ICellStyle dateTimeStyle = workbook.CreateCellStyle();
            dateTimeStyle.SetFont(contentFont);
            dateTimeStyle.WrapText = false;
            dateTimeStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            dateTimeStyle.DataFormat = format.GetFormat("yyyy-MM-dd HH:mm:ss");
            cellStyleDictionary.Add("datetime", dateStyle);

            // Money样式
            ICellStyle doubleStyle = workbook.CreateCellStyle();
            doubleStyle.SetFont(contentFont);
            doubleStyle.WrapText = false;
            doubleStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            doubleStyle.DataFormat = HSSFDataFormat.GetBuiltinFormat("0.00");
            cellStyleDictionary.Add("double", doubleStyle);

            // 文本样式
            ICellStyle textStyle = workbook.CreateCellStyle();
            textStyle.SetFont(contentFont);
            textStyle.WrapText = false;
            textStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            cellStyleDictionary.Add("text", textStyle);
        }

        /// <summary>
        /// 初始化标题样式
        /// </summary>
        private void InitTitleStyle()
        {
            // 标题字体
            IFont titleFont = workbook.CreateFont();
            titleFont.FontHeightInPoints = 9;
            titleFont.FontName = "宋体";
            titleFont.Boldweight = (short)FontBoldWeight.Bold;

            // 标题样式
            titleStyle = workbook.CreateCellStyle();
            titleStyle.WrapText = false;
            titleStyle.SetFont(titleFont);
            titleStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            titleStyle.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
        }

        /// <summary>
        /// 获取Excel工作簿
        /// </summary>
        /// <returns></returns>
        public IWorkbook GetWorkBook()
        {
            return workbook;
        }

        /// <summary>
        /// 根据Sheet表名称获取Sheet表
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ISheet GetSheet(string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException("sheetName为空");
            }

            return workbook.GetSheet(sheetName);
        }

        /// <summary>
        /// 根据Sheet表索引位置获取Sheet表
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public ISheet GetSheet(int index)
        {
            if (index >= 0)
            {
                return workbook.GetSheetAt(index);
            }

            return null;
        }

        /// <summary>
        /// 创建指定名称的Sheet表
        /// </summary>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public ISheet CreateSheet(string sheetName)
        {
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                throw new ArgumentNullException("sheetName为空");
            }

            return workbook.CreateSheet(sheetName);
        }

        #region 写数据
        /// <summary>
        /// 将DataTable数据写入到Sheet表，默认从Sheet表的第0行开始
        /// </summary>
        /// <param name="sheetName">Sheet表名</param>
        /// <param name="data">DataTable数据</param>
        /// <param name="isWriteTitle">是否写入DataTable列名作为Sheet表标题</param>
        /// <returns>写入的最后行号</returns>
        public int WriteSheet(string sheetName, DataTable data, bool isWriteTitle)
        {
            if (workbook == null)
            {
                throw new NullReferenceException("workbook为null，请先调用Open方法初始化workbook");
            }

            ISheet sheet = GetSheet(sheetName);
            if (sheet == null)
            {
                CreateSheet(sheetName);
            }

            return WriteSheet(sheet, data, isWriteTitle);
        }

        /// <summary>
        /// 将DataTable数据写入到Sheet表，默认从Sheet表的第0行开始
        /// </summary>
        /// <param name="sheet">Sheet表</param>
        /// <param name="data">DataTable数据</param>
        /// <param name="isWriteTitle">是否写入DataTable列名作为Sheet表标题</param>
        /// <returns>写入的最后行号</returns>
        public int WriteSheet(ISheet sheet, DataTable data, bool isWriteTitle)
        {
            return WriteSheet(sheet, data, isWriteTitle, 0);
        }

        /// <summary>
        /// 将DataTable数据写入到Sheet表
        /// </summary>
        /// <param name="sheet">Sheet表</param>
        /// <param name="data">DataTable数据</param>
        /// <param name="isWriteTitle">是否写入DataTable列名作为Sheet表标题</param>
        /// <param name="startRowNumber">开始写入数据的行号（从0开始）</param>
        /// <returns>写入的最后行号</returns>
        public int WriteSheet(ISheet sheet, DataTable data, bool isWriteTitle, int startRowNumber)
        {
            int endRowNumber = 0;

            // 创建标题行
            if (isWriteTitle == true)
            {
                IRow titleRow = sheet.CreateRow(startRowNumber);
                WriteTitle(titleRow, data.Columns);
                startRowNumber++;
            }

            // 创建数据行
            endRowNumber = WriteContent(sheet, data, startRowNumber);

            // 自动适应列宽
            AutoFitColumnWidth(sheet, data.Columns.Count);

            return endRowNumber;
        }

        /// <summary>
        /// 写入数据到Sheet表标题
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="columns"></param>
        public void WriteTitle(ISheet sheet, DataColumnCollection columns)
        {
            IRow row = sheet.CreateRow(0);

            WriteTitle(row, columns);
        }

        /// <summary>
        /// 写入数据到Row
        /// </summary>
        /// <param name="row"></param>
        /// <param name="columns"></param>
        public void WriteTitle(IRow row, DataColumnCollection columns)
        {
            for (int j = 0; j < columns.Count; ++j)
            {
                var cell = row.CreateCell(j);
                cell.SetCellValue(columns[j].ColumnName);
                cell.CellStyle = titleStyle;
            }
        }

        /// <summary>
        /// 写入数据到内容
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="data"></param>
        /// <param name="strartRowNumber"></param>
        /// <returns></returns>
        public int WriteContent(ISheet sheet, DataTable data, int strartRowNumber)
        {
            // 创建内容行
            for (int i = 0; i < data.Rows.Count; ++i)
            {
                IRow row = sheet.CreateRow(strartRowNumber);
                row.Height = 26 * 20;

                WriteRow(row, data.Columns, data.Rows[i]);

                strartRowNumber++;
            }

            return strartRowNumber;
        }

        /// <summary>
        /// 写入数据到某一行
        /// </summary>
        /// <param name="row"></param>
        /// <param name="data"></param>
        /// <param name="cols"></param>
        public void WriteRow(IRow row, DataColumnCollection cols, DataRow data)
        {
            for (int j = 0; j < cols.Count; ++j)
            {
                var cell = row.CreateCell(j);

                CellDataType cellDataType = CellDataType.None;
                if (data[j] != null)
                {
                    cellDataType = GetCellDataTypeByColumn(cols[j]);
                }

                WriteCell(cell, data[j], cellDataType);
            }
        }

        /// <summary>
        /// 写入数据到单元格
        /// </summary>
        /// <param name="sheet">要写入数据的sheet表实例</param>
        /// <param name="rowNumber">要写入数据的行号</param>
        /// <param name="columnNumber">要写入数据的列号</param>
        /// <param name="value">要写入的数据</param>
        /// <param name="dataType">要写入数据的类型</param>
        public void WriteCell(ISheet sheet, int rowNumber, int columnNumber, object value, CellDataType dataType)
        {
            var row = sheet.GetRow(rowNumber);
            if (row == null)
            {
                row = sheet.CreateRow(rowNumber);
            }

            WriteCell(row, columnNumber, value, dataType);
        }

        /// <summary>
        /// 写入数据到单元格
        /// </summary>
        /// <param name="row">要写入数据的行实例</param>
        /// <param name="columnNumber">要写入数据的列号</param>
        /// <param name="value">要写入的数据</param>
        /// <param name="dataType">要写入数据的类型</param>
        public void WriteCell(IRow row, int columnNumber, object value, CellDataType dataType)
        {
            var cell = row.GetCell(columnNumber);
            if (cell == null)
            {
                cell = row.CreateCell(columnNumber);
            }

            WriteCell(cell, value, dataType);
        }

        /// <summary>
        /// 写入数据到单元格
        /// </summary>
        /// <param name="cell">要写入数据的单元格实例</param>
        /// <param name="value">要写入的数据</param>
        /// <param name="dataType">要写入数据的类型</param>
        public void WriteCell(ICell cell, object value, CellDataType dataType)
        {
            if (dataType == CellDataType.None || dataType == CellDataType.Null || value == null || value == DBNull.Value)
            {
                cell.SetCellType(CellType.Blank);
                return;
            }

            if (value is DateTime)
            {
                cell.SetCellValue((DateTime)value);
            }
            else if (value is int)
            {
                cell.SetCellValue((int)value);
            }
            else if (value is double)
            {
                cell.SetCellValue((double)value);
            }
            else if (value is decimal)
            {
                cell.SetCellValue(Convert.ToDouble(value));
            }
            else
            {
                cell.SetCellValue(value.ToString());
            }

            if (dataType != CellDataType.None)
            {
                switch (dataType)
                {
                    case CellDataType.Date:
                        cell.CellStyle = cellStyleDictionary["date"];
                        break;
                    case CellDataType.DateTime:
                        cell.CellStyle = cellStyleDictionary["datetime"];
                        break;
                    case CellDataType.Double:
                        cell.CellStyle = cellStyleDictionary["double"];
                        break;
                    default:
                        cell.CellStyle = cellStyleDictionary["text"];
                        break;
                }
            }
        }

        /// <summary>
        /// 将内存中的数据写入到Excel数据文件
        /// </summary>
        public void Flush()
        {
            using (var dataFileStream = new FileStream(dataFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                workbook.Write(dataFileStream);
            }
        }
        #endregion

        #region 读数据
        /// <summary>
        /// 自动适应列宽
        /// </summary>
        /// <param name="sheet">需要自适应列宽的sheet表</param>
        /// <param name="columnCount">列数</param>
        public void AutoFitColumnWidth(ISheet sheet, int columnCount)
        {
            //列宽自适应，只对英文和数字有效
            for (int ci = 0; ci < columnCount; ci++)
            {
                sheet.AutoSizeColumn(ci);
            }
            //获取当前列的宽度，然后对比本列的长度，取最大值
            for (int columnNum = 0; columnNum < columnCount; columnNum++)
            {
                int columnWidth = sheet.GetColumnWidth(columnNum) / 256;
                for (int rowNum = 0; rowNum < sheet.LastRowNum; rowNum++)
                {
                    if (rowNum == 0 || rowNum == sheet.LastRowNum - 1 || rowNum == sheet.LastRowNum / 2)
                    {
                        IRow currentRow;

                        //当前行未被使用过
                        if (sheet.GetRow(rowNum) == null)
                        {
                            currentRow = sheet.CreateRow(rowNum);
                        }
                        else
                        {
                            currentRow = sheet.GetRow(rowNum);
                        }

                        if (currentRow.GetCell(columnNum) != null)
                        {
                            ICell currentCell = currentRow.GetCell(columnNum);
                            int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                            if (columnWidth < length)
                            {
                                columnWidth = length;
                            }
                        }
                    }
                }

                sheet.SetColumnWidth(columnNum, columnWidth * 256);
            }
        }

        /// <summary>
        /// 将Sheet表数据读取到DataTable
        /// </summary>
        /// <param name="sheetName">Excel Sheet表名</param>
        /// <param name="isReadTitle">是否将Sheet表标题作为DataTable列名</param>
        /// <returns>读取到的数据，如果Sheet表不存在则返回null</returns>
        public DataTable ReadSheet(string sheetName, bool isReadTitle)
        {
            ISheet sheet = null;
            DataTable data = null;

            if (workbook == null)
            {
                throw new NullReferenceException("workbook为null，请先调用Open方法初始化workbook");
            }

            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = GetSheet(sheetName);
            }

            if (sheet != null)
            {
                data = ReadSheet(sheet, isReadTitle);
            }

            return data;
        }

        /// <summary>
        /// 将Sheet表数据读取到DataTable
        /// </summary>
        /// <param name="sheet">Sheet表</param>
        /// <param name="isReadTitle">是否将Sheet表标题作为DataTable列名</param>
        /// <returns>读取到的数据，如果Sheet表不存在则返回null</returns>
        public DataTable ReadSheet(ISheet sheet, bool isReadTitle)
        {
            return ReadSheet(sheet, isReadTitle, 0);
        }

        /// <summary>
        /// 将Sheet表数据读取到DataTable
        /// </summary>
        /// <param name="sheet">Sheet表</param>
        /// <param name="isReadTitle">是否将Sheet表标题作为DataTable列名</param>
        /// <param name="startRowNumber">起始行号</param>
        /// <returns>读取到的数据，如果Sheet表不存在则返回null</returns>
        public DataTable ReadSheet(ISheet sheet, bool isReadTitle, int startRowNumber)
        {
            DataTable data = null;

            if (sheet != null)
            {
                data = new DataTable();
                IRow titleRow = sheet.GetRow(startRowNumber);

                // 是否读取标题行
                if (isReadTitle)
                {
                    var cols = ReadTitle(titleRow);
                    if (cols != null)
                    {
                        data.Columns.AddRange(cols);
                    }

                    startRowNumber++;
                }

                // 读取数据
                int rowCount = sheet.LastRowNum;
                ReadContent(sheet, startRowNumber, rowCount, data);
            }

            return data;
        }

        /// <summary>
        /// 获取Sheet表的标题，读取第一行
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <returns>标题列数组</returns>
        public DataColumn[] ReadTitle(ISheet sheet)
        {
            IRow firstRow = sheet.GetRow(0);
            if (firstRow == null)
            {
                return new DataColumn[0];
            }

            return ReadTitle(firstRow);
        }

        /// <summary>
        /// 从Row实例获取标题
        /// </summary>
        /// <param name="row">Row实例</param>
        /// <returns>标题列数组</returns>
        public DataColumn[] ReadTitle(IRow row)
        {
            int cellCount = row.LastCellNum;
            DataColumn[] cols = new DataColumn[cellCount];

            int j = 0;
            for (int i = row.FirstCellNum; i < cellCount; i++)
            {
                DataColumn column = new DataColumn(row.GetCell(i).StringCellValue);
                cols[j] = column;
                j++;
            }

            return cols;
        }

        /// <summary>
        /// 读取sheet表数据到DataTable中
        /// </summary>
        /// <param name="sheet">sheet表</param>
        /// <param name="startRowNumber">起始行号</param>
        /// <param name="endRowNumber">终止行号</param>
        /// <param name="table">要读入数据的DataTable</param>
        /// <returns>读取的数据行数</returns>
        public int ReadContent(ISheet sheet, int startRowNumber, int endRowNumber, DataTable table)
        {
            int columnCount = table.Columns.Count;
            int rowCount = 0;
            for (int i = startRowNumber; i <= endRowNumber; ++i)
            {
                IRow row = sheet.GetRow(i);
                if (row == null) continue;

                var rowData = ReadRow(row, table.Columns);
                if (rowData != null)
                {
                    DataRow dataRow = table.NewRow();
                    for (int j = 0; j < rowData.Length; j++)
                    {
                        dataRow[j] = rowData[j];
                    }

                    table.Rows.Add(dataRow);
                }

                rowCount++;
            }

            return rowCount;
        }

        /// <summary>
        /// 读取一行数据
        /// </summary>
        /// <param name="row"></param>
        /// <param name="cols"></param>
        /// <returns></returns>
        public object[] ReadRow(IRow row, DataColumnCollection cols)
        {
            if (cols.Count == 0)
            {
                return null;
            }

            object[] data = new object[cols.Count - row.FirstCellNum];

            for (int j = row.FirstCellNum; j < cols.Count; ++j)
            {
                var cell = row.GetCell(j);
                if (cell != null)
                {
                    CellDataType cellDataType = CellDataType.Text;

                    if (cols.Count > 0)
                    {
                        cellDataType = GetCellDataTypeByColumn(cols[j]);
                    }
                    else
                    {
                        cellDataType = GetCellDataTypeByCellType(cell.CellType);
                    }

                    data[j] = ReadCell(cell, cellDataType);
                }
            }

            return data;
        }

        /// <summary>
        /// 读取单元格数据
        /// </summary>
        /// <param name="sheet">要读取数据的sheet表实例</param>
        /// <param name="rowNumber">要读取数据的行号</param>
        /// <param name="columnNumber">要读取数据的列号</param>
        public object ReadCell(ISheet sheet, int rowNumber, int columnNumber, CellDataType cellDataType)
        {
            var row = sheet.GetRow(rowNumber);
            if (row != null)
            {
                return ReadCell(row, columnNumber, cellDataType);
            }

            return null;
        }

        /// <summary>
        /// 读取单元格数据
        /// </summary>
        /// <param name="row">要读取数据的行实例</param>
        /// <param name="columnNumber">要读取数据的列号</param>
        public object ReadCell(IRow row, int columnNumber, CellDataType cellDataType)
        {
            var cell = row.GetCell(columnNumber);
            if (cell != null)
            {
                return ReadCell(cell, cellDataType);
            }

            return null;
        }

        /// <summary>
        /// 读取单元格数据
        /// </summary>
        /// <param name="cell">要读取数据的单元格</param>
        public object ReadCell(ICell cell, CellDataType cellDataType)
        {
            if (cellDataType == CellDataType.Boolean)
            {
                return cell.BooleanCellValue;
            }
            else if (cellDataType == CellDataType.Date)
            {
                return cell.DateCellValue;
            }
            else if (cellDataType == CellDataType.DateTime)
            {
                return cell.DateCellValue;
            }
            else if (cellDataType == CellDataType.Double)
            {
                return cell.NumericCellValue;
            }
            else if (cellDataType == CellDataType.Formula)
            {
                return cell.CellFormula;
            }
            else if (cellDataType == CellDataType.Int)
            {
                return cell.NumericCellValue;
            }
            else if (cellDataType == CellDataType.RichText)
            {
                return cell.RichStringCellValue;
            }
            else if (cellDataType == CellDataType.Text)
            {
                return cell.StringCellValue;
            }
            else if (cellDataType == CellDataType.None)
            {
                return string.Empty;
            }
            else if (cellDataType == CellDataType.Null)
            {
                return null;
            }

            return null;
        }
        #endregion

        /// <summary>
        /// 根据模板或数据excel文件创建工作簿实例
        /// </summary>
        /// <returns></returns>
        private IWorkbook CreateWorkbook()
        {
            if (!string.IsNullOrEmpty(templateFilePath))
            {
                using (var templateFile = new FileStream(templateFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    if (templateFilePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook(templateFile);
                    }
                    else if (templateFilePath.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook(templateFile);
                    }
                }
            }
            else
            {
                if (File.Exists(dataFilePath))
                {
                    using (var dataFile = new FileStream(dataFilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
                    {
                        if (dataFilePath.IndexOf(".xlsx") > 0)
                        {
                            workbook = new XSSFWorkbook(dataFile);
                        }
                        else if (dataFilePath.IndexOf(".xls") > 0)
                        {
                            workbook = new HSSFWorkbook(dataFile);
                        }
                    }
                }
                else
                {
                    if (dataFilePath.IndexOf(".xlsx") > 0)
                    {
                        workbook = new XSSFWorkbook();
                    }
                    else if (dataFilePath.IndexOf(".xls") > 0)
                    {
                        workbook = new HSSFWorkbook();
                    }
                }
            }

            return workbook;
        }

        /// <summary>
        /// 根据列属性获取数据类型
        /// </summary>
        /// <param name="column">列构造实例</param>
        /// <returns></returns>
        private CellDataType GetCellDataTypeByColumn(DataColumn column)
        {
            var cellDataType = CellDataType.None;

            if (isUseBuiltInStyle)
            {
                if (column.ExtendedProperties != null && column.ExtendedProperties.Count > 0)
                {
                    if (column.ExtendedProperties.ContainsKey("DataType"))
                    {
                        var propertyValue = column.ExtendedProperties["DataType"].ToString();
                        cellDataType = GetDataTypeFromColumnExtendedProperty(propertyValue);
                    }
                }
                else
                {
                    Type columnDataType = column.DataType;
                    cellDataType = GetDataTypeFromColumnDataType(columnDataType);
                }
            }

            return cellDataType;
        }

        /// <summary>
        /// 根据单元格类型获取数据类型
        /// </summary>
        /// <param name="cellType">单元格类型</param>
        /// <returns></returns>
        private CellDataType GetCellDataTypeByCellType(CellType cellType)
        {
            var cellDataType = CellDataType.None;

            if (cellType == CellType.Blank)
            {
                cellDataType = CellDataType.None;
            }
            else if (cellType == CellType.Boolean)
            {
                cellDataType = CellDataType.Boolean;
            }
            else if (cellType == CellType.Error)
            {
                cellDataType = CellDataType.Text;
            }
            else if (cellType == CellType.Formula)
            {
                cellDataType = CellDataType.Formula;
            }
            else if (cellType == CellType.Numeric)
            {
                cellDataType = CellDataType.Double;
            }
            else if (cellType == CellType.String)
            {
                cellDataType = CellDataType.Text;
            }
            else if (cellType == CellType.Unknown)
            {
                cellDataType = CellDataType.Text;
            }

            return cellDataType;
        }

        /// <summary>
        /// 从数据列的类型获取DataType
        /// </summary>
        /// <param name="type">数据类型</param>
        /// <returns>单元格数据类型</returns>
        private static CellDataType GetDataTypeFromColumnDataType(Type type)
        {
            var dataType = CellDataType.None;

            if (type == typeof(DateTime))
            {
                dataType = CellDataType.DateTime;
            }
            else if (type == typeof(int))
            {
                dataType = CellDataType.Int;
            }
            else if (type == typeof(double))
            {
                dataType = CellDataType.Double;
            }
            else if (type == typeof(decimal))
            {
                dataType = CellDataType.Double;
            }
            else if (type == typeof(bool))
            {
                dataType = CellDataType.Boolean;
            }
            else
            {
                dataType = CellDataType.Text;
            }

            return dataType;
        }

        /// <summary>
        /// 从列的扩展属性获取DataType
        /// </summary>
        /// <param name="propertyValue">属性值</param>
        /// <returns>单元格数据类型</returns>
        private static CellDataType GetDataTypeFromColumnExtendedProperty(string propertyValue)
        {
            var dataType = CellDataType.None;

            switch (propertyValue)
            {
                case "Null":
                    dataType = CellDataType.Null;
                    break;
                case "Date":
                    dataType = CellDataType.Date;
                    break;
                case "DateTime":
                    dataType = CellDataType.DateTime;
                    break;
                case "Int":
                    dataType = CellDataType.Int;
                    break;
                case "Double":
                    dataType = CellDataType.Double;
                    break;
                case "Boolean":
                    dataType = CellDataType.Boolean;
                    break;
                case "Formula":
                    dataType = CellDataType.Formula;
                    break;
                case "RichText":
                    dataType = CellDataType.RichText;
                    break;
                default:
                    dataType = CellDataType.Text;
                    break;
            }

            return dataType;
        }
    }
}
