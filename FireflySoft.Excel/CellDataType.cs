using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FireflySoft.Excel
{
    /// <summary>
    /// 单元格数据类型
    /// </summary>
    public enum CellDataType
    {
        None = 0,
        Null = 1,
        Text,
        Date,
        DateTime,
        Int,
        Double,
        Formula,
        Boolean,
        RichText
    }
}
