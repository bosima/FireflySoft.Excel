using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FireflySoft.Excel
{
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
