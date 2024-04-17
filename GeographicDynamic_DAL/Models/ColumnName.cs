using System;
using System.Collections.Generic;

namespace GeographicDynamic_DAL.Models;

public partial class ColumnName
{
    public int Id { get; set; }

    public string? Sqlname { get; set; }

    public string? ExcelName { get; set; }

    public int? SortValue { get; set; }

    public string? DataType { get; set; }

    public int? ColN { get; set; }

    public string? AccessName { get; set; }

    public bool? IsAccessToExcel { get; set; }

    public string? GroupMethod { get; set; }
}
