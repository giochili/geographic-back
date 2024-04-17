using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.DTOs.Windbreak
{
    public class ColumnNameDTO
    {
        public int? Id { get; set; }
        public string? Sqlname { get; set; }
        public string? ExcelName { get; set; }
        public int? SortValue { get; set; }
        public string? DataType { get; set; }
        public string? AccessName { get; set; }
        public string? GroupMethod { get; set; }
        public bool? IsAccessToExcel {  get; set; } 
        public int? ColN { get; set; }
    }
}
