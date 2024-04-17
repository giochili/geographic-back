using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Interface
{
    public interface IColumnName
    {
        public Result<ColumnNameDTO> ColumnNameTransfer(string ExcelPath);

        public Result<ColumnNameDTO> ColumnNameTransferFromAccess(string AccessPath, string AccessSheetName);

        public Result<ColumnNameDTO> GetSQLColumnNamesList();

        public Result<bool> SaveColumnName(List<ColumnNameDTO> columnNameDTO);

        public Result<int> DeleteRow(ColumnNameDTO Id);

    }
}
