using Microsoft.AspNetCore.Mvc;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;

namespace GeographicDynamicWebAPI.Controllers
{

    [ApiController]
    public class ColumnName : Controller
    {
        public readonly IColumnName _columnName;

        public ColumnName(IColumnName columnName)
        {
            _columnName = columnName;
        }


        [HttpGet("ColumnNameTransfer")]
        public IActionResult ColumnNameTransfer(string ExcelPath)
        {
            var result = _columnName.ColumnNameTransfer(ExcelPath);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpGet("ColumnNameTransferFromAccess")]
        public IActionResult ColumnNameTransferFromAccess(string AccessPath, string AccessSheetName)
        {
            var result = _columnName.ColumnNameTransferFromAccess(AccessPath, AccessSheetName);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpGet("GetSQLColumnNamesList")]
        public IActionResult GetSQLColumnNamesList()
        {
            var result = _columnName.GetSQLColumnNamesList();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpPost("SaveColumnName")]
        public IActionResult SaveColumnName(List<ColumnNameDTO> columnNameDTO)
        {
            var result = _columnName.SaveColumnName(columnNameDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpPost ("DeleteRow")]

        public IActionResult DeleteRow(ColumnNameDTO Id)
        {
            var result = _columnName.DeleteRow(Id);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
    }
}
