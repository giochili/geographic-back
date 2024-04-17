using GeographicDynamic_DAL.Interface;
using Microsoft.AspNetCore.Mvc;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Models;
using AutoMapper;
using GeographicDynamic_DAL.DTOs;

namespace GeographicDynamicWebAPI.Controllers
{

    [ApiController]
    public class VarjisFartiController : Controller
    {

        private readonly IVarjisFarti _varjisFarti;

        public VarjisFartiController(IVarjisFarti varjisFarti)
        {
            _varjisFarti = varjisFarti;
        }
        [HttpPost("SaveVarjisFarti")] // Define the route for the POST action
        public IActionResult SaveVarjisFarti(List<VarjisFartiDTO> varjisFartiDTO)
        {
            var result = _varjisFarti.SaveVarjisFarti(varjisFartiDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("SaveSaxeobebi")]

        public IActionResult SaveSaxeobebi(List<DictionaryDTO> dictionaryDTO)
        {
            var result = _varjisFarti.SaveSaxeobebi(dictionaryDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("DeleteVarjisfarti")]
        public IActionResult DeleteVarjisfarti(VarjisFartiDTO varjisFartiDTO)
        {
            var result = _varjisFarti.DeleteVarjisfarti(varjisFartiDTO);
            if(result.Success) return Ok( result);
            return BadRequest(result);
        }
    }
}
