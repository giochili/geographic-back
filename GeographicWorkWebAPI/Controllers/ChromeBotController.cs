using Microsoft.AspNetCore.Mvc;
using BotReestriClassLibrary;
using BotReestriClassLibrary.Interface;
using BotReestriClassLibrary.DTOs;
using BotReestriClassLibrary.Repository;
using BotReestriClassLibrary.Wrapper;
using GeographicDynamic_DAL.Interface;
namespace GeographicDynamicWebAPI.Controllers
{

    [ApiController]
    public class ChromeBot : Controller
    {
        private readonly IChromeBot _chromeBot;

        public ChromeBot(IChromeBot chromeBot)
        {
            _chromeBot = chromeBot;
        }

        [HttpPost("BotChromeArguments")]
        public IActionResult BotChromeArguments(ChromeBotDTO chromeBotDTO)
        {
            var result = _chromeBot.BotChromeArguments(chromeBotDTO.ExcelPath, chromeBotDTO.Destination);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
    }
}
