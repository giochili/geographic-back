using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BotReestriClassLibrary.Wrapper;
using BotReestriClassLibrary.DTOs;

namespace BotReestriClassLibrary.Interface
{
    public interface IChromeBot
    
    {
        public Result<ChromeBotDTO> BotChromeArguments(string ExcelPath, string Destination);
    


    }

}
