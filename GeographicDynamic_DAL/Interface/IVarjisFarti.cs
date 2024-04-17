using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using GeographicDynamicWebAPI.Wrappers;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.DTOs;

namespace GeographicDynamic_DAL.Interface
{
    public interface IVarjisFarti
    {
        public Result<bool> SaveVarjisFarti(List<VarjisFartiDTO> varjisFartiDTO);
        public Result<bool> SaveSaxeobebi(List<DictionaryDTO> dictionaryDTO);
        public Result<VarjisFartiDTO> DeleteVarjisfarti(VarjisFartiDTO varjisFartiDTO);
    }
}
