using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Models;
using Microsoft.Extensions.Logging;
namespace GeographicDynamic_DAL.Configurations
{
    public class MapperConfig : Profile
    {
        public MapperConfig()
        {
            CreateMap<VarjisFarti,
                VarjisFartiDTO>().ReverseMap();
            CreateMap<Dictionary,
    DictionaryDTO>().ReverseMap();
            CreateMap<ColumnName, ColumnNameDTO>().ReverseMap();
        }
    }
}
