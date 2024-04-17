using GeographicDynamic_DAL.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.DTOs.Windbreak
{
    public class VarjisFartiDTO
    {
        public int? Id { get; set; }
        public string? Name { get;set; }
        public int? SaxeobaId { get; set; }
        public int? AreaNameId { get; set; }
        public double? VarjisFarti1 { get; set; }
        public virtual Dictionary? Saxeoba { get; set; }

    }
}
