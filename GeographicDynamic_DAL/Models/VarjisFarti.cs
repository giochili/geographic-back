using System;
using System.Collections.Generic;

namespace GeographicDynamic_DAL.Models;

public partial class VarjisFarti
{
    public int Id { get; set; }

    public int? SaxeobaId { get; set; }

    public int? AreaNameId { get; set; }

    public double? VarjisFarti1 { get; set; }

    public virtual Dictionary? Saxeoba { get; set; }
}
