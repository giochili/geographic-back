using System;
using System.Collections.Generic;

namespace GeographicDynamic_DAL.Models;

public partial class Dictionary
{
    public int Id { get; set; }

    public string? Name { get; set; }

    public int? Code { get; set; }

    public int? Sort { get; set; }

    public virtual ICollection<VarjisFarti> VarjisFartis { get; set; } = new List<VarjisFarti>();
}
