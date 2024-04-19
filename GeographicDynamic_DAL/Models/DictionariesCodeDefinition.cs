using System;
using System.Collections.Generic;

namespace GeographicDynamic_DAL.Models;

public partial class DictionariesCodeDefinition
{
    public int Id { get; set; }

    public string? Definition { get; set; }

    public int? Code { get; set; }
}
