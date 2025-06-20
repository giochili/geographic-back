using GeographicDynamic_DAL.DTOs.Windbreak;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Interface
{
    public interface IWriteOnImage
    {
        string WriteInfoOnImage(RenamePhotoDTO renamePhotoDTO);
    }
}
