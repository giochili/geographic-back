using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.DTOs.Windbreak
{
    public class RenamePhotoDTO
    {
        public string  FolderPath { get; set; }
        public int FolderStartNumber { get; set; }
        public int PhotoStartNumber { get; set; }
        public bool Gadanomrilia {  get; set; }
    }
}
