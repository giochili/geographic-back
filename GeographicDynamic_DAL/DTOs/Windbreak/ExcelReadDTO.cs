using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.DTOs.Windbreak
{
    public class ExcelReadDTO
    {

        public string ExcelPath { get; set; }
        public int UnicIDStartNumber { get; set; }
        public string ExcelDestinationPath { get; set; }
        public string AccessFilePath { get; set; }
        public int ProjectNameID { get; set; }
        public bool CalcVarjisFartiCheckbox {  get; set; }
        public string AccessShitName {  get; set; }
        public int ProjectID { get; set; }
        public int EtapiID { get; set; }

        //ფოტოების გადასანომრად ველები
        public string FolderPath { get; set; }
        public int PhotoStartNumber { get; set; }

    }
}
