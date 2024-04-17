﻿using Microsoft.AspNetCore.Mvc;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Models;
using GeographicDynamicWebAPI.Wrappers;
namespace GeographicDynamicWebAPI.Controllers
{
    [ApiController]
    public class WindbreakController : Controller
    {
        /*test*/
        private readonly IWindbreak _windbreak;

        public WindbreakController(IWindbreak windbreak)
        {
            _windbreak = windbreak;
        }
        [HttpPost("RenamePhotosInFolder")]
        public IActionResult RenamePhotosInFolder(RenamePhotoDTO renamePhotoDTO)
        {
            var result = _windbreak.RenamePhotosInFolder(renamePhotoDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("ExcelCalculations")]
        public IActionResult ExcelCalculations(ExcelReadDTO excelReadDTO)
        {
            var result = _windbreak.ExcelCalculations(excelReadDTO);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpGet("GetProjectNamesList")]
        public IActionResult GetProjectNamesList()
        {
            var result = _windbreak.GetProjectNames();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpGet("getSaxeobaList")]
        public IActionResult getSaxeobaList()
        {
            var result = _windbreak.getSaxeobaList();
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }
        [HttpGet("GetVarjisFartebiList")]
        public IActionResult GetVarjisFartebiList(int AreaNameID)
        {
            var result = _windbreak.GetVarjisFartebi(AreaNameID);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost ("GetCheckPhotoDate")]
        public IActionResult GetCheckPhotoDate(CheckPhotoDateDTO checkPhotoDateDTO)
        {
            var result = _windbreak.GetCheckPhotoDate(checkPhotoDateDTO.folderPath, checkPhotoDateDTO.resultPath);
            if (result.Success) return Ok(result);
            return BadRequest(result);
        }

        [HttpPost("FotoebisGayofa")]
        public IActionResult FotoebisGayofa()
        {
            //var result = _windbreak.GetCheckPhotoDate(checkPhotoDateDTO.folderPath, checkPhotoDateDTO.resultPath);
            //if (result.Success) return Ok(result);
            GeographicDynamicDbContext windBreakContext = new GeographicDynamicDbContext();
            try
            {

                    GadanomriliFotoebi photo = new GadanomriliFotoebi();

                    var directories = Directory.GetDirectories(@"D:\\Documents\\Desktop\\I_etapi\\Photoes").OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                    foreach (var folderPath in directories)
                    {
                        var idxLiter = folderPath.LastIndexOf('\\');
                        string literIDstr = folderPath.Substring(idxLiter + 1);

                        double literID = Convert.ToDouble(literIDstr);

                        var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                        var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));


                        foreach (var item in list)
                        {
                            DirectoryInfo d5 = new DirectoryInfo(item);
                            FileInfo[] infos1 = d5.GetFiles();

                            var idxUniqid = item.LastIndexOf('\\');

                            string uniqIDstr = item.Substring(idxUniqid + 1);

                            double uniqID = Convert.ToDouble(uniqIDstr);

                            string photoN = "";

                            var PhotoDate = "";

                            //photo.UniqId = uniqID;
                            bool ismoved = true;
                            foreach (FileInfo f6 in infos1)
                            {
                                if (!f6.Name.Contains(".db"))
                                {

                                    QarsafariGrouped? qarsafaritest = windBreakContext.QarsafariGroupeds.FirstOrDefault(m => m.UniqId == uniqID);

     
                                    if (qarsafaritest?.Sakutreba == "კერძო" || qarsafaritest?.Sakutreba == "იურიდიული პირი")
                                    {
                                        if (ismoved)
                                        {

                                            photo.LiterId = literID;
                                            string destinationFolder = Path.Combine((string.Concat(@"D:\Documents\Desktop\I_etapi\Split" + "\\" + "photoSplit" + "\\" + "Kerdzo")), literID.ToString());
                                            if (!Directory.Exists(destinationFolder))
                                            {
                                                Directory.CreateDirectory(destinationFolder);
                                            }
                                            string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                            //File.Copy(item, destinationFile);
                                            Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);

                                            ismoved = false;
                                        }

                                    }
                                    if (qarsafaritest?.Sakutreba != "კერძო" && qarsafaritest?.Sakutreba != "იურიდიული პირი")
                                    {
                                        if (ismoved)
                                        {
                                            photo.LiterId = literID;


                                            string destinationFolder = Path.Combine((string.Concat(@"D:\Documents\Desktop\I_etapi\Split" + "\\" + "photoSplit" + "\\" + "Saxelmwifo")), literID.ToString());

                                            if (!Directory.Exists(destinationFolder))
                                            {
                                                Directory.CreateDirectory(destinationFolder);
                                            }
                                            string destinationFile = Path.Combine(destinationFolder, uniqID.ToString());
                                            //File.Copy(item, destinationFile);
                                            Microsoft.VisualBasic.FileIO.FileSystem.CopyDirectory(item, destinationFile);
                                            ismoved = false;
                                        }
                                    }
                                }
                            }
                        }

                    }
                
            }

            catch (Exception ex)
            {
                return BadRequest("false");
            }


            return BadRequest("false");
        }

    }
}
