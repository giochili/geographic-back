using Azure;
using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using GeographicDynamic_DAL.Models;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Internal;
using Microsoft.EntityFrameworkCore.Query.Internal;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static Azure.Core.HttpHeader;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory;
using GeographicDynamic_DAL.Models.WindbreakMethods;

namespace GeographicDynamic_DAL.Repository
{
    public class WindbreakRepository : IWindbreak
    {
        WindbreakMethods _windbreakMethods = new WindbreakMethods();

        public Result<bool> GetCheckPhotoDate(string folderPath, string resultPath)
        {
            List<string> UnMachedPhotos = new List<string>();
            try
            {
                string ForExcelName = Path.GetFileName(resultPath);
                var directories = Directory.GetDirectories(folderPath);

                //for proggress bar 
                int currentDirectoryIndex = 0;
                foreach (var Liters in directories)
                {

                    var actualPathOfFolderLitter = Liters.LastIndexOf('\\');
                    var numericalPathLitter = Liters.Substring(actualPathOfFolderLitter + 1);
                    var litterPath = Directory.GetDirectories(Liters);
                    int totalDirectories = directories.Length;
                    currentDirectoryIndex++;

                    foreach (var d in litterPath)
                    {

                        var actualPathOfFolder = d.LastIndexOf('\\');
                        var numericalPath = d.Substring(actualPathOfFolder + 1);

                        //var directoriesWithoutExtention = Directory.GetDirectories(d).OrderBy(x => Convert.ToInt32(Path.GetFileNameWithoutExtension(x)));
                        var directoriesWithoutExtention = Directory.GetFiles(d);

                        //var list = directoriesWithoutExtention.OrderBy(y => int.Parse(y.Split('\\')[5])).ToList();



                        //foreach (var item in list)
                        //{

                        DirectoryInfo d5 = new DirectoryInfo(d);
                        FileInfo[] infos1 = d5.GetFiles();
                        var innerdir = d;
                        var files = Directory.GetFiles(d).Where(m => !m.Contains(".db") && (m.Contains(".jpg") || m.Contains(".jpeg"))).ToList();

                        int photoLength = 0;
                        bool photoebiaremtxveva = false;
                        foreach (var file in files)
                        {

                            //DateTime firstFileDate = File.GetCreationTime(files[0]); // შექმნის თარიღი (გადაკოპირების)    
                            DateTime firstFileDate = File.GetLastWriteTime(files[0]); // მოდიფიკაციის თარიღი( გადაღების) 
                            string formattedDate = firstFileDate.ToString("MM/dd/yy");

                            photoLength++;
                            var extensionTest = Path.GetExtension(file);

                            //DateTime fileDate = File.GetCreationTime(file);
                            DateTime fileDate = File.GetLastWriteTime(file);


                            string formattedDateToCompare = fileDate.ToString("MM/dd/yy");

                            if (formattedDateToCompare != formattedDate && !photoebiaremtxveva)
                            {
                                UnMachedPhotos.Add(numericalPathLitter + " | " + "/" + numericalPath.ToString() + "/" + " |თარიღები არ ემთხვევა");
                                photoebiaremtxveva = true;
                                // MessageBox.Show("მოხდა შეცდომა ფაილების თარიღები არ ემთხვევა ერთმანეთს !");
                            }


                            //}

                        }
                        if (photoLength < 3)
                        {
                            UnMachedPhotos.Add(numericalPathLitter + " | " + "/" + numericalPath.ToString() + "/" + " |ნაკლები ფოტოა");
                            //MessageBox.Show("მოხდა შეცდომა ფოლდერში 3-ზე ნაკლები ფოტოს ფაილია აღმოჩენილი !");
                        }
                        // აქ ემატება შემოწმება იმისთვის რომ ინახოს თუ არის 10 ზე მეტი ფოტო ფილდერში და თუ კი მაშინ წაშალოს იმდენი რამდენითაც მეტია 10 ზე 
                        // წასაშლელად ირჩევა შემთხვევითი პრინციპით რიცხვი მასივში და შენდეგ იშლება
                        if (photoLength > 10)
                        {
                            Random random = new Random();
                            int filesToDeleteCount = photoLength - 10; // Calculate the number of files to delete

                            for (int i = 0; i < filesToDeleteCount; i++)
                            {
                                // Select a random file from the list
                                int randomIndex = random.Next(files.Count);
                                string fileToDelete = files[randomIndex];

                                // Delete the file
                                File.Delete(fileToDelete);

                                // Remove the file from the list
                                files.RemoveAt(randomIndex);
                            }

                            // Update photoLength after deletion
                            photoLength = files.Count;
                        }


                    }
                }

                WriteToExcel(UnMachedPhotos, ForExcelName);

                void WriteToExcel(List<string> UnMatchedPhotos, string ForExcelName)
                {
                    Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    Workbook ExcelWorkBook = null;
                    Worksheet ExcelWorkSheet = null;

                    // Set Excel application to not be visible
                    ExcelApp.Visible = true;


                    ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);

                    ExcelWorkBook.Worksheets.Add(); //Adding New Sheet in Excel Workbook

                    try
                    {
                        ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // Compulsory Line in which sheet you want to write data
                                                                      //Writing data into excel of 100 rows with 10 column 
                        ExcelWorkSheet.Cells[1, "A"] = "შეცდომები";
                        //ExcelWorkSheet.Cells[1, "B"] = "UNIQ_ID";
                        for (int r = 0; r < UnMatchedPhotos.Count(); r++) //r stands for ExcelRow and c for ExcelColumn
                        {
                            string[] parts = UnMatchedPhotos[r].Split('/');
                            ExcelWorkSheet.Cells[r + 2, "A"] = string.Concat(parts);
                            //ExcelWorkSheet.Cells[r + 2, "B"] = parts[1];

                        }
                        ExcelWorkBook.Worksheets[1].Name = "ResultSheet";//Renaming the Sheet1 to MySheet
                        ExcelWorkBook.SaveAs(resultPath + "\\Results-" + ForExcelName + ".xlsx");
                        // ExcelWorkBook.Close();
                        // ExcelApp.Quit();
                        Marshal.ReleaseComObject(ExcelWorkSheet);
                        Marshal.ReleaseComObject(ExcelWorkBook);
                        Marshal.ReleaseComObject(ExcelApp);

                        //Process.Start(resultPath + "\\Results-" + ForExcelName + ".xlsx");
                        //Process.Start(new ProcessStartInfo { FileName = @"${resultPath}\\Results-{ForExcelName}.xlsx", UseShellExecute = true });
                    }

                    catch (Exception exHandle)

                    {

                        Console.WriteLine("Exception: " + exHandle.Message);

                        Console.ReadLine();

                    }
                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "მოხდა შეცდომა:" + ex.Message
                };
            }
        }


        //ფოტოების გაყოფის ფუნცქია 
        public Result<bool> PhotoSplitKerdzoSaxelmwifo(string GadanomriliPhotoFolderPath, string DestinationFolderPath)
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
            try
            {
                GadanomriliFotoebi photo = new GadanomriliFotoebi();
                var directories = Directory.GetDirectories(GadanomriliPhotoFolderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

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

                        photo.UniqId = uniqIDstr;
                        bool ismoved = true;
                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {

                                QarsafariGrouped? qarsafaritest = geographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(m => m.UniqId == uniqID);

                                // აქ გვჭირდება რომ მოწმდებოდეს მარტო კერძო ან სახელმწიფო რადგან ბაზაში იურიდიული პირიდა მუნიცპალიტეტი აღარაა მარტო კერძო ან სახელმწიფო
                                if (qarsafaritest?.Owner == "კერძო" || qarsafaritest?.Owner == "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {

                                        photo.LiterId = literID;
                                        string destinationFolder = Path.Combine((string.Concat(DestinationFolderPath + "\\" + "photoSplit" + "\\" + "Kerdzo")), literID.ToString());
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
                                if (qarsafaritest?.Owner != "კერძო" && qarsafaritest?.Owner != "იურიდიული პირი")
                                {
                                    if (ismoved)
                                    {
                                        photo.LiterId = literID;


                                        string destinationFolder = Path.Combine((string.Concat(DestinationFolderPath + "\\" + "photoSplit" + "\\" + "Saxelmwifo")), literID.ToString());

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
            catch(Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "მოხდა შეცდომა:" + ex.Message
                };
            }





            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }



        //ფოტოების გადანომრვის ფუნქციონალი ეშვება აქ 
        public Result<bool> RenamePhotosInFolder(RenamePhotoDTO renamePhotoDTO)
        {

            #region GIORGI
            //var directories = Directory.GetDirectories(renamePhotoDTO.FolderPath).OrderBy(filePath => Convert.ToString(Path.GetFileNameWithoutExtension(filePath)));
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            geographicDynamicDbContext.GadanomriliFotoebis.ExecuteDelete();
            var directories = Directory.GetDirectories(renamePhotoDTO.FolderPath).OrderBy(filePath => int.Parse(Path.GetFileNameWithoutExtension(filePath)));

            int foldercount = renamePhotoDTO.FolderStartNumber;
            int photocount = renamePhotoDTO.PhotoStartNumber;
            var random = new Random();
            var tempFolderCount = random.Next(100000, 999999);
            if (renamePhotoDTO.Gadanomrilia == false)
            {

                ////test for _ 

                //foreach (var directory in directories)
                //{
                //    var directories1 = Directory.GetDirectories(directory).OrderBy(filePath => Convert.ToString(Path.GetFileNameWithoutExtension(filePath)));
                //    var list = directories1.OrderBy(filePath => Convert.ToString(Path.GetFileNameWithoutExtension(filePath)));
                //    foreach (var item in list)
                //    {
                //        int idx2 = item.LastIndexOf('\\');
                //        var kk2 = Convert.ToString(item.Substring(idx2 + 1));
                //        if (kk2.Contains("_"))
                //        {
                //           var splited = kk2.Split("_");

                //            if (splited.Length == 2 && int.TryParse(splited[0], out int start) && int.TryParse(splited[1], out int end))
                //            {
                //                for (int i = start; i <= end; i++)
                //                {
                //                    string newFolderPath = Path.Combine(directory, $"{i}");
                //                    if (!Directory.Exists(newFolderPath))
                //                    {
                //                        // Create the new folder
                //                        Directory.CreateDirectory(newFolderPath);
                //                        Console.WriteLine("Folder created successfully.");

                //                        string sourceDirectory = Path.Combine(directory, kk2); 
                //                        string[] imageFiles = Directory.GetFiles(sourceDirectory);
                //                        foreach (string imageFile in imageFiles)
                //                        {

                //                            string fileName = Path.GetFileName(imageFile);
                //                            string destinationFile = Path.Combine(newFolderPath, fileName);
                //                            File.Copy(imageFile, destinationFile, true);
                //                            Console.WriteLine($"Copied {fileName} to {newFolderPath}");
                //                        }

                //                    }
                //                }
                //                Directory.Delete(Path.Combine(directory, kk2), true);
                //            }

                //        }

                //    }

                //}
                ////Test end 





                try
                {
                    //ფაილების გადანომვრა 
                    foreach (var folderPath in directories)
                    {                 //ფაილების გადანომვრა რენდომ რიცხვით რომ გამოირიცხოს დუპლიკატი

                        var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));
                        var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));
                        foreach (var items in list)
                        {
                            var idx = items.LastIndexOf('\\');
                            string kk = items.Substring(idx + 1);


                            var newname = items.Replace(kk, Convert.ToString(foldercount)); //es mushaobs

                            int idx11 = items.LastIndexOf('\\');
                            string oldfoldername = items.Substring(0, idx11);
                            string newnamefolder = oldfoldername + "\\" + tempFolderCount; //ჯერ ეს უნდა გავუშვათ 

                            Directory.Move(items, newnamefolder);
                            tempFolderCount++;

                        }
                    }

                    //გადანომვრის ციკლი შერჩეული რიცხვით სადანაც გვინდა დაიწყოს 
                    foreach (var folderPath in directories)
                    {
                        int idx2 = folderPath.LastIndexOf('\\');
                        var kk2 = Convert.ToInt32(folderPath.Substring(idx2 + 1));

                        var directories1 = Directory.GetDirectories(folderPath).OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                        var list = directories1.OrderBy(filePath => Convert.ToInt32(Path.GetFileNameWithoutExtension(filePath)));

                        foreach (var items in list)
                        {
                            var idx = items.LastIndexOf('\\');
                            string kk = items.Substring(idx + 1);


                            var newname = items.Replace(kk, Convert.ToString(foldercount)); //es mushaobs

                            int idx11 = items.LastIndexOf('\\');
                            string oldfoldername = items.Substring(0, idx11);
                            string newnamefolder = oldfoldername + "\\" + Convert.ToString(foldercount);
                            Directory.Move(items, newnamefolder);

                            foldercount++;

                        }
                    }

                }
                catch (Exception ex) { }
            }
            //ფოტოების გადანომვრა 
            try
            {

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

                        string photoN = "";

                        var PhotoDate = "";
                        string photoNCorrected = "";
                        // ფოტოების გადასანომრი ციკლი 
                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {
                                var ext = Path.GetExtension(f6.FullName);
                                File.Move(f6.FullName, f6.FullName.Replace(f6.Name, Convert.ToString(photocount) + ext));


                                photoN += Convert.ToString(photocount) + "/";



                                //ფოტოს თარიღის წამოღება
                                bool isWritten = false;
                                if (!isWritten)
                                {
                                    var modifiedDate1 = f6.LastWriteTime;
                                    var formatInfo = new CultureInfo("en-US").DateTimeFormat;
                                    formatInfo.DateSeparator = "-";
                                    PhotoDate = modifiedDate1.ToString("dd-MM-yyyy", formatInfo);
                                }
                                isWritten = true;
                                photocount++;

                            }

                        }
                        // SQL ბაზაში დამატება და ცვლილებების დამახსოვრება 
                        GadanomriliFotoebi photo = new GadanomriliFotoebi();
                        photo.UniqId = uniqIDstr;
                        photo.LiterId = literID;
                        photoNCorrected = photoN.TrimEnd('/'); // ბოლოში სლექშებს უშლის 
                        photo.PhotoN = photoNCorrected;
                        photo.PhotoDate = PhotoDate;
                        geographicDynamicDbContext.GadanomriliFotoebis.Add(photo); // ამით ემატება ბაზაში 
                        geographicDynamicDbContext.SaveChanges();//ამით ამახსოვრებს 
                    }
                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK, Message = "წარმატებით გადაინომრა" };
            }
            catch (Exception ex)
            {
                return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.OK, Message = "ჩტო ტა ნიტო" };
            }
            #endregion


            return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
        }



        //ექსელის ფაილებთან სამუშაო ფუნქცია სადაც უნდა წამოვიდეს უშუალოდ ექსელის წაკითხვისას 
        public Result<bool> ExcelCalculations(ExcelReadDTO excelReadDTO)

        {

            ////////// ჩეკბოქსების გამოტანა ცვლადში რომ ადვილად აიწყოს შემოწმება 
            var CalcVarjisFartiCheckbox = excelReadDTO.CalcVarjisFartiCheckbox;
            var AccessShitNameTextbox = excelReadDTO.AccessShitName;


            /////////////ფუნქციების გამოძახებები თავის შეცდომიანად თუ სადმე რამე იყო 
            ////// ვკითხულობთ ექსელიდან ინფორმაციას და შეგვაქვს sql ში qarsafari ცხრილი
            var ExcelisWakitxvaRestult = _windbreakMethods.ExcelisWakitxva(excelReadDTO);
            if (ExcelisWakitxvaRestult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელის წაკითხვის დროს" + ExcelisWakitxvaRestult.Message
                };
            }
            ////ვამოწმებთ Excel-ში თუ არის დუპლიკატი Unic-Liter-ID -ები
            //var ShemowmebaUnicLiterExcelshiResult = _windbreakMethods.ShemowmebaUnicLiterExcelshi();
            //if (ShemowmebaUnicLiterExcelshiResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = "მოხდა შეცდომა Access ფაილის წაკითხვისას" + ShemowmebaUnicLiterExcelshiResult.Message
            //    };
            //}

            //// ვკითხულობთ აქსესიდან ინფორმაციას და შეგვაქვს sql ში WindbreakMDB ცხრილი
            //if (AccessShitNameTextbox != "")
            //{
            //    var AccessReadingResult = _windbreakMethods.AccessWakitxva(excelReadDTO.AccessFilePath, excelReadDTO.AccessShitName);

            //    if (AccessReadingResult.Success == false)
            //    {
            //        return new Result<bool>
            //        {
            //            Success = false,
            //            StatusCode = System.Net.HttpStatusCode.BadGateway,
            //            Message = "მოხდა შეცდომა Access ფაილის წაკითხვისას" + AccessReadingResult.Message
            //        };
            //    }
            //}

            /////////////////ეს ფუქნცია ამოწმებს excel და access ცხრილებს და ადარებს UNIQID ებს თუ ემთხვევა ერთმანეთს
            //var ShemowmebaAccessExcelUnicLiterDublicatsResult = _windbreakMethods.ShemowmebaAccessExcelUnicLiterDublicats();
            //if (ShemowmebaAccessExcelUnicLiterDublicatsResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = ShemowmebaAccessExcelUnicLiterDublicatsResult.Message + " ზედმეტია: " + ShemowmebaAccessExcelUnicLiterDublicatsResult.Data
            //    };
            //}

            //if (CalcVarjisFartiCheckbox == true)
            //{
            //    var ChaweraVarjisPartiFunction = _windbreakMethods.ChaweraVarjisParti(excelReadDTO.ProjectNameID);

            //    if (ChaweraVarjisPartiFunction.Success == false)
            //    {
            //        return new Result<bool>
            //        {
            //            Success = false,
            //            StatusCode = System.Net.HttpStatusCode.BadGateway,
            //            Message = "ვარჯის ფართის დათვლის დროს მოხდა შეცდომა" + ChaweraVarjisPartiFunction.Message
            //        };
            //    }
            //}
            //var CheckerOfVarjisFartiandSaxeobaResult = _windbreakMethods.CheckerOfVarjisFartiandSaxeoba();
            //if (CheckerOfVarjisFartiandSaxeobaResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = "მოხდა შეცდომა ვარჯის ფართის ჩაწერისას " + CheckerOfVarjisFartiandSaxeobaResult.Message
            //    };
            //}
            //var UIDReplaceExcelResult = _windbreakMethods.UIDReplaceExcel();
            //if (UIDReplaceExcelResult.Success == false)
            //{
            //    return new Result<bool>
            //    {
            //        Success = false,
            //        StatusCode = System.Net.HttpStatusCode.BadGateway,
            //        Message = "მოხდა შეცდომა ექსელში UID-ის ჩაწერის დროს" + UIDReplaceExcelResult.Message
            //    };
            //}

            var QarsafariGadanomrvaResult = _windbreakMethods.QarsafariGadanomrva(excelReadDTO.UnicIDStartNumber);
            if (QarsafariGadanomrvaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariGadanomrvaResult.Message
                };
            }

            // აქ უკვე გავაკეთოთ თუ დაჩეკილია ფოტოების გადანომვრა დეფოლტად იყოს დაჩეკილი
            // თუ დაჩეკილია გადავნომრავთ
            //ექსელის ახალი უნიკიდი უნდა ჩავწეროთ ფოლდერის  პაპკის უნიკიდიში
            // და ჯოინი ხდება არსებული ფოლდერის უნიკიდი = ექსელის ძველი უნიკიდი
            // აქ უნდა შევინარჩუნოთ ლოგიკა რომ ყველას ვუწერთ 00000 დუბლიკატები რომ გამოვრიცხოთ

            //ექსელის ძველი უნიკიდებით ვეძებთ ფოლდერებში უნიკიდიებს თავისთავად რომელ ლიტერშია და ვუწერთ ექსელის ახალ უნიკიდს
            // და თუ ძველ უნიკიდის ვერ ვიპოვით ვერცერთ ლიტერში მაშინ ERROR ! და აღარ უნდა გააგრძელოს და გამოიტანოს ის ლიტერი უნიკიდი რაც ვერ იპოვა ფოტოებში


            RenamePhotoDTO renamePhotoDTOTEST = new RenamePhotoDTO();
            renamePhotoDTOTEST.FolderPath = excelReadDTO.FolderPath;


            var GadanomvraFolderebisExcelidanResult = RenamePhotosInFolder(renamePhotoDTOTEST);
            //var GadanomvraFolderebisExcelidanResult = _windbreakMethods.GadanomvraFolderebisExcelidan(RenamePhotoDTO renamePhotoDTO);
            if (GadanomvraFolderebisExcelidanResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = GadanomvraFolderebisExcelidanResult.Message
                };
            }
               

            var ProcentisDatvlaResult = _windbreakMethods.QarsafariProcentisDatvla();
            if (ProcentisDatvlaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = ProcentisDatvlaResult.Message
                };
            }

            // გადაგვაქვს ინფორმაცია აქსესიდან ექსელში
            var UpdateFromAccessToExcellResult = _windbreakMethods.UpdateFromAccessToExcell();
            if (UpdateFromAccessToExcellResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UpdateFromAccessToExcellResult.Message
                };
            }
            //საკუთრებაში ვწერთ სახელმწიფოა თუ კერძო
            var FillSakutrebaIsKerdzoOrSaxelmwifoResult = FillSakutrebaIsKerdzoOrSaxelmwifo();
            if (FillSakutrebaIsKerdzoOrSaxelmwifoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა საკუთრების მინიჭების დროს "
                };
            }

            var UIDReplaceAccessResult = _windbreakMethods.UIDReplaceAccess();
            if (UIDReplaceAccessResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UIDReplaceAccessResult.Message
                };
            }


            ////////////axali funqcia UIDREPLACE () {} // table qarsafarshi
            ////////////SET UID = str([ლიტერი ID]) + str([უნიკ ID]) // str chventan aris Convert.ToString()
            ///////////// ჯერ არ ვიყიენებთ მარა გამოსაყენებელია ხეხილში ვამოწმებთ დუბლიკატები ხომ არ არის
            var QarsafariXexilisShemowmebaResult = _windbreakMethods.QarsafariXexilisShemowmeba();
            if (QarsafariXexilisShemowmebaResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariXexilisShemowmebaResult.Message
                };
            }



            // qarsafari ცხრილის დაგრუპვა uniqid ის მიხედვით და გადატანა qarsafariGrouped ში
            var QarsafariToQarsafariGroupedResult = _windbreakMethods.QarsafariToQarsafariGrouped();
            if (QarsafariToQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = QarsafariToQarsafariGroupedResult.Message
                };
            }



            var UIDReplaceQarsafariGroupedResult = _windbreakMethods.UIDReplaceQarsafariGrouped();
            if (UIDReplaceQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UIDReplaceQarsafariGroupedResult.Message
                };
            }


            var GadanomriliFotoebiToQarsafariGroupedResult = _windbreakMethods.GadanomriliFotoebiToQarsafariGrouped();
            if (GadanomriliFotoebiToQarsafariGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = GadanomriliFotoebiToQarsafariGroupedResult.Message
                };
            }


            var UPDTFromExcelToAccessResult = _windbreakMethods.UPDTFromExcelToAccess(excelReadDTO.AccessShitName);
            if (UPDTFromExcelToAccessResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = UPDTFromExcelToAccessResult.Message
                };
            }
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            /////////////// ქარსაფარი გრუპდის ცხრილები რომ ამოექსპორტდეს 
            List<QarsafariGrouped> qarsafariGroupeds = geographicDynamicDbContext.QarsafariGroupeds.OrderBy(m => m.UniqId).ToList();
            List<QarsafariGrouped> qarsafariGroupedsSaxelmwifo = geographicDynamicDbContext.QarsafariGroupeds.Where(x => x.Sakutreba == "სახელმწიფო" || x.Sakutreba == "მუნიციპალიტეტი").OrderBy(m => m.UniqId).ToList();
            List<QarsafariGrouped> qarsafariGroupedsKerdzo = geographicDynamicDbContext.QarsafariGroupeds.Where(x => x.Sakutreba == "კერძო" || x.Sakutreba == "იურიდიული პირი").OrderBy(m => m.UniqId).ToList();

            ///////////////ფუნქციის გამოძახებები

            var WriteToExcelGroupedResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupeds, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelGroupedResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelGroupedResult.Message,
                };
            }

            var WriteToExcelGroupedSaxelmwifoResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupedsSaxelmwifo, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-Saxelmwifo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelGroupedSaxelmwifoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelGroupedSaxelmwifoResult.Message,
                };
            }
            var WriteToExcelGroupedKerdzoResult = _windbreakMethods.WriteToExcelGrouped(qarsafariGroupedsKerdzo, excelReadDTO.ExcelDestinationPath, "QarsafariGrouped-Kerdzo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelGroupedKerdzoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelGroupedKerdzoResult.Message,
                };
            }

            ///////////////////ქარსაფარის ცხრილები რომ ამოექსპორტდეს 

            List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).ToList();//მთლიანი ცხრილი 
            List<Qarsafari> qarsafarisSaxelmwifo = geographicDynamicDbContext.Qarsafaris.Where(x => x.Sakutreba == "სახელმწიფო" || x.Sakutreba == "მუნიციპალიტეტი").ToList();//სახელმწიფო საკუთრების ცხრილი 
            List<Qarsafari> qarsafarisKerdzo = geographicDynamicDbContext.Qarsafaris.Where(x => x.Sakutreba == "კერძო" || x.Sakutreba == "იურიდიული პირი").ToList();// კერძო საკუთრების ცხრილი 
            //////////////////// ფუნქციის გამოძახებები 

            var WriteToExcelQarsafariResult = _windbreakMethods.WriteToExcel(qarsafaris, excelReadDTO.ExcelDestinationPath, "Qarsafari-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelQarsafariResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelQarsafariResult.Message,
                };
            }
            var WriteToExcelQarsafariSaxelmwifoResult = _windbreakMethods.WriteToExcel(qarsafarisSaxelmwifo, excelReadDTO.ExcelDestinationPath, "Qarsafari-Saxelmwifo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelQarsafariSaxelmwifoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelQarsafariSaxelmwifoResult.Message,
                };
            }
            var WriteToExcelQarsafariKerdzoResult = _windbreakMethods.WriteToExcel(qarsafarisKerdzo, excelReadDTO.ExcelDestinationPath, "Qarsafari-Kerdzo-" + DateTime.Now.ToString("yyyy-MM-dd"));
            if (WriteToExcelQarsafariKerdzoResult.Success == false)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = WriteToExcelQarsafariKerdzoResult.Message,
                };
            }


            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };

        }



        // აქ ხდება ფუნქციების აღწერა 



        // გიოს ნახლაფორთარი რომელიც ერთ ხაზში დაიწერა საკუთრების ველის შევსებისთვის 
        #region ნახლაფორთარი გიოსი 
        public Result<bool> FillSakutrebaIsKerdzoOrSaxelmwifo() // ფუნქცია გამოიყენება იმისთვის რომ საკუთრების ველში შეივსოს სახელწმიფოა თუ კერძო 
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            try
            {
                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();

                string Sakutrebastore = "";// ინახება მნიშვნელობა როცა isUniqIdLiterIdtrue 
                foreach (var item in qarsafaris)
                {
                    if (item.IsUniqLiterNull == "true" && item.Owner != null)
                    {
                        if (item.Owner == "მუნიციპალიტეტი" || item.Owner == "სახელმწიფო" || String.IsNullOrEmpty(item.Owner))
                        {
                            item.Sakutreba = "სახელმწიფო";
                            Sakutrebastore = "სახელმწიფო";
                        }
                        else
                        {
                            item.Sakutreba = "კერძო";
                            Sakutrebastore = "კერძო";
                        }
                    }

                    if (item.IsUniqLiterNull == "false" && item.Owner == null)
                    {
                        item.Sakutreba = Sakutrebastore;
                    }


                    geographicDynamicDbContext.SaveChanges();
                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა IsUniqLitterNull ველის შევსების დროს" + ex.Message
                };

            }
        }
        #endregion






        // ეშვება GET ფუნქცია რომ მიიღოს მუნიციპალიტეტების სია FRONT-ში 
        public Result<DictionaryDTO> GetProjectNames()
        {
            try
            {

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                var list = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 2).ToList();

                List<DictionaryDTO> DictionaryDTOs = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 2).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name }).ToList();

                return new Result<DictionaryDTO>
                {
                    Success = true,
                    Data = DictionaryDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<DictionaryDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };

            }


        }
        //ეშვება GET გუნქცია რომ მიიღოს ეტაპის სია FRont-ში
        public Result<DictionaryDTO> GetEtapiID()
        {
            try
            {

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                var list = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 3).ToList();

                List<DictionaryDTO> DictionaryDTOs = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 3).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name }).ToList();


                return new Result<DictionaryDTO>
                {
                    Success = true,
                    Data = DictionaryDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა  ეტაპების სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<DictionaryDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };

            }
        }

        // ეშვება GET  ფუნქცია რომ წამოიღოს ვარჯის ფართები მუნიციპალიტეტის მიხედვით ,
        public Result<VarjisFartiDTO> GetVarjisFartebi(int AreaNameID)
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
                List<VarjisFartiDTO> VarjisFartiDTOs = geographicDynamicDbContext.VarjisFartis.Where(x => x.AreaNameId == AreaNameID).Select(x => new VarjisFartiDTO { Id = x.Id, Name = x.Saxeoba.Name, SaxeobaId = x.SaxeobaId, AreaNameId = x.AreaNameId, VarjisFarti1 = x.VarjisFarti1 }).ToList();
                return new Result<VarjisFartiDTO>
                {
                    Success = true,
                    Data = VarjisFartiDTOs,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დაბრუნდა სია"
                };
            }
            catch (Exception ex)
            {
                return new Result<VarjisFartiDTO>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "შეცდომა მოხდა " + ex.Message
                };
            }
        }

        public Result<DictionaryDTO> getSaxeobaList()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();



            List<DictionaryDTO> saxeobaList = geographicDynamicDbContext.Dictionaries.Where(x => x.Code == 1).Select(x => new DictionaryDTO { ID = x.Id, Name = x.Name, Code = x.Code }).OrderBy(x => x.Name).ToList();
            //saxeobaList.OrderByDescending(x=>x.Name);

            return new Result<DictionaryDTO>
            {
                Success = true,
                Data = saxeobaList,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }



        //public Result<bool> GroupingFormulas()
        //{
        //    GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
        //    List<Qarsafari> qarsafari = geographicDynamicDbContext.Qarsafaris.ToList();
        //    List<QarsafariGrouped> qarsafariGrouped = geographicDynamicDbContext.QarsafariGroupeds.ToList();
        //    try
        //    {
        //        foreach( var item in qarsafari)
        //        {

        //        }

        //        return new Result<bool>
        //        {
        //            Success = true,
        //            StatusCode = System.Net.HttpStatusCode.OK
        //        };
        //    }
        //    catch (Exception ex)
        //    {
        //        return new Result<bool>
        //        {
        //            Success = false,
        //            StatusCode = System.Net.HttpStatusCode.BadGateway,
        //            Message = "მოხდა შეცდომა დაგრუპვის ფორმულის დროს" + ex.Message
        //        };
        //    }
        //}


        // ფუნცქია კითხულობს ბაზას და ქმნის ახალ ექსელის ფაილს რომ შევიდეს მონაცემები შიგნით და გაკეთდეს ექსელის დიდი ფაილი მხოლოდ ქარსაფარისთვის 



    }



}
