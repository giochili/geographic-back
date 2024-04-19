﻿using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamicWebAPI.Wrappers;
using Microsoft.EntityFrameworkCore;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace GeographicDynamic_DAL.Models.WindbreakMethods
{
    public class WindbreakMethods
    {
        #region Methods
        public Result<bool> FillIsUniqLitterNull()
        {
            try
            {
                GeographicDynamicDbContext GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.ToList();

                foreach (var item in qarsafaris)
                {
                    if ((item.UniqId == null || item.LiterId == null) || (item.UniqId == 0 || item.LiterId == 0))
                    {
                        item.IsUniqLiterNull = "false";

                    }
                    if (item.UniqId != null && item.LiterId != null)
                    {
                        item.IsUniqLiterNull = "true";
                    }



                    GeographicDynamicDbContext.SaveChanges();

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
        public Result<bool> ExcelisWakitxva(ExcelReadDTO excelReadDTO)
        {
            GeographicDynamicDbContext GeographicDynamicDbContext = new GeographicDynamicDbContext();
            var test = excelReadDTO.UnicIDStartNumber;
            var test1 = excelReadDTO.ExcelDestinationPath;
            var ExcelPath = excelReadDTO.ExcelPath;
            var test3 = excelReadDTO.AccessFilePath;
            var municipality = excelReadDTO.ProjectNameID;
            Application xlApp = new Application();
            try
            {
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
                //Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1 ];
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = (Microsoft.Office.Interop.Excel._Worksheet)xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                //label2.Visible = true;
                Microsoft.Office.Interop.Excel.Range lastCell = xlWorksheet.Cells.SpecialCells(Microsoft.Office.Interop.Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                int colCount = lastCell.Row;
                //ExcelProgressBar.Minimum = 1;
                //ExcelProgressBar.Maximum = colCount - 1;
                //ExcelProgressBar.Step = 1;

                //if (colCount == ExcelProgressBar.Maximum)
                //{
                //    ExcelProgressBar.Maximum = 100;
                //}




                // წინასწარ ცხრილის გასუფთავება მანამ ჩანაწერებს შევიტანთ

                // ----აქეეედააააან
                GeographicDynamicDbContext.Qarsafaris.ExecuteDelete();
                Type myType = typeof(Qarsafari);
                //იტერაცია ექსელის ფაილში
                for (int i = 2; i <= colCount; i++) // colCount tu sworad wakikitxavs
                {
                    Qarsafari qarsafari = new Qarsafari();

                    // თუ უნიკიდ ან ლიტერ აიდი ცარიელია მაშინ ჩაიწერება false თუ არაა ცარიელი მაშინ true
                    if (String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, "A"].Value2)) || String.IsNullOrEmpty(Convert.ToString(xlRange.Cells[i, "B"].Value2)))
                    {
                        qarsafari.IsUniqLiterNull = "false";

                    }
                    else
                    {
                        qarsafari.IsUniqLiterNull = "true";
                    }


                    foreach (var columnName in GeographicDynamicDbContext.ColumnNames)
                    {
                        //Get cell type
                        if (columnName.ColN != null)
                        {


                            object cellValue = xlRange.Cells[i, columnName.ColN].Value2;
                            if (cellValue != null)
                            {


                                Type cellType = cellValue.GetType();
                                PropertyInfo propertyInfo = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                                if (propertyInfo != null)
                                {

                                    // Handle conversion based on cell type
                                    if (cellType == typeof(double))
                                    {
                                        propertyInfo.SetValue(qarsafari, (double)cellValue);
                                    }
                                    else if (cellType == typeof(string))
                                    {
                                        propertyInfo.SetValue(qarsafari, cellValue);
                                    }
                                    else if (cellType == typeof(DateTime))
                                    {
                                        DateTime dateTimeValue;
                                        if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                                        {
                                            propertyInfo.SetValue(qarsafari, dateTimeValue);
                                        }
                                        // Handle DateTime conversion if necessary
                                    }
                                    // Add other type conversions as necessary
                                }
                                
                            }
                            //else
                            //{
                            //    cellValue = (int)cellValue + 1;
                            //    Type cellType = cellValue.GetType();
                            //    PropertyInfo propertyInfo = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                            //    if (propertyInfo != null)
                            //    {

                            //        // Handle conversion based on cell type
                            //        if (cellType == typeof(double))
                            //        {
                            //            propertyInfo.SetValue(qarsafari, (double)cellValue);
                            //        }
                            //        else if (cellType == typeof(string))
                            //        {
                            //            propertyInfo.SetValue(qarsafari, cellValue);
                            //        }
                            //        else if (cellType == typeof(DateTime))
                            //        {
                            //            DateTime dateTimeValue;
                            //            if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                            //            {
                            //                propertyInfo.SetValue(qarsafari, dateTimeValue);
                            //            }
                            //            // Handle DateTime conversion if necessary
                            //        }
                            //        // Add other type conversions as necessary
                            //    }
                            //}
                        }
                    }
                    //qarsafari.UniqId = Convert.ToDouble(xlRange.Cells[i, "A"].Value2);
                    //qarsafari.LiterId = Convert.ToDouble(xlRange.Cells[i, "B"].Value2);
                    //qarsafari.PhotoN = xlRange.Cells[i, "C"].Value2;
                    //qarsafari.Region = xlRange.Cells[i, "D"].value2;
                    //qarsafari.Municipality = xlRange.Cells[i, "E"].Value2;
                    //qarsafari.Shrubbery = Convert.ToDouble(xlRange.Cells[i, "J"].Value2);
                    ////qarsafari.WoodyPlantPercent = Convert.ToDouble(xlRange.Cells[i,"K"].Value2);
                    //qarsafari.WoodyPlantQuantity = Convert.ToDouble(xlRange.Cells[i, "L"].Value2);
                    //qarsafari.WoodyPlantSpecies = xlRange.Cells[i, "M"].Value2;
                    //qarsafari.InGoodCondition = Convert.ToDouble(xlRange.Cells[i, "N"].Value2);
                    //qarsafari.ChoppedDown = Convert.ToDouble(xlRange.Cells[i, "O"].Value2);
                    //qarsafari.Rampike = Convert.ToDouble(xlRange.Cells[i, "P"].Value2);
                    //qarsafari.SpeciesMediumAge = Convert.ToDouble(xlRange.Cells[i, "Q"].Value2);
                    //qarsafari.Company = xlRange.Cells[i, "S"].Value2;
                    //qarsafari.FieldOperator = xlRange.Cells[i, "T"].Value2;
                    //qarsafari.Owners = xlRange.Cells[i, "AA"].Value2;
                    //qarsafari.LandFieldOperator = xlRange.Cells[i, "AB"].Value2;
                    //qarsafari.Note1 = xlRange.Cells[i, "AC"].Value2;
                    //qarsafari.Date2 = xlRange.Cells[i, "AD"].Value2;
                    //qarsafari.LandGisOperator = xlRange.Cells[i, "AE"].Value2;
                    //qarsafari.Note11 = xlRange.Cells[i, "AF"].Value2;
                    //qarsafari.Date3 = xlRange.Cells[i, "AG"].Value2;
                    //qarsafari.CadCod = xlRange.Cells[i, "AH"].Value2;

                    GeographicDynamicDbContext.Qarsafaris.Add(qarsafari); // ახალი ობიექტის დამატება ბაზაში
                    GeographicDynamicDbContext.SaveChanges(); // ცვლილებების შენახვა


                    //label2.Text = Convert.ToString($"{i - 1} / {colCount - 1}");
                    //GeographicDynamicDbContext.PerformStep();

                }
                xlApp.Application.Quit();
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                xlApp.Application.Quit();
                return new Result<bool> { Success = false, StatusCode = System.Net.HttpStatusCode.BadGateway, Message = "შეცდომა მოხდა " + ex.Message };
                throw;
            }
        }


        ////აქ უნდა შემოწმდეს ლიტერი უნიკიდი  თუ მეორედება ექსელში მაშინ აღარ უდნა გააგრძელოს პროცესი 

        public Result<double?> ShemowmebaUnicLiterExcelshi()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();
            List<string> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.Where(x => x.IsUniqLiterNull == "true")
                                                                                        .OrderBy(m => m.UniqId)
                                                                                        .Select(q => $"{q.UniqId}-{q.LiterId}")
                                                                                        .ToList();
            try
            {
                return new Result<double?>
                {
                    Success = true,
                    // Data = uniqIdsNotInAccessList,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა შემოწმება Excel-ში UniqId-ის "
                };

            }
            catch
            {
                return new Result<double?>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად დასრულდა შემოწმება Excel-ში UniqId-ის "

                };
            }
        }

        //// ეს ფუქნცია ამოწმებს excel და access ცხრილებს და ადარებს UNIQID ებს თუ ემთხვევა ერთმანეთს 
        public Result<string?> ShemowmebaAccessExcelUnicLiterDublicats()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            string uniqIdsNotInAccessList = "";
            try
            {

                #region ALEKS

                //List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(m => m.IsUniqLiterNull == "true").ToList();
                //List<WindbreakMdb> windbreakMdbs = geographicDynamicDbContext.WindbreakMdbs.ToList();

                //bool emtxveva = true;
                //foreach (var excel in qarsafaris)
                //{
                //    foreach (var mdb in windbreakMdbs)
                //    {
                //        if (mdb.UniqId == excel.UniqId && mdb.LiterId == excel.LiterId)
                //        {
                //            emtxveva = false;
                //        }
                //    }
                //}


                #endregion


                #region gio
                //List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();
                //List<double?> AccessList = geographicDynamicDbContext.WindbreakMdbs.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();

                //foreach (var item in AccessList)
                //{
                //    if (!distinctUniqIds.Contains(item))
                //    {
                //        uniqIdsNotInAccessList.Add(item);
                //    }
                //}

                //if (uniqIdsNotInAccessList.Count > 0)
                //{
                //    return new Result<double?>
                //    {
                //        Success = false,
                //        Data = uniqIdsNotInAccessList,
                //        StatusCode = System.Net.HttpStatusCode.BadGateway,
                //        Message = "მოხდა შეცდომა ! Access და Excel UniqId-ები არ ემთხვევა ერთმანეთს !"
                //    };
                //}

                //return new Result<double?>
                //{
                //    Success = true,
                //    Data = uniqIdsNotInAccessList,
                //    StatusCode = System.Net.HttpStatusCode.OK,
                //    Message = "წარნატებით დასრულდა შემოწმება Access და Excel UniqId-ის "
                //};
                #endregion


                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(m => m.IsUniqLiterNull == "true").Select(x => new Qarsafari { UniqId = x.UniqId, LiterId = x.LiterId }).ToList();

                var duplicates = qarsafaris.GroupBy(q => new { q.UniqId, q.LiterId }).Where(g => g.Count() > 1).SelectMany(g => g);

                if (duplicates.Any())
                {
                    //Console.WriteLine("Duplicates found:");
                    foreach (var duplicate in duplicates)
                    {
                        uniqIdsNotInAccessList += $"{duplicate.LiterId}-{duplicate.UniqId})";
                    }
                    return new Result<string?>
                    {
                        Success = false,
                        //Data = uniqIdsNotInAccessList,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "მოხდა შეცდომა ! Excel UniqId  !"
                    };
                }

                List<WindbreakMdb> windbreakMdbs = geographicDynamicDbContext.WindbreakMdbs.Select(x => new WindbreakMdb { UniqId = x.UniqId, LiterId = x.LiterId }).ToList();

                var duplicatesMDB = qarsafaris.GroupBy(q => new { q.UniqId, q.LiterId }).Where(g => g.Count() > 1).SelectMany(g => g);

                if (duplicates.Any())
                {

                    //Console.WriteLine("Duplicates found:");
                    foreach (var duplicate in duplicatesMDB)
                    {
                        uniqIdsNotInAccessList += $"{duplicate.LiterId}-{duplicate.UniqId})";
                        return new Result<string?>
                        {
                            Success = false,
                            //Data = uniqIdsNotInAccessList,
                            StatusCode = System.Net.HttpStatusCode.BadGateway,
                            Message = "მოხდა შეცდომა ! Excel UniqId  !"
                        };
                    }



                    List<Qarsafari> resultList = qarsafaris.Where(u => windbreakMdbs.Any(l => l.LiterId == u.LiterId && l.UniqId == u.UniqId)).ToList();
                    //foreach (var excel in qarsafaris)
                    //{
                    //    bool existsInList = windbreakMdbs.Any(x => x.UniqId == excel.UniqId && x.LiterId == excel.LiterId);
                    //    if (!existsInList)
                    //    {
                    //        uniqIdsNotInAccessList.Add(string.Concat(excel.UniqId, "-", excel.LiterId, "excel"));
                    //    }
                    //}
                    //foreach (var access in windbreakMdbs)
                    //{
                    //    bool existsInList = qarsafaris.Any(x => x.UniqId == access.UniqId && x.LiterId == access.LiterId);
                    //    if (!existsInList)
                    //    {
                    //        uniqIdsNotInAccessList.Add(string.Concat(access.UniqId, "-", access.LiterId, "access"));
                    //    }
                    //}sultList = qarsafaris.Where(u => windbreakMdbs.Any(l => l.LiterId == u.LiterId && l.UniqId == u.UniqId)).ToList();


                    if (resultList.Count != 0)
                    {
                        string? concatenatedString = "";
                        foreach (var item in resultList)
                        {
                            concatenatedString += $"{item.LiterId}-{item.UniqId}";
                        }

                        return new Result<string?>
                        {
                            Success = false,
                            // Data = concatenatedString,
                            StatusCode = System.Net.HttpStatusCode.BadGateway,
                            Message = "მოხდა შეცდომა ! Excel და Access რაოდენობა არ ემთხვევა!"
                        };
                    }







                    return new Result<string?>
                    {
                        Success = true,
                        //Data = uniqIdsNotInAccessList,
                        StatusCode = System.Net.HttpStatusCode.OK,
                        Message = "წარნატებით დასრულდა შემოწმება access და excel uniqid-ის "
                    };
                }

                return new Result<string?> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };

            }
            catch
            {
                return new Result<string?>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "წარუმატებლად შესრულდა შემოწმება access და excel uniqid-ის "
                };
            }
        }
        //ეს ფუნქცია მიდის და ქარსაფარის ცხრილში სახეობების მიხედვით აკეთებს ვარჯის ფართების ჩაწერას
        public Result<bool> ChaweraVarjisParti(int ProjectNameID)
        {
            try
            {
                // Creating a GeographicDynamicDbContext instance
                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<VarjisFarti> varjisFartis = GeographicDynamicDbContext.VarjisFartis.Where(x => x.AreaNameId == ProjectNameID).ToList();
                foreach (var item in varjisFartis)
                {
                    var saxeobaName = GeographicDynamicDbContext.Dictionaries.FirstOrDefault(m => m.Id == item.SaxeobaId).Name;
                    List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.Where(x => x.WoodyPlantSpecies == saxeobaName).ToList();
                    foreach (var qarsafariItem in qarsafaris)
                    {
                        qarsafariItem.VarjisFarti = item.VarjisFarti1;

                    }
                    GeographicDynamicDbContext.SaveChanges();
                }

                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                // Returning failure result with error message
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ვარჯის ფართების გადათვლისას: " + ex.Message
                };
            }
        }

        public Result<bool> CheckerOfVarjisFartiandSaxeoba()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();
            try
            {
                foreach (var item in qarsafaris)
                {
                    if (item.VarjisFarti == null && item.WoodyPlantSpecies != null)
                    {
                        return new Result<bool>
                        {
                            Success = false,
                            StatusCode = System.Net.HttpStatusCode.BadRequest,
                            Message = "ვარჯისფართი არ ჩაიწერა სადაც სახეობა გვაქ! "
                        };
                    }

                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "შემოწმებისას მოხდა შეცდომა ვარჯისფართი არ წერია სადაც სახეობა გვაქ ! " + ex.Message
                };
            }
        }
        // ეშვება მეთოდი იმისთვის რომ LITER_ID და UNIQ-ID შეერთდეს და ჩაიწეროს UID-ში
        public Result<bool> UIDReplaceExcel()
        {
            try
            {
                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.ToList();
                foreach (var item in qarsafaris)
                {
                    item.Uid = item.LiterId.ToString() + item.UniqId.ToString();
                    //item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();
                }
                return new Result<bool> { Success = true, StatusCode = System.Net.HttpStatusCode.OK };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }
        // ფუნქცია გამოიყენება რომ წაიკითხოს Access ფაილი და შეყაროს SQL ბაზაში 
        public Result<bool> AccessWakitxva(string AccessFilePath, string AccessShitName)
        {
            var GeographicDynamicDbContext = new GeographicDynamicDbContext();
            var AccessPath = AccessFilePath;
            var AccessShitN = AccessShitName;

            GeographicDynamicDbContext.WindbreakMdbs.ExecuteDelete();

            #region

            //OleDbConnection

            //string connectionString = @"Provider=Microsoft.ACE.OLEDB.16.0;Data Source=C:\Users\gioch\OneDrive\Desktop\GEOGraphics\test.accdb";
            //string connectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\gioch\\OneDrive\\Desktop\\GEOGraphics\Dedoplistskaro.mdb";
            string connectionString = "";
            if (Path.GetExtension(AccessPath).ToLower().Trim() == ".mdb" && Environment.Is64BitOperatingSystem == false)
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + AccessPath;
                connectionString = "Provider=Microsoft.Jet.OLEDBMicrosoft.Jet.OLEDB.4.0;Data Source=" + AccessPath + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
            }
            else
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + AccessPath;
            }

            #endregion

            string strSQL = "SELECT * FROM " + AccessShitN;
            // Create a connection    
            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                // Create a command and set its connection    
                OleDbCommand command = new OleDbCommand(strSQL, connection);
                // Open the connection and execute the select command.    
                try
                {

                    // Open connecton    
                    connection.Open();
                    // Execute command    
                    using (OleDbDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {

                            WindbreakMdb windbreakMdb = new WindbreakMdb();// იყო ზევით და არ მუშაობდა რადგან ყოველ ჯერზე ახალი შექმნას და არ იყოს შევსებული 
                            foreach (var columnName in GeographicDynamicDbContext.ColumnNames)
                            {


                                //Get cell type
                                if (columnName.AccessName != null)
                                {
                                    object cellValue = reader[columnName.AccessName];
                                    if (cellValue != null)
                                    {
                                        Type cellType = cellValue.GetType();
                                        PropertyInfo propertyInfo = typeof(WindbreakMdb).GetProperty(columnName.Sqlname);
                                        if (propertyInfo != null)
                                        {
                                            ;
                                            // Handle conversion based on cell type
                                            if (cellType == typeof(System.Single))
                                            {
                                                Double? doubleValue = Convert.ToDouble(cellValue);
                                                propertyInfo.SetValue(windbreakMdb, doubleValue);
                                            }
                                            else if (cellType == typeof(string))
                                            {
                                                propertyInfo.SetValue(windbreakMdb, cellValue);
                                            }
                                            else if (cellType == typeof(DateTime))
                                            {
                                                DateTime dateTimeValue;
                                                if (DateTime.TryParse((string)cellValue, out dateTimeValue))
                                                {
                                                    propertyInfo.SetValue(windbreakMdb, dateTimeValue);
                                                }
                                                // Handle DateTime conversion if necessary
                                            }
                                            // Add other type conversions as necessary
                                        }
                                    }

                                }

                            }

                            GeographicDynamicDbContext.WindbreakMdbs.Add(windbreakMdb);
                            GeographicDynamicDbContext.SaveChanges();
                            //// ველების წაკითხვის და ბაზაში გაშვების/დამახსოვრების ციკლი 
                            //while (reader.Read())
                            //{
                            //    WindbreakMdb windbreakMdbs = new WindbreakMdb();

                            //    windbreakMdbs.UniqId = (float?)reader["UNIQ_ID"];
                            //    windbreakMdbs.LiterId = (float?)reader["LITER_ID"];
                            //    windbreakMdbs.AdmMun = reader["Adm_Mun"].ToString();
                            //    windbreakMdbs.CityTownVillage = reader["City_Town_Village"].ToString();
                            //    windbreakMdbs.LandAreaSqM = (float?)reader["Land_Area_Sq_m"];
                            //    windbreakMdbs.LandAreaHa = (float?)reader["Land_Area_Ha"];
                            //    windbreakMdbs.LegalPerson = reader["Legal_person"].ToString();
                            //    windbreakMdbs.Note = reader["Note_"].ToString();
                            //    windbreakMdbs.Date = reader["Date_"].ToString();
                            //    windbreakMdbs.GisOperator = reader["Gis_Operator"].ToString();
                            //    windbreakMdbs.DaTe1 = reader["DaTe_1"].ToString();
                            //    windbreakMdbs.OverlapCadCode = reader["Overlap_CAD_COD"].ToString();
                            //    windbreakMdbs.Owner = reader["Owner"].ToString();



                            //    // ამატებს SQL ბაზაში და ამახსოვრებს ცვლილებებს 
                            //    GeographicDynamicDbContext.WindbreakMdbs.Add(windbreakMdbs);
                            //    GeographicDynamicDbContext.SaveChanges();

                            //    //Console.WriteLine("{0} {1}", reader["Name"].ToString(), reader["Address"].ToString());
                            //}


                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                // The connection is automatically closed becasuse of using block.    
            }

            return new Result<bool>
            {
                Success = true,
                StatusCode = System.Net.HttpStatusCode.OK
            };
        }

        //ამ ფუნქციაზი ხდება შემოწმება UniqID - ების ექსელში და აქსესში 
        // ფუნქცია გამოიყენება რომ მოხდეს Access ფაილიდან წაკითხული მონაცემები და გადასული ინფორმაციის UID ველის შევსება ლიტერის და უნიკაიდის კონკატენაციით 
        public Result<bool> UIDReplaceAccess()
        {
            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<WindbreakMdb> windbreakMdbs = GeographicDynamicDbContext.WindbreakMdbs.ToList();
                foreach (var item in windbreakMdbs)
                {
                    item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();

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
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }
        // ფუნქცია გამოიყენება რომ დაგაიწეროს Access ფაილიდან საჭირო მონაცემები Excel-ში 
        public Result<bool> UpdateFromAccessToExcell()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            List<WindbreakMdb> AccessList = geographicDynamicDbContext.WindbreakMdbs.ToList();
            List<Qarsafari> ExcelList = geographicDynamicDbContext.Qarsafaris.ToList();
            try
            {
                foreach (var excel in ExcelList)
                {
                    if (excel.IsUniqLiterNull == "true")
                    {
                        WindbreakMdb access = geographicDynamicDbContext.WindbreakMdbs.FirstOrDefault(x => x.LiterId == excel.LiterId && x.UniqId == excel.UniqIdOld);
                        if (access != null)
                        {
                            foreach (var ColumnName in geographicDynamicDbContext.ColumnNames)
                            {

                                if (ColumnName.IsAccessToExcel == true)
                                {
                                    PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(ColumnName.Sqlname);

                                    if (propertyInfoQarsafari != null)
                                    {
                                        PropertyInfo propertyInfoWindbreakMDB = typeof(WindbreakMdb).GetProperty(ColumnName.Sqlname);
                                        object propertyValueWindbreakMDB = propertyInfoWindbreakMDB.GetValue(access);
                                        // თუ ცარიელია ვამოწმებთ და ვწერთ null -ს 
                                        if (propertyValueWindbreakMDB != null)
                                        {

                                            //აქ სადაც double ია მოაქვს System.Single ამიტომ ვამოწმებთ
                                            if (propertyValueWindbreakMDB.GetType() == typeof(System.Single))
                                            {
                                                Double? intValue = Convert.ToDouble(propertyValueWindbreakMDB);
                                                propertyInfoQarsafari.SetValue(excel, intValue, null);
                                            }
                                            else // სხვა შემთხვევაში არის string ან date ან bit
                                            {
                                                propertyInfoQarsafari.SetValue(excel, propertyInfoWindbreakMDB.GetValue(access, null), null);
                                            }
                                        }
                                        else
                                        {
                                            propertyInfoQarsafari.SetValue(excel, null, null);
                                        }
                                    }
                                }
                            }

                            geographicDynamicDbContext.SaveChanges();
                        }
                    }
                }
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა Access-დან Excel-ში "
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა Access-დან Excel-ში გადაწერის დროს" + ex.Message
                };
            }




        }

        public Result<bool> QarsafariProcentisDatvla()
        {
            try
            {
                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                List<Qarsafari> ExcelList = geographicDynamicDbContext.Qarsafaris.ToList();

                foreach (var Excel in ExcelList)
                {
                    Excel.ChoppedDownQuantity = Excel.ChoppedDown;
                    if (Excel.ChoppedDown != 0 && Excel.ChoppedDown != null)
                    {
                        Excel.ChoppedDown = (Excel.ChoppedDown / Excel.WoodyPlantQuantity) * 100;
                        Excel.Gachexili = (Excel.ChoppedDown / Excel.WoodyPlantQuantity) * 100;
                    }
                    if (Excel.InGoodCondition != 0 && Excel.InGoodCondition != null)
                    {
                        Excel.InGoodCondition = (Excel.InGoodCondition / Excel.WoodyPlantQuantity) * 100;
                    }
                    if (Excel.Rampike != 0 && Excel.Rampike != null)
                    {
                        Excel.Rampike = (Excel.Rampike / Excel.WoodyPlantQuantity) * 100;
                    }

                    geographicDynamicDbContext.SaveChanges();
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა ქარსაფარში პროცენტის დათვლა "
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ქარსაფარში პროცენტის დათვლის დროს" + ex.Message
                };
            }

        }

        //ფუნქცია გამოიყენება რომ დაგაინომროს UNIQ_ID ები ქარსაფარის ცხრილში 
        public Result<bool> QarsafariGadanomrva(int UnicIDStartNumber)
        {
            var geographicDynamicDbContext = new GeographicDynamicDbContext();
            try
            {
                #region ALEKS
                List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.ToList();
                //UniqId გადაგვაქვს UniqIdOld -ში ძველი უნიკიდის შესანახად
                foreach (var qarsafari in qarsafaris)
                {
                    if (qarsafari.IsUniqLiterNull == "true")
                    {
                        qarsafari.UniqIdOld = Convert.ToInt32(qarsafari.UniqId);
                    }
                }
                geographicDynamicDbContext.SaveChanges();

                //გადანომრვა
                var newUniqueID = UnicIDStartNumber - 1;
                //გლობალურად ვინათავთ ლიტერაიდის რომ შემდეგ იტერაციაში გამოვიყენოთ 
                Double? literid = null;
                foreach (var qarsafari in qarsafaris)
                {
                    if (qarsafari.IsUniqLiterNull == "true")
                    {
                        newUniqueID++;
                        //აქ იღებს ლიტერაიდი მნიშვნელობას როდესაც ზედა if პირობა სრულდება მაშინ იცვლის მნიშვნელობას 
                        literid = qarsafari.LiterId;
                    }
                    qarsafari.UniqId = newUniqueID;

                    //აქ უკვე იწერება ქარსაფარში 
                    qarsafari.LiterId = literid;
                }
                //ვიმახსოვრებთ შედეგებს 
                geographicDynamicDbContext.SaveChanges();
                #endregion
                #region GIO
                //List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.OrderBy(x => x.LiterId).ThenBy(x => x.UniqId).ToList();

                //foreach (var item in qarsafaris)
                //{
                //    var implement = UnicIDStartNumber;
                //    item.UniqIdOld = Convert.ToInt32(item.UniqId);
                //    if (item.Municipality != null)
                //    {
                //        item.UniqId = implement;
                //    }

                //    foreach (var item1 in qarsafaris)
                //    {
                //        if(item1.Municipality == null)
                //        {
                //            //qarsafaris.FirstOrDefault()
                //        }
                //    }
                //    implement++;
                //}
                #endregion
                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა გადანომვრის პროცესი"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა გადანომვრის პროცესის დროს" + ex.Message
                };
            }
        }

        // აქ გვჭირდება შემმოწმება ფუნქციის ჩაწერა რომელიც გადაამოწმებს თუ სადმე ხეხილი მეორდება უბანზე 
        public Result<bool> QarsafariXexilisShemowmeba()
        {
            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

            try
            {
                List<Qarsafari> qarsafariList = geographicDynamicDbContext.Qarsafaris.ToList();
                List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();
                var count = 0;
                foreach (var uniqueID in distinctUniqIds)
                {
                    List<string> xexilebiList = new List<string>();
                    List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(x => x.UniqId == uniqueID).ToList();
                    foreach (var item in qarsafaris)
                    {
                        if (!string.IsNullOrEmpty(item.WoodyPlantSpecies))
                        {
                            xexilebiList.Add(item.WoodyPlantSpecies);
                        }
                    }
                    if (xexilebiList != null && xexilebiList.Count != xexilebiList.Distinct().Count())
                    {
                        count++;
                    }
                }
                if (count > 0)
                {
                    return new Result<bool>
                    {
                        Success = false,
                        StatusCode = System.Net.HttpStatusCode.BadGateway,
                        Message = "მოხდა შეცდომა! ხეხილის ჯიში მეორდება"
                    };

                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა გადაწერა გადანომვრის პროცესი"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა გადანომვრის პროცესის დროს" + ex.Message
                };
            }
        }

        //ფუნქცია გამოიყენება რომ დაიგრუპოს ხეხილის სახეობები და ამასთან მიყვეს სხვა პროცედურებიც რაც დაგრუპვაში შედის (პატარა ექსელი) 
        public Result<bool> QarsafariToQarsafariGrouped()
        {
            try
            {

                //თავიდან უნდა გავასუფთავოთ qarsafariGrouped ცხრილი

                GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext();

                geographicDynamicDbContext.QarsafariGroupeds.ExecuteDelete();
                // List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.Where(m => m.UniqId ).Select(q => q.UniqId).Distinct().ToList(); ერთი კონკრეტული როუს გასატესტად 
                List<double?> distinctUniqIds = geographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).Select(q => q.UniqId).Distinct().ToList();

                foreach (var uniqueID in distinctUniqIds)
                {
                    List<Qarsafari> qarsafaris = geographicDynamicDbContext.Qarsafaris.Where(x => x.UniqId == uniqueID).ToList();
                    QarsafariGrouped qarsafariGrouped = new QarsafariGrouped();



                    Qarsafari qarsafariExcel = qarsafaris.FirstOrDefault(m => m.IsUniqLiterNull == "true");


                    #region ახალი გადინამიურებული დაგრუპვა (დარჩენილია ფორმულების გაკეთება)
                    //if (qarsafariExcel != null)
                    //{
                    //    foreach (var columnName in geographicDynamicDbContext.ColumnNames)
                    //    {
                    //        //Get cell type
                    //        if (columnName.GroupMethod == "MAX")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {
                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);
                    //                object propertyValueQarsafari = propertyInfoQarsafari.GetValue(qarsafariExcel);

                    //                // თუ ცარიელია ვამოწმებთ და ვწერთ null -ს 
                    //                if (propertyValueQarsafari != null)
                    //                {

                    //                    //აქ სადაც double ია მოაქვს System.Single ამიტომ ვამოწმებთ

                    //                    if (propertyValueQarsafari.GetType() == typeof(System.Single) || propertyValueQarsafari.GetType() == typeof(System.Double))
                    //                    {
                    //                        Double? doubleValue = Convert.ToDouble(propertyValueQarsafari);

                    //                        propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, doubleValue, null);
                    //                    }
                    //                    else // სხვა შემთხვევაში არის string ან date ან bit
                    //                    {

                    //                        propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, propertyInfoQarsafari.GetValue(qarsafariExcel, null), null);
                    //                    }
                    //                }
                    //                else
                    //                {
                    //                    propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, null, null);
                    //                }

                    //                //if (propertyInfo != null)
                    //                //{

                    //                //    propertyInfo.SetValue(qarsafariGrouped, qarsafaris);
                    //                //}
                    //            }
                    //        }
                    //        else if (columnName.GroupMethod == "SUBSTRING")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {

                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);

                    //                var storeStringValue = "";
                    //                foreach (var item in qarsafaris)
                    //                {
                    //                    object propertyValueQarsfari = propertyInfoQarsafari.GetValue(item);

                    //                    if (propertyValueQarsfari != null)
                    //                    {

                    //                        if (propertyValueQarsfari.GetType() == typeof(System.String))
                    //                        {
                    //                            storeStringValue += string.Concat(propertyValueQarsfari, "/"); ;
                    //                            propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, storeStringValue.TrimEnd('/'), null);
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }


                    //        else if (columnName.GroupMethod == "SUM")
                    //        {
                    //            if (qarsafariExcel != null)
                    //            {

                    //                PropertyInfo propertyInfoQarsafariGrouped = typeof(QarsafariGrouped).GetProperty(columnName.Sqlname);
                    //                PropertyInfo propertyInfoQarsafari = typeof(Qarsafari).GetProperty(columnName.Sqlname);

                    //                Double? storedValue = 0;
                    //                foreach (var item in qarsafaris)
                    //                {
                    //                    object propertyValueQarsafari = propertyInfoQarsafari.GetValue(item);
                    //                    if (propertyValueQarsafari != null)
                    //                    {
                    //                        if ((propertyValueQarsafari.GetType() == typeof(System.Single)) || (propertyValueQarsafari.GetType() == typeof(System.Double)))
                    //                        {
                    //                            storedValue += (double)propertyValueQarsafari;

                    //                            propertyInfoQarsafariGrouped.SetValue(qarsafariGrouped, storedValue, null);
                    //                        }
                    //                    }
                    //                }
                    //            }
                    //        }
                    //    }
                    //    geographicDynamicDbContext.QarsafariGroupeds.Add(qarsafariGrouped);
                    //    geographicDynamicDbContext.SaveChanges();
                    //}

                    #endregion



                    #region ჩვეულებრივი ხელით
                    if (qarsafariExcel != null)
                    {
                        //Double? woodyplantqunatity = 0;
                        Double? merqmcenarisPr = 0;
                        Double? romgitxariGadanomrili = 0;
                        string mcenarisSaxeobebi = "";
                        string mcenarisSaxeobebiCorrected = "";
                        //double? InGoodCondition = 100;


                        //კარგ მდგომარეობაში ჩასაწერი ველისთვის 
                        Double? sumWoodyPlantQuantitymultiplyChoppedDown = 0;
                        Double? sumWoodyPlantQuantity = 0;
                        Double? sumWoodyPlantQuantitymultiplyRampike = 0;

                        //საშუალო ხმოვანებაში ჩასაწერი ველისთვის 
                        double? speciesMidAge = 0;
                        foreach (var qarsafari in qarsafaris)
                        {


                            //qarsafari.Sakutreba = qarsafariExcel.Owner == null ? "სახელმწიფო" : qarsafariExcel.Owner;
                            ////ვითვლით საკუთრებას სადაც NULL ებია რომ შეივსოს ყველასთვის რომ დადუბლირდეს რაც მთავარშია და IsUniqLiterNull არის "true"
                            //if (qarsafari.Owner == "მუნიციპალიტეტი") //მოწმდება სადმე თუ წერი Owner-ში მუნიციპალიტეტი და იწერება სახელმწიფო საკუთრებაში
                            //{
                            //    qarsafari.Sakutreba = "მუნიციპალიტეტი";
                            //}
                            //if (qarsafari.Owner == "იურიდიული პირი")//მოწმდება სადმე თუ წერი Owner-ში იურიდიული პირი და იწერება კერძო საკუთრებაში
                            //{
                            //    qarsafari.Sakutreba = "იურიდიული პირი";
                            //}
                            //woodyplantqunatity += qarsafari.WoodyPlantQuantity;
                            // ამაში ამოვაგდეთ გაჩეხილი (chopped_down)
                            merqmcenarisPr += qarsafari.ChoppedDownQuantity != null ? (((qarsafari.WoodyPlantQuantity - qarsafari.ChoppedDownQuantity) * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100 : ((qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100;
                            //merqmcenarisPr += ((qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti) / qarsafariExcel.LandAreaSqM) * 100;
                            romgitxariGadanomrili += qarsafari.WoodyPlantQuantity * qarsafari.VarjisFarti;
                            //mcenarisSaxeobebi += string.Concat(qarsafari.WoodyPlantSpecies, "/ "); // აქ იყო  "/ "
                            if (!string.IsNullOrEmpty(qarsafari.WoodyPlantSpecies))// კეთდება შემოწმება იმისთვის რომ გავიგოთ სლექში ჩაიწეროს თუ არა ველში 
                            {
                                mcenarisSaxeobebi += qarsafari.WoodyPlantSpecies + "/";
                            }

                            //InGoodCondition -= ((qarsafari.WoodyPlantQuantity * qarsafari.ChoppedDown) / 100) / qarsafari.WoodyPlantQuantity;

                            // კარგ მდომარეობაში ჩასაწერი 
                            // ვითვლით SUM([Woody_plant_quantity] * [chopped_down] / 100)
                            sumWoodyPlantQuantitymultiplyChoppedDown += ((qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity) * (qarsafari.ChoppedDown == null ? 0 : qarsafari.ChoppedDown)) / 100;
                            // ვითვლით მცენარეების რაოდენობის მთლიან ჯამს
                            sumWoodyPlantQuantity += qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity;
                            // ვითვლით SUM([Woody_plant_quantity] * [rampike] / 100)
                            sumWoodyPlantQuantitymultiplyRampike += ((qarsafari.WoodyPlantQuantity == null ? 0 : qarsafari.WoodyPlantQuantity) * (qarsafari.Rampike == null ? 0 : qarsafari.Rampike)) / 100;

                            //საშუალო ხმოვანეის ჩასაწერი 
                            speciesMidAge += (qarsafari.SpeciesMediumAge * qarsafari.WoodyPlantQuantity);

                        }


                        qarsafariGrouped.UniqId = qarsafariExcel.UniqId;
                        qarsafariGrouped.LiterId = qarsafariExcel.LiterId;
                        qarsafariGrouped.PhotoN = qarsafariExcel.PhotoN;
                        qarsafariGrouped.Region = qarsafariExcel.Region;
                        qarsafariGrouped.Municipality = qarsafariExcel.Municipality;
                        qarsafariGrouped.AdmMun = qarsafariExcel.AdmMun;
                        qarsafariGrouped.CityTownVillage = qarsafariExcel.CityTownVillage;
                        qarsafariGrouped.LandAreaSqM = qarsafariExcel.LandAreaSqM;
                        qarsafariGrouped.LandAreaHa = Math.Round(Convert.ToDouble(qarsafariExcel.LandAreaHa), 1);
                        qarsafariGrouped.Shrubbery = qarsafariExcel.Shrubbery;
                        qarsafariGrouped.WoodyPlantPercent = Math.Round(Convert.ToDouble(merqmcenarisPr), 1);
                        qarsafariGrouped.WoodyPlantQuantity = sumWoodyPlantQuantity;
                        qarsafariGrouped.VarjisFarti = romgitxariGadanomrili;
                        // ვაშორებთ ბოლო "/" -ს
                        mcenarisSaxeobebiCorrected = mcenarisSaxeobebi.TrimEnd('/');
                        qarsafariGrouped.WoodyPlantSpecies = mcenarisSaxeobebiCorrected;
                        // კარგ მდგომარეობაში არის 100 - გამხმარი - გაჩეხილი
                        //თუ ხეები არ გვაქვს ჩაწეროს 0

                        qarsafariGrouped.InGoodCondition = sumWoodyPlantQuantity == 0 ? 0 : 100 - (sumWoodyPlantQuantity == 0 ? 0 : (Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyChoppedDown / sumWoodyPlantQuantity) * 100), 0) + Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyRampike / sumWoodyPlantQuantity) * 100), 0)));
                        qarsafariGrouped.ChoppedDown = sumWoodyPlantQuantity == 0 ? 0 : (Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyChoppedDown / sumWoodyPlantQuantity) * 100), 0));
                        qarsafariGrouped.Rampike = sumWoodyPlantQuantity == 0 ? 0 : Math.Round(Convert.ToDouble((sumWoodyPlantQuantitymultiplyRampike / sumWoodyPlantQuantity) * 100), 0);
                        qarsafariGrouped.SpeciesMediumAge = RoundToNearest(Convert.ToDouble(speciesMidAge / sumWoodyPlantQuantity)); // Math.Round(Convert.ToDouble((((speciesMidAge / sumWoodyPlantQuantity / 5) * 0.5) * 10)), 0);
                        qarsafariGrouped.Note = qarsafariExcel.Note;
                        qarsafariGrouped.Company = qarsafariExcel.Company;
                        qarsafariGrouped.LandGisOperator = qarsafariExcel.LandGisOperator;
                        qarsafariGrouped.Date = qarsafariExcel.Date;
                        qarsafariGrouped.GisOperator = qarsafariExcel.GisOperator;
                        qarsafariGrouped.DaTe1 = qarsafariExcel.DaTe1;
                        qarsafariGrouped.OverlapCadCode = qarsafariExcel.OverlapCadCode;
                        qarsafariGrouped.Owner = qarsafariExcel.Owner;
                        qarsafariGrouped.Sakutreba = qarsafariExcel.Sakutreba;
                        qarsafariGrouped.LegalPerson = qarsafariExcel.LegalPerson;
                        qarsafariGrouped.UniqIdOld = qarsafariExcel.UniqIdOld;

                        //რაუნდები დაუწერე აქ, ზემოთ არ დაუწერო
                        //qarsafariGrouped.WoodyPlantPercent = merqmcenarisPr;


                        geographicDynamicDbContext.QarsafariGroupeds.Add(qarsafariGrouped);
                        geographicDynamicDbContext.SaveChanges();
                    }
                    #endregion
                }

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                    Message = "წარმატებით დასრულდა დაგრუპვა ქარსაფარის ცხრილის"
                };
            }
            catch (Exception ex)
            {
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ქარსაფარის ცხრილის დაგრუპვის დროს" + ex.Message
                };
            }
        }

        // qarsafariGrouped ცხრილის UID ველის შევსება ლიტერით და უნიკაიდით 
        public Result<bool> UIDReplaceQarsafariGrouped()
        {

            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();
                List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();
                foreach (var item in qarsafariGroupeds)
                {
                    item.Uid = String.Concat(item.LiterId, item.UniqId);
                    GeographicDynamicDbContext.SaveChanges();

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
                    Message = "მოხდა შეცდომა UID replace-ის დროს" + ex.Message
                };
            }
        }

        //ფუნქცია იმისთვის რომ Access ფაილში ჩაიწეროს QarsafariGroupded-ან 
        public Result<bool> UPDTFromExcelToAccess(string AccessShitName)
        {
            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<WindbreakMdb> windbreakMdbs = GeographicDynamicDbContext.WindbreakMdbs.ToList();
                List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();

                if (!string.IsNullOrEmpty(AccessShitName))
                {
                    foreach (var item in windbreakMdbs)
                    {
                        //WindbreakMdb access = GeographicDynamicDbContext.WindbreakMdbs.FirstOrDefault(x => x.LiterId == excel.LiterId && x.UniqId == excel.UniqId);
                        QarsafariGrouped ExcelGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.Uid == item.Uid);
                        if (ExcelGrouped != null)
                        {
                            item.PhotoN = ExcelGrouped.PhotoN;
                            item.Shrubbery = Convert.ToDouble(ExcelGrouped.Shrubbery);
                            item.WoodyPlantPercent = Convert.ToString(ExcelGrouped.WoodyPlantPercent);
                            item.WoodyPlantQuantity = Convert.ToDouble(ExcelGrouped.WoodyPlantQuantity);
                            item.WoodyPlantSpecies = ExcelGrouped.WoodyPlantSpecies;
                            item.InGoodCondition = Convert.ToDouble(ExcelGrouped.InGoodCondition);
                            item.ChoppedDown = Convert.ToDouble(ExcelGrouped.ChoppedDown);
                            item.Rampike = Convert.ToDouble(ExcelGrouped.Rampike);
                            item.SpeciesMediumAge = Convert.ToDouble(ExcelGrouped.SpeciesMediumAge);
                            item.Company = ExcelGrouped.Company;
                            item.FieldOperator = ExcelGrouped.FieldOperator;
                            item.UniqId = (float?)ExcelGrouped.UniqId;

                            GeographicDynamicDbContext.SaveChanges();

                        }
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
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა Excel-ცხრილიდან Access-ცხრილში გადაწერის დროს" + ex.Message
                };
            }
        }

        // ფუნქცია ყრის მონაცემებს gadanomriliFotoebi-დან qarsafariGroupded-ში 
        public Result<bool> GadanomriliFotoebiToQarsafariGrouped()
        {
            try
            {

                var GeographicDynamicDbContext = new GeographicDynamicDbContext();

                List<GadanomriliFotoebi> FotoList = GeographicDynamicDbContext.GadanomriliFotoebis.ToList();
                List<QarsafariGrouped> qarsafariGroupeds = GeographicDynamicDbContext.QarsafariGroupeds.ToList();

                {
                    foreach (var item in FotoList)
                    {
                        QarsafariGrouped ExcelGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.LiterId == item.LiterId && x.UniqId == Convert.ToDouble(item.UniqId));
                        if (ExcelGrouped != null)
                        {
                            ExcelGrouped.PhotoN = item.PhotoN;
                            ExcelGrouped.Date = item.PhotoDate;

                            GeographicDynamicDbContext.SaveChanges();

                        }
                    }


                    return new Result<bool>
                    {
                        Success = true,
                        StatusCode = System.Net.HttpStatusCode.OK
                    };

                }
            }

            catch (Exception ex)
            {

                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა GadanomriliPhotoebi-ცხრილიდან QarsafariGrouped-ცხრილში გადაწერის დროს" + ex.Message
                };
            }
        }
        public Result<bool> WriteToExcel(List<Qarsafari> qarsafaris, string ExcelDestinationPath, string ExcelName)
        {
            var GeographicDynamicDbContext = new GeographicDynamicDbContext();

            //List<Qarsafari> qarsafaris = GeographicDynamicDbContext.Qarsafaris.OrderBy(m => m.UniqId).ToList();
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook ExcelWorkBook = null;
            ExcelApp.Visible = false;
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            Worksheet ExcelWorkSheet = ExcelWorkBook.Worksheets[1] as Worksheet;
            try
            {



                // ამით ივსება სათაურების ველები 
                ExcelWorkSheet.Cells[1, "A"] = "UNIQ_ID";
                ExcelWorkSheet.Cells[1, "B"] = "Liter_ID";
                ExcelWorkSheet.Cells[1, "C"] = "Photo_N";
                ExcelWorkSheet.Cells[1, "C"].EntireColumn.NumberFormat = "@";
                ExcelWorkSheet.Cells[1, "D"] = "Region";
                ExcelWorkSheet.Cells[1, "E"] = "Municipality";
                ExcelWorkSheet.Cells[1, "F"] = "Adm_Mun";
                ExcelWorkSheet.Cells[1, "G"] = "City_Town_Village";
                ExcelWorkSheet.Cells[1, "H"] = "Land_Area_Sq_M";
                ExcelWorkSheet.Cells[1, "I"] = "Land_Area_Ha";
                ExcelWorkSheet.Cells[1, "J"] = "Shrubbery";
                ExcelWorkSheet.Cells[1, "K"] = "Woody_Plant_Percent";
                ExcelWorkSheet.Cells[1, "L"] = "Woody_Plant_Quantity";
                ExcelWorkSheet.Cells[1, "M"] = "Woody_Plant_Spices";
                ExcelWorkSheet.Cells[1, "N"] = "VarjisFarti";
                ExcelWorkSheet.Cells[1, "O"] = "In_Good_Condition";
                ExcelWorkSheet.Cells[1, "P"] = "Chopped_down";
                ExcelWorkSheet.Cells[1, "Q"] = "Rampike";
                ExcelWorkSheet.Cells[1, "R"] = "Spices_Medium_Age";
                ExcelWorkSheet.Cells[1, "S"] = "Note_";
                ExcelWorkSheet.Cells[1, "T"] = "Company";
                ExcelWorkSheet.Cells[1, "U"] = "Field_Operator";
                ExcelWorkSheet.Cells[1, "V"] = "Date_";
                ExcelWorkSheet.Cells[1, "V"].EntireColumn.NumberFormat = "@"; //ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "W"] = "Gis_Operator";
                ExcelWorkSheet.Cells[1, "X"] = "DaTe_1";
                ExcelWorkSheet.Cells[1, "X"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "Y"] = "Overlap_CAD_CODE";
                ExcelWorkSheet.Cells[1, "Z"] = "Owner";
                ExcelWorkSheet.Cells[1, "AA"] = "Legal_person";
                ExcelWorkSheet.Cells[1, "AB"] = "Owners";
                ExcelWorkSheet.Cells[1, "AC"] = "Land_Field_Operator";
                ExcelWorkSheet.Cells[1, "AD"] = "Note1";
                ExcelWorkSheet.Cells[1, "AE"] = "Date_2";
                ExcelWorkSheet.Cells[1, "AF"] = "Land_Gis_Operator";
                ExcelWorkSheet.Cells[1, "AG"] = "Note1_1";
                ExcelWorkSheet.Cells[1, "AH"] = "Date_3";
                ExcelWorkSheet.Cells[1, "AI"] = "CAD_COD";
                ExcelWorkSheet.Cells[1, "AJ"] = "UNIQ_ID_OLD";
                ExcelWorkSheet.Cells[1, "AK"] = "UNIQ_ID_NEW";
                ExcelWorkSheet.Cells[1, "AL"] = "UID";
                ExcelWorkSheet.Cells[1, "AM"] = "ID";

                // ამ ციკლით ივსება Rows სათაურების ქვეშ 
                for (int r = 0; r < qarsafaris.Count(); r++) //r stands for ExcelRow and c for ExcelColumn
                {


                    QarsafariGrouped qarsafariGrouped = GeographicDynamicDbContext.QarsafariGroupeds.FirstOrDefault(x => x.UniqId == qarsafaris[r].UniqId);
                    //ExcelWorkSheet.Cells[r + 2, "A"] = qarsafaris[r].UniqId;
                    if (qarsafaris[r].IsUniqLiterNull == "true")
                    {
                        ExcelWorkSheet.Cells[r + 2, "A"] = qarsafaris[r].UniqId;
                        ExcelWorkSheet.Cells[r + 2, "B"] = qarsafaris[r].LiterId;
                        ExcelWorkSheet.Cells[r + 2, "AJ"] = qarsafaris[r].UniqIdOld;
                        ExcelWorkSheet.Cells[r + 2, "K"] = Math.Round(Convert.ToDouble(qarsafariGrouped.WoodyPlantPercent), 1);
                        ExcelWorkSheet.Cells[r + 2, "C"] = qarsafariGrouped.PhotoN; // ფოტოები მოდის დაგრუპულიდან 
                        ExcelWorkSheet.Cells[r + 2, "D"] = qarsafaris[r].Region;
                        ExcelWorkSheet.Cells[r + 2, "E"] = qarsafaris[r].Municipality;
                        ExcelWorkSheet.Cells[r + 2, "F"] = qarsafaris[r].AdmMun;
                        ExcelWorkSheet.Cells[r + 2, "G"] = qarsafaris[r].CityTownVillage;
                        ExcelWorkSheet.Cells[r + 2, "H"] = qarsafaris[r].LandAreaSqM;
                        ExcelWorkSheet.Cells[r + 2, "I"] = qarsafaris[r].LandAreaHa;
                        ExcelWorkSheet.Cells[r + 2, "V"] = qarsafariGrouped.Date;
                    }




                    ExcelWorkSheet.Cells[r + 2, "J"] = qarsafaris[r].Shrubbery;
                    ExcelWorkSheet.Cells[r + 2, "L"] = qarsafaris[r].WoodyPlantQuantity;
                    ExcelWorkSheet.Cells[r + 2, "M"] = qarsafaris[r].WoodyPlantSpecies;
                    ExcelWorkSheet.Cells[r + 2, "N"] = qarsafaris[r].VarjisFarti;

                    // მოწმდება თუ სადმე sumofGoodChoppedRampike განსხვავდება 0-ს ან 100-ს იდეაში 0.1 ან მეტია ან ნაკლები და მაგის მიხედვით 
                    // ხორციელდება გამოკლება ან მიმატება 0.1-ის 
                    var sumofGoodChoppedRampike = Math.Round(Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1)
                        + Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1)
                        + Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1), 1);

                    switch (sumofGoodChoppedRampike)
                    {
                        case 99.9:
                            ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1) + 0.1;
                            ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                            ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            break;
                        case 100.1:
                            if (qarsafaris[r].InGoodCondition > 0)
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1) - 0.1;
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            }
                            else if (qarsafaris[r].ChoppedDown > 0)
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1) - 0.1;
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            }
                            else
                            {
                                ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                                ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                                ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1) - 0.1;
                            }
                            break;
                        default:
                            ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                            ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                            ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                            break;
                    }
                    //ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafaris[r].InGoodCondition), 1);
                    //ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafaris[r].ChoppedDown), 1);
                    //ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafaris[r].Rampike), 1);
                    ExcelWorkSheet.Cells[r + 2, "R"] = qarsafaris[r].SpeciesMediumAge;
                    ExcelWorkSheet.Cells[r + 2, "S"] = qarsafaris[r].Note;
                    ExcelWorkSheet.Cells[r + 2, "T"] = qarsafaris[r].Company;
                    ExcelWorkSheet.Cells[r + 2, "U"] = qarsafaris[r].FieldOperator;
                    //ExcelWorkSheet.Cells[r + 2, "V"] = qarsafaris[r].Date;
                    ExcelWorkSheet.Cells[r + 2, "W"] = qarsafaris[r].GisOperator;
                    ExcelWorkSheet.Cells[r + 2, "X"] = qarsafaris[r].DaTe1;
                    ExcelWorkSheet.Cells[r + 2, "Y"] = qarsafaris[r].OverlapCadCode;
                    ExcelWorkSheet.Cells[r + 2, "Z"] = qarsafaris[r].Sakutreba;
                    ExcelWorkSheet.Cells[r + 2, "AA"] = qarsafaris[r].LegalPerson;
                    ExcelWorkSheet.Cells[r + 2, "AB"] = qarsafaris[r].Owners;
                    ExcelWorkSheet.Cells[r + 2, "AC"] = qarsafaris[r].LandFieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "AD"] = qarsafaris[r].Note1;
                    ExcelWorkSheet.Cells[r + 2, "AE"] = qarsafaris[r].Date2;
                    ExcelWorkSheet.Cells[r + 2, "AF"] = qarsafaris[r].LandGisOperator;
                    ExcelWorkSheet.Cells[r + 2, "AG"] = qarsafaris[r].Note11;
                    ExcelWorkSheet.Cells[r + 2, "AH"] = qarsafaris[r].Date3;
                    ExcelWorkSheet.Cells[r + 2, "AI"] = qarsafaris[r].CadCod;
                    ExcelWorkSheet.Cells[r + 2, "AK"] = qarsafaris[r].UniqIdNew;
                    ExcelWorkSheet.Cells[r + 2, "AL"] = qarsafaris[r].Uid;
                    ExcelWorkSheet.Cells[r + 2, "AM"] = qarsafaris[r].Id;
                }

                ExcelWorkBook.Worksheets[1].Name = "დიდი-ექსელი";//Renaming the Sheet1 to MySheet
                ExcelWorkBook.SaveAs(ExcelDestinationPath + $"\\{ExcelName}.xlsx");
                ExcelWorkBook.Close();
                ExcelApp.Quit();
                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }

            catch (Exception ex)
            {
                ExcelWorkBook.Close();
                ExcelApp.Quit();
                return new Result<bool>
                {

                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა SQL-დან ახალ Execl-ში გადაწერის დროს" + ex.Message
                };
            }
        }


        //ფუნქცია კითხულობს ბაზას და ქმნის ახალ ექსელის ფაილს რომ ჩაიწეროს მონაცემები მხოლოდ დაგრუპულისთვის 
        public Result<bool> WriteToExcelGrouped(List<QarsafariGrouped> qarsafariGroupeds, string ExcelDestinationPath, string ExcelName)
        {

            GeographicDynamicDbContext geographicDynamicDbContext = new GeographicDynamicDbContext(); //უკავშირდება კონტექსტს რომ გაიგოს ცხრილები SQL-დან 

            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application(); //იქმნება აპლიკაცია წინასწარ 
            Workbook ExcelWorkBook = null; // წინასწარ იქმნება ვორკბუკის ცვლადი რომ შემდეგში გამოვიყენოთ 
            Worksheet ExcelWorkSheet = null; // ასევე წინასწარ იქმნება შიტის ცვლადი რომ გამოვიყენოთ შემდეგ 
            ExcelApp.Visible = false; // აქ ვანიჭებთ ექსელის ფანჯარას რომ გამოჩნდეს პროგრამის მსვლელობა 
            ExcelWorkBook = ExcelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);  // იქმნება ახალი ექსელის workbook რომელშიც გვაქ ერთი შიტი და ამ შიტს ვიყენებთ სამომავლოდ 

            try
            {
                ExcelWorkSheet = ExcelWorkBook.Worksheets[1]; // აქ ვირჩევთ სამუშაო შიტს ვორკბუკიდან(ექსელიდან) ჩვენ შემთხვევაში ერთია და მაგიტომ გვაქ Worksheets[1} ინდექსად 1 


                // ამ კოდის ფრაგმენტებში ივსება სათაურის ველები სხვა სიტყვებით რომ ვთქვათ პირველ row-ში იწერება მნიშვნელობები რამდენი სვეტიც გვაქ (column) 
                ExcelWorkSheet.Cells[1, "A"] = "UNIQ_ID";
                ExcelWorkSheet.Cells[1, "B"] = "Liter_ID";
                ExcelWorkSheet.Cells[1, "C"] = "Photo_N";
                ExcelWorkSheet.Cells[1, "C"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "D"] = "REGION";
                ExcelWorkSheet.Cells[1, "E"] = "Municipality";
                ExcelWorkSheet.Cells[1, "F"] = "Adm_Mun";
                ExcelWorkSheet.Cells[1, "G"] = "City_Town_Village";
                ExcelWorkSheet.Cells[1, "H"] = "Land_Area_Sq_m";
                ExcelWorkSheet.Cells[1, "I"] = "Land_Area_Ha";
                ExcelWorkSheet.Cells[1, "J"] = "shrubbery";
                ExcelWorkSheet.Cells[1, "K"] = "Woody_plant_percent";
                ExcelWorkSheet.Cells[1, "L"] = "Woody_plant_quantity";
                ExcelWorkSheet.Cells[1, "M"] = "VarjisFarti";
                ExcelWorkSheet.Cells[1, "N"] = "woody_plant_species";
                ExcelWorkSheet.Cells[1, "O"] = "In_good_condition";
                ExcelWorkSheet.Cells[1, "P"] = "chopped_down";
                ExcelWorkSheet.Cells[1, "Q"] = "rampike";
                ExcelWorkSheet.Cells[1, "R"] = "species_medium_age";
                ExcelWorkSheet.Cells[1, "S"] = "Note_";
                ExcelWorkSheet.Cells[1, "T"] = "Company";
                ExcelWorkSheet.Cells[1, "U"] = "Field_Operator";
                ExcelWorkSheet.Cells[1, "V"] = "Date_";
                ExcelWorkSheet.Cells[1, "V"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "W"] = "Gis_Operator";
                ExcelWorkSheet.Cells[1, "X"] = "DaTe_1";
                ExcelWorkSheet.Cells[1, "X"].EntireColumn.NumberFormat = "@";//ფორმატტდება დეითის ველის ტიპი ტექსტად
                ExcelWorkSheet.Cells[1, "Y"] = "Overlap_CAD_CODE";
                ExcelWorkSheet.Cells[1, "Z"] = "Owner";
                ExcelWorkSheet.Cells[1, "AA"] = "Legal_person";
                ExcelWorkSheet.Cells[1, "AB"] = "Owners";
                ExcelWorkSheet.Cells[1, "AC"] = "Land_Field_Operator";
                ExcelWorkSheet.Cells[1, "AD"] = "Note1";
                ExcelWorkSheet.Cells[1, "AE"] = "Date_2";
                ExcelWorkSheet.Cells[1, "AF"] = "Land_Gis_Operator";
                ExcelWorkSheet.Cells[1, "AG"] = "Note1_1";
                ExcelWorkSheet.Cells[1, "AH"] = "Date_3";
                ExcelWorkSheet.Cells[1, "AI"] = "CAD_COD";
                ExcelWorkSheet.Cells[1, "AJ"] = "UNIQ_ID_OLD";
                ExcelWorkSheet.Cells[1, "AK"] = "UNIQ_ID_NEW";
                ExcelWorkSheet.Cells[1, "AL"] = "UID";
                ExcelWorkSheet.Cells[1, "AM"] = "ID";

                for (var r = 0; r < qarsafariGroupeds.Count(); r++) // კეთდება ციკლი იმისთვის რო დაიაროს სათითაო ველი და ჩაიწეროს ექსელის შიტში 
                                                                    // R ამ შემთხვევაში ნიშნავს RowNumbers რომ ჩაწერა დაიწყოს მეროე რიგიდან რადგან პირველიში სვეტების სახელები წერია 
                                                                    // ყოველ იტერაციაზე R-ს ერთი ემატება რის გამოც შემდეგ რიგში გადადის ინფორმაციის შევსება 
                {
                    // ციკლის შიგნით იწერება რომელ სვეტში რა ინფორმაცია ჩაიწეროს 

                    ExcelWorkSheet.Cells[r + 2, "A"] = qarsafariGroupeds[r].UniqId;
                    ExcelWorkSheet.Cells[r + 2, "B"] = qarsafariGroupeds[r].LiterId;
                    ExcelWorkSheet.Cells[r + 2, "C"] = qarsafariGroupeds[r].PhotoN;
                    ExcelWorkSheet.Cells[r + 2, "D"] = qarsafariGroupeds[r].Region;
                    ExcelWorkSheet.Cells[r + 2, "E"] = qarsafariGroupeds[r].Municipality;
                    ExcelWorkSheet.Cells[r + 2, "F"] = qarsafariGroupeds[r].AdmMun;
                    ExcelWorkSheet.Cells[r + 2, "G"] = qarsafariGroupeds[r].CityTownVillage;
                    ExcelWorkSheet.Cells[r + 2, "H"] = qarsafariGroupeds[r].LandAreaSqM;
                    ExcelWorkSheet.Cells[r + 2, "I"] = qarsafariGroupeds[r].LandAreaHa;
                    ExcelWorkSheet.Cells[r + 2, "J"] = qarsafariGroupeds[r].Shrubbery;
                    ExcelWorkSheet.Cells[r + 2, "K"] = qarsafariGroupeds[r].WoodyPlantPercent;
                    ExcelWorkSheet.Cells[r + 2, "L"] = qarsafariGroupeds[r].WoodyPlantQuantity;
                    ExcelWorkSheet.Cells[r + 2, "M"] = qarsafariGroupeds[r].VarjisFarti;
                    ExcelWorkSheet.Cells[r + 2, "N"] = qarsafariGroupeds[r].WoodyPlantSpecies;

                    // სადაც სახეობა არ გვიწერია და ხეხილის რაოდენობა იქ იწერება კარგ მდომარეობაში 0 
                    if (qarsafariGroupeds[r].WoodyPlantQuantity == 0)
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "P"] = 0;
                        ExcelWorkSheet.Cells[r + 2, "Q"] = 0;
                    }
                    else
                    {
                        ExcelWorkSheet.Cells[r + 2, "O"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].InGoodCondition), 1);
                        ExcelWorkSheet.Cells[r + 2, "P"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].ChoppedDown), 1);
                        ExcelWorkSheet.Cells[r + 2, "Q"] = Math.Round(Convert.ToDouble(qarsafariGroupeds[r].Rampike), 1);
                    }
                    ExcelWorkSheet.Cells[r + 2, "R"] = qarsafariGroupeds[r].SpeciesMediumAge;
                    ExcelWorkSheet.Cells[r + 2, "S"] = qarsafariGroupeds[r].Note;
                    ExcelWorkSheet.Cells[r + 2, "T"] = qarsafariGroupeds[r].Company;
                    ExcelWorkSheet.Cells[r + 2, "U"] = qarsafariGroupeds[r].FieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "V"] = qarsafariGroupeds[r].Date;
                    ExcelWorkSheet.Cells[r + 2, "W"] = qarsafariGroupeds[r].GisOperator;
                    ExcelWorkSheet.Cells[r + 2, "X"] = qarsafariGroupeds[r].DaTe1;
                    ExcelWorkSheet.Cells[r + 2, "Y"] = qarsafariGroupeds[r].OverlapCadCode;
                    ExcelWorkSheet.Cells[r + 2, "Z"] = qarsafariGroupeds[r].Sakutreba;
                    ExcelWorkSheet.Cells[r + 2, "AA"] = qarsafariGroupeds[r].LegalPerson;
                    ExcelWorkSheet.Cells[r + 2, "AB"] = qarsafariGroupeds[r].Owners;
                    ExcelWorkSheet.Cells[r + 2, "AC"] = qarsafariGroupeds[r].LandFieldOperator;
                    ExcelWorkSheet.Cells[r + 2, "AD"] = qarsafariGroupeds[r].Note1;
                    ExcelWorkSheet.Cells[r + 2, "AE"] = qarsafariGroupeds[r].Date2;
                    ExcelWorkSheet.Cells[r + 2, "AF"] = qarsafariGroupeds[r].LandGisOperator;
                    ExcelWorkSheet.Cells[r + 2, "AG"] = qarsafariGroupeds[r].Note11;
                    ExcelWorkSheet.Cells[r + 2, "AH"] = qarsafariGroupeds[r].Date3;
                    ExcelWorkSheet.Cells[r + 2, "AI"] = qarsafariGroupeds[r].CadCod;
                    ExcelWorkSheet.Cells[r + 2, "AJ"] = qarsafariGroupeds[r].UniqIdOld;
                    ExcelWorkSheet.Cells[r + 2, "AK"] = qarsafariGroupeds[r].UniqIdNew;
                    ExcelWorkSheet.Cells[r + 2, "AL"] = qarsafariGroupeds[r].Uid;
                    ExcelWorkSheet.Cells[r + 2, "AM"] = qarsafariGroupeds[r].Id;
                }
                ExcelWorkBook.Worksheets[1].Name = "პატარა-ექსელი"; // ვარქმევთ ჩვენ მიერ ზევით შექმნილ შიტს სახელს 
                ExcelWorkBook.SaveAs(ExcelDestinationPath + $"\\{ExcelName}.xlsx"); // ვაძლევთ სახელს ექსელის ფაილს ჩვენ მიერ გადმოწოდებული ცვლადის მიხედვით 
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 

                Marshal.ReleaseComObject(ExcelWorkSheet);
                Marshal.ReleaseComObject(ExcelWorkBook);
                Marshal.ReleaseComObject(ExcelApp);

                return new Result<bool>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK
                };

            }
            catch (Exception ex)
            {
                ExcelWorkBook.Close(); // იხურება ექსელის ფაილი ვეღარ მოვახდენთ მასზე ცვლილებას 
                ExcelApp.Quit(); //გამოვდივართ აპლიკაციიდან 
                return new Result<bool>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "მოხდა შეცდომა ექსელში ჩაწერისას!" + ex.Message
                };
            }

        }

        // ამრგვალებს 5 ის ჯერადზე გადაცემულ რიცხვს
        static int RoundToNearest(double number)
        {
            int i = Convert.ToInt32(number);
            return (i % 5) == 0 ? i : (i % 5) >= 2.5 ? i + 5 - (i % 5) : i - (i % 5);
        }

        #endregion
    }
}
