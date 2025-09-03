using GeographicDynamicWebAPI.Wrappers;
using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using GeographicDynamic_DAL.Interface;
using System.IO;
using System.Drawing;
using Font = System.Drawing.Font;

namespace GeographicDynamic_DAL.Repository
{
    public class AeroGadagebaRepository : IAero
    {
        public Result<List<AeroRecord>> ExcelisWakiTxvaAero()
        {
            List<AeroRecord> ExcelInfo = new List<AeroRecord>();
            Application xlApp = new Application();
            Workbook xlWorkbook = null;

            try
            {
                string ExcelPath = @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\aero.xlsx";
                xlWorkbook = xlApp.Workbooks.Open(ExcelPath);
                _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;

                for (int i = 2; i <= rowCount; i++) // assuming first row is headers
                {
                    var objectID = xlRange.Cells[i, 1]?.Value2?.ToString();
                    var cellDate = xlRange.Cells[i, 2]?.Value2;
                    DateTime? dataTaken = null;

                    if (cellDate != null)
                    {
                        if (cellDate is double d)
                            dataTaken = DateTime.FromOADate(d);
                        else if (cellDate is string s && DateTime.TryParse(s, out var parsed))
                            dataTaken = parsed;
                    }

                    var xValue = xlRange.Cells[i, 3]?.Value2?.ToString();
                    var yValue = xlRange.Cells[i, 4]?.Value2?.ToString();

                    ExcelInfo.Add(new AeroRecord
                    {
                        ObjectID = objectID,
                        DataTaken = dataTaken,
                        X = xValue,
                        Y = yValue
                    });
                }

                // ===== Organize images & write coordinates =====
                OrganizeImagesAndWriteCoordinates(ExcelInfo, @"D:\Projects\2025\QarsaffariDatvlebi\TestAero\images");

                return new Result<List<AeroRecord>>
                {
                    Success = true,
                    StatusCode = System.Net.HttpStatusCode.OK,
                };
            }
            catch (Exception ex)
            {
                return new Result<List<AeroRecord>>
                {
                    Success = false,
                    StatusCode = System.Net.HttpStatusCode.BadGateway,
                    Message = "ექსელის წაკითხვა ვერ მოხერხდა: " + ex.Message
                };
            }
            finally
            {
                xlWorkbook?.Close(false);
                xlApp.Quit();
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
            }
        }

        // ===== Combined organize & write coordinates method =====
        private void OrganizeImagesAndWriteCoordinates(List<AeroRecord> excelData, string imagePath)
        {
            // Step 1: Move images into ObjectID folders based on DataTaken ±5 minutes
            string[] files = Directory.GetFiles(imagePath, "*.*", SearchOption.AllDirectories);

            foreach (var record in excelData)
            {
                if (record.DataTaken == null) continue;
                DateTime targetTime = record.DataTaken.Value;

                foreach (var file in files)
                {
                    try
                    {
                        DateTime? photoDate = GetPhotoDateTaken(file);
                        if (photoDate == null) continue;

                        if (Math.Abs((photoDate.Value - targetTime).TotalMinutes) <= 5)
                        {
                            string targetFolder = Path.Combine(imagePath, record.ObjectID);
                            if (!Directory.Exists(targetFolder))
                                Directory.CreateDirectory(targetFolder);

                            string destFile = Path.Combine(targetFolder, Path.GetFileName(file));
                            File.Move(file, destFile, true);
                        }
                    }
                    catch(Exception ex)
                    {
                        continue;
                    }
                }
            }

            // Step 2: Iterate folders and write coordinates
            string[] objectFolders = Directory.GetDirectories(imagePath);

            foreach (var folder in objectFolders)
            {
                string folderName = Path.GetFileName(folder);
                string[] images = Directory.GetFiles(folder);

                foreach (var imagePathFile in images)
                {
                    try
                    {
                        DateTime? photoDate = GetPhotoDateTaken(imagePathFile);
                        if (photoDate == null) continue;

                        AeroRecord match = excelData.Find(r =>
                            r.ObjectID == folderName &&
                            r.DataTaken.HasValue &&
                            Math.Abs((r.DataTaken.Value - photoDate.Value).TotalMinutes) <= 5);

                        if (match != null)
                        {
                            WriteCoordinates(imagePathFile, match.X, match.Y);
                        }
                    }
                    catch(Exception ex)
                    {
                        continue;
                    }
                }
            }
        }

        private DateTime? GetPhotoDateTaken(string path)
        {
            try
            {
                using (var img = Image.FromFile(path))
                {
                    const int PropertyTagDateTaken = 36867;

                    if (img.PropertyIdList.Contains(PropertyTagDateTaken))
                    {
                        var prop = img.GetPropertyItem(PropertyTagDateTaken);
                        string dateStr = System.Text.Encoding.ASCII.GetString(prop.Value).Trim('\0');
                        if (DateTime.TryParseExact(dateStr, "yyyy:MM:dd HH:mm:ss",
                            System.Globalization.CultureInfo.InvariantCulture,
                            System.Globalization.DateTimeStyles.None, out DateTime dt))
                        {
                            return dt;
                        }
                    }
                }
            }
            catch { }

            // fallback to file last write time
            return File.GetLastWriteTime(path);
        }

        private void WriteCoordinates(string imagePath, string x, string y)
        {
            using (Image original = Image.FromFile(imagePath))
            {
                using (Bitmap img = new Bitmap(original)) // copy to new Bitmap
                {
                    using (Graphics g = Graphics.FromImage(img))
                    {
                        string text = $"X: {x}, Y: {y}";
                        Font font = new Font("Arial", 24, FontStyle.Bold);
                        SizeF textSize = g.MeasureString(text, font);

                        int xPos = img.Width - (int)textSize.Width - 10;
                        int yPos = img.Height - (int)textSize.Height - 10;

                        // Draw outline for readability
                        g.DrawString(text, font, Brushes.White, xPos + 1, yPos + 1);
                        g.DrawString(text, font, Brushes.Black, xPos, yPos);
                    }
                    // Save with proper format
                    img.Save(imagePath, original.RawFormat);
                }
            }
        }
    }

    public class AeroRecord
    {
        public string ObjectID { get; set; }
        public DateTime? DataTaken { get; set; }
        public string X { get; set; }
        public string Y { get; set; }
    }
}
