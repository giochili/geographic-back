using GeographicDynamic_DAL.DTOs.Windbreak;
using GeographicDynamic_DAL.Interface;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Text;
using DotSpatial.Projections;
using ProjNet.CoordinateSystems;
using ProjNet.CoordinateSystems.Transformations;

namespace GeographicDynamic_DAL.Repository
{
    public class TextImageRepository : IWriteOnImage
    {
        public string WriteInfoOnImage(RenamePhotoDTO renamePhotoDTO)
        {
            StringBuilder result = new StringBuilder();

            try
            {
                var directories = Directory.GetDirectories(renamePhotoDTO.FolderPath)
                    .OrderBy(filePath =>
                    {
                        int.TryParse(Path.GetFileName(filePath), out int number);
                        return number;
                    });

                int photocount = renamePhotoDTO.PhotoStartNumber;

                foreach (var folderPath in directories)
                {
                    foreach (var item in Directory.GetDirectories(folderPath))
                    {
                        DirectoryInfo d5 = new DirectoryInfo(item);
                        FileInfo[] infos1 = d5.GetFiles();

                        foreach (FileInfo f6 in infos1)
                        {
                            if (!f6.Name.Contains(".db"))
                            {
                                string newPhotoNamePath = f6.FullName;
                                string outputPath = Path.Combine(Path.GetDirectoryName(newPhotoNamePath), $"Processed_{f6.Name}");

                                try
                                {
                                    using (Image myImage = Image.FromFile(newPhotoNamePath))
                                    {
                                        // Get the date taken
                                        DateTime dtaken = f6.LastWriteTime;
                                        try
                                        {
                                            PropertyItem propItem = myImage.GetPropertyItem(306); // Date Taken
                                            string sdate = Encoding.UTF8.GetString(propItem.Value).Trim();
                                            string firsthalf = sdate.Substring(0, 10).Replace(":", "-");
                                            string secondhalf = sdate.Substring(sdate.IndexOf(" "));
                                            dtaken = DateTime.Parse(firsthalf + secondhalf, CultureInfo.InvariantCulture);
                                        }
                                        catch (Exception)
                                        {
                                            // Ignore if Date Taken property doesn't exist
                                        }

                                        // Get GPS coordinates
                                        double latitude = GetGpsCoordinate(myImage, 2); // Latitude
                                        double longitude = GetGpsCoordinate(myImage, 4); // Longitude
                                                                                         // Get the UTM Zone for the given longitude
                                                                                         // If coordinates are valid, proceed
                                        if (latitude != 0 && longitude != 0)
                                        {
                                            // Get the UTM Zone for the given longitude
                                            int utmZone = (int)Math.Floor((longitude + 180) / 6) + 1;

                                            // Define the WGS84 (GPS) and UTM coordinate systems
                                            var wgs84 = GeographicCoordinateSystem.WGS84;
                                            var utm = ProjectedCoordinateSystem.WGS84_UTM(utmZone, latitude >= 0); // Northern Hemisphere

                                            // Create the coordinate transformation
                                            var transformFactory = new CoordinateTransformationFactory();
                                            var transformation = transformFactory.CreateFromCoordinateSystems(wgs84, utm);

                                            // Convert GPS coordinates to UTM
                                            double[] utmCoordinates = transformation.MathTransform.Transform(new double[] { longitude, latitude });
                                            myImage.Dispose();

                                            // **Write UTM Coordinates onto Image**
                                            //// ამით მხოლოდ კოორდინატები ეწერება 
                                            //WriteCoordinatesOnImage(newPhotoNamePath, utmCoordinates[0], utmCoordinates[1], utmZone);
                                            /////ამით უკვე თარიღიც 
                                            WriteCoordinatesOnImage(newPhotoNamePath, utmCoordinates[0], utmCoordinates[1], utmZone, dtaken);

                                            string photoInfo = $"Photo: {f6.Name}, Date Taken: {dtaken}, GPS: {latitude} ||| {longitude}, UTM: {utmCoordinates[0]}, {utmCoordinates[1]}, Zone {utmZone}";
                                            result.AppendLine(photoInfo);
                                        }
                                        else
                                        {
                                            result.AppendLine($"Photo: {f6.Name} has no GPS data.");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    result.AppendLine($"Error processing {f6.Name}: {ex.Message}");
                                }

                                photocount++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }

            return result.Length > 0 ? result.ToString() : "No photos found.";
        }

        private double GetGpsCoordinate(Image image, int propId)
        {
            try
            {
                PropertyItem propItem = image.GetPropertyItem(propId);
                if (propItem != null)
                {
                    return ConvertGpsBytesToDouble(propItem.Value);
                }
            }
            catch (Exception) { }
            return 0.0; // Return a default value if the coordinate cannot be retrieved
        }
        // Helper method to convert byte array to GPS coordinate
        private double ConvertGpsBytesToDouble(byte[] value)
        {
            if (value.Length != 24) // 3 rational numbers (each 8 bytes)
                throw new ArgumentException("Invalid GPS byte array format.");

            // Degrees
            uint degNumerator = BitConverter.ToUInt32(value, 0);
            uint degDenom = BitConverter.ToUInt32(value, 4);
            double degrees = (double)degNumerator / degDenom;

            // Minutes
            uint minNumerator = BitConverter.ToUInt32(value, 8);
            uint minDenom = BitConverter.ToUInt32(value, 12);
            double minutes = (double)minNumerator / minDenom;

            // Seconds
            uint secNumerator = BitConverter.ToUInt32(value, 16);
            uint secDenom = BitConverter.ToUInt32(value, 20);
            double seconds = (double)secNumerator / secDenom;

            return degrees + (minutes / 60.0) + (seconds / 3600.0);
        }

        #region  ეს კოდი მუშაობს ოღონდ არ წერს თარიღს მხოლოდ აწერს სურათზე კოორდინატებს
        //private static void WriteCoordinatesOnImage(string imagePath, double utmEast, double utmNorth, int utmZone, DateTime dateTaken)
        //{
        //    try
        //    {
        //        if (!File.Exists(imagePath))
        //        {
        //            Console.WriteLine($"File does not exist: {imagePath}");
        //            return;
        //        }

        //        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(imagePath);
        //        string tempPath = Path.Combine(Path.GetDirectoryName(imagePath), $"{fileNameWithoutExtension}_temp.jpeg");

        //        using (Image image = Image.FromFile(imagePath))
        //        {
        //            const int ExifOrientationId = 0x0112;
        //            int originalOrientation = 1;

        //            if (image.PropertyIdList.Contains(ExifOrientationId))
        //            {
        //                var prop = image.GetPropertyItem(ExifOrientationId);
        //                originalOrientation = BitConverter.ToUInt16(prop.Value, 0);

        //                switch (originalOrientation)
        //                {
        //                    case 3:
        //                        image.RotateFlip(RotateFlipType.Rotate180FlipNone);
        //                        break;
        //                    case 6:
        //                        image.RotateFlip(RotateFlipType.Rotate90FlipNone);
        //                        break;
        //                    case 8:
        //                        image.RotateFlip(RotateFlipType.Rotate270FlipNone);
        //                        break;
        //                }
        //            }

        //            using (Graphics graphics = Graphics.FromImage(image))
        //            {
        //                string utmText = $"{utmZone}T , {utmEast:F2}, {utmNorth:F2}";
        //                int fontSize = image.Width / 50;
        //                Font font = new Font("Arial", fontSize, FontStyle.Bold);
        //                Brush textBrush = new SolidBrush(Color.Red);

        //                SizeF textSize = graphics.MeasureString(utmText, font);
        //                float x = 20;
        //                float y = image.Height - textSize.Height - 20;

        //                if (x + textSize.Width > image.Width)
        //                {
        //                    x = image.Width - textSize.Width - 20;
        //                }
        //                if (y + textSize.Height > image.Height)
        //                {
        //                    y = image.Height - textSize.Height - 20;
        //                }

        //                Brush backgroundBrush = new SolidBrush(Color.FromArgb(150, 0, 0, 0));
        //                graphics.FillRectangle(backgroundBrush, x - 10, y - 5, textSize.Width + 20, textSize.Height + 10);
        //                graphics.DrawString(utmText, font, textBrush, new PointF(x, y));
        //            }

        //            switch (originalOrientation)
        //            {
        //                case 3:
        //                    image.RotateFlip(RotateFlipType.Rotate180FlipNone);
        //                    break;
        //                case 6:
        //                    image.RotateFlip(RotateFlipType.Rotate270FlipNone);
        //                    break;
        //                case 8:
        //                    image.RotateFlip(RotateFlipType.Rotate90FlipNone);
        //                    break;
        //            }

        //            // Save edited image to a temporary file
        //            image.Save(tempPath, ImageFormat.Jpeg);
        //        } // Image object is disposed here, releasing file lock

        //        // Ensure the file is writable before deleting
        //        File.SetAttributes(imagePath, FileAttributes.Normal);

        //        // Delete the original file
        //        File.Delete(imagePath);

        //        // Rename the temp file to the original file name
        //        File.Move(tempPath, imagePath);
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Error writing text on image: {ex.Message}");
        //    }
        //}
        #endregion



        #region ეს კოდიც მუშაობს ოღონდ კოორდინატებთან ერთად აწერს თარიღსაც 
        private static void WriteCoordinatesOnImage(string imagePath, double utmEast, double utmNorth, int utmZone, DateTime dateTaken)
        {
            try
            {
                if (!File.Exists(imagePath)) return;

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(imagePath);
                string tempPath = Path.Combine(Path.GetDirectoryName(imagePath), $"{fileNameWithoutExtension}_temp.jpeg");

                using (Image image = Image.FromFile(imagePath))
                {
                    const int ExifOrientationId = 0x0112;
                    int originalOrientation = 1;

                    if (image.PropertyIdList.Contains(ExifOrientationId))
                    {
                        var prop = image.GetPropertyItem(ExifOrientationId);
                        originalOrientation = BitConverter.ToUInt16(prop.Value, 0);

                        switch (originalOrientation)
                        {
                            case 3: image.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                            case 6: image.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                            case 8: image.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
                        }
                    }

                    using (Graphics graphics = Graphics.FromImage(image))
                    {
                        string utmText = $"{utmZone}T , {utmEast:F2}, {utmNorth:F2}";
                        string dateText = dateTaken.ToString("yyyy-MM-dd HH:mm:ss");
                        string fullText = $"{utmText}\nDate: {dateText}";

                        int fontSize = image.Width / 50;
                        Font font = new Font("Arial", fontSize, FontStyle.Bold);
                        Brush textBrush = new SolidBrush(Color.Red);

                        SizeF textSize = graphics.MeasureString(fullText, font);
                        float x = 20;
                        float y = image.Height - textSize.Height - 20;

                        if (x + textSize.Width > image.Width) x = image.Width - textSize.Width - 20;
                        if (y + textSize.Height > image.Height) y = image.Height - textSize.Height - 20;

                        Brush backgroundBrush = new SolidBrush(Color.FromArgb(150, 0, 0, 0));
                        graphics.FillRectangle(backgroundBrush, x - 10, y - 5, textSize.Width + 20, textSize.Height + 10);
                        graphics.DrawString(fullText, font, textBrush, new PointF(x, y));
                    }

                    // Rotate back if needed
                    switch (originalOrientation)
                    {
                        case 3: image.RotateFlip(RotateFlipType.Rotate180FlipNone); break;
                        case 6: image.RotateFlip(RotateFlipType.Rotate270FlipNone); break;
                        case 8: image.RotateFlip(RotateFlipType.Rotate90FlipNone); break;
                    }

                    image.Save(tempPath, ImageFormat.Jpeg);
                }

                File.SetAttributes(imagePath, FileAttributes.Normal);
                File.Delete(imagePath);
                File.Move(tempPath, imagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error writing text on image: {ex.Message}");
            }
        }
        #endregion

    }

}