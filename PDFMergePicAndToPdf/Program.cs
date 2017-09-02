using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using Ghostscript.NET;
using Ghostscript.NET.Rasterizer;
using iTextSharp.text.pdf;

namespace PDFMergePicAndToPdf
{
    class Program
    {
        private const int DesiredXDpi = 300;
        private const int DesiredYDpi = 300;
        private static GhostscriptVersionInfo GhostScriptVersion => GetGhostscriptVersionInfo();

        static void Main(string[] args)
        {
            string outPdfFileName,
                outImageFileName,
                oriPdfPath = @"W:\Demo.pdf",
                signFileFullPath = @"W:\Sign.png",
                outPdfPath = @"W:\",
                outPdfTempImagePath = @"W:\",
                outPdfImagePath = @"W:\";

            MergeSign(outPdfPath,
                oriPdfPath,
                signFileFullPath, 
                out outPdfFileName);

            PdfToImage(outPdfTempImagePath,
                outPdfImagePath,
                Path.Combine(outPdfPath, outPdfFileName),
                out outImageFileName);

        }

        /// <summary>
        ///     將Pdf檔匯出成圖檔
        /// </summary>
        /// <param name="outPdfImagePath"></param>
        /// <param name="file"></param>
        /// <param name="outImageFileName"></param>
        /// <param name="outPdfTempImagePath"></param>
        private static void PdfToImage(string outPdfTempImagePath, string outPdfImagePath, string file,
            out string outImageFileName)
        {
            var tempFiles = new List<string>();

            using (var rasterizer = new GhostscriptRasterizer())
            {
                rasterizer.Open(file, GhostScriptVersion, false);

                for (var i = 1; i <= rasterizer.PageCount; i++)
                {
                    var fileName = Path.Combine(outPdfTempImagePath, GetUniqueFileName(".png", outPdfTempImagePath));
                    tempFiles.Add(fileName);

                    var img = rasterizer.GetPage(DesiredXDpi, DesiredYDpi, i);

                    if (i == 1)
                        img.Save(fileName, ImageFormat.Png);
                    else
                    {
                        img.RotateFlip(RotateFlipType.Rotate90FlipNone);
                        img.Save(fileName, ImageFormat.Png);
                    }
                }
                rasterizer.Close();
            }

            outImageFileName = GetUniqueFileName(".png", outPdfImagePath);
            MergeImages(tempFiles).Save(Path.Combine(outPdfImagePath, outImageFileName));
        }

        /// <summary>
        ///     將PDF檔與簽名檔合併
        /// </summary>
        /// <param name="outPdfPath">PDF輸出的檔案路徑</param>
        /// <param name="oriPdfFullPath">原始PDF檔案完整路徑</param>
        /// <param name="signFileFullPath">簽名檔檔案完整路徑</param>
        /// <param name="outPdfFileName">合併後輸出的檔案名稱</param>
        private static void MergeSign(string outPdfPath, string oriPdfFullPath, string signFileFullPath,
            out string outPdfFileName)
        {
            outPdfFileName = GetUniqueFileName(".pdf", outPdfPath);

            using (var reader = new PdfReader(oriPdfFullPath))
            using (var stamper = new PdfStamper(reader,
                new FileStream(Path.Combine(outPdfPath, outPdfFileName), FileMode.Create)))
            {
                var pageCount = reader.NumberOfPages;

                for (var i = 1; i <= pageCount; i++)
                {
                    var content = stamper.GetOverContent(i);

                    var image = iTextSharp.text.Image.GetInstance(signFileFullPath);
                    image.ScaleAbsolute(100, 100);
                    image.SetAbsolutePosition(70, 10);
                    content.AddImage(image);
                }
            }
        }

        /// <summary>
        ///     取得唯一的檔案名稱
        /// </summary>
        /// <param name="extenType">檔案類型 （.png、.pdf）</param>
        /// <param name="filePath">檔案路徑</param>
        /// <returns></returns>
        private static string GetUniqueFileName(string extenType, string filePath)
        {
            var fileName = Guid.NewGuid().ToString("N") + extenType;

            while (File.Exists(Path.Combine(filePath, fileName)))
            {
                fileName = Guid.NewGuid().ToString("N") + extenType;
            }

            return fileName;
        }

        /// <summary>
        ///     將多張圖檔合併為一張
        /// </summary>
        /// <param name="files">檔案路徑</param>
        /// <returns></returns>
        private static Bitmap MergeImages(List<string> files)
        {
            //read all images into memory
            var images = new List<Bitmap>();
            Bitmap finalImage = null;

            try
            {
                var width = 0;
                var height = 0;

                foreach (var image in files)
                {
                    //create a Bitmap from the file and add it to the list
                    var bitmap = new Bitmap(image);

                    //update the size of the final bitmap
                    width += bitmap.Width;
                    height = bitmap.Height > height ? bitmap.Height : height;

                    images.Add(bitmap);
                }

                //create a bitmap to hold the combined image
                finalImage = new Bitmap(width, height);

                //get a graphics object from the image so we can draw on it
                using (var g = Graphics.FromImage(finalImage))
                {
                    //set background color
                    g.Clear(Color.Black);

                    //go through each image and draw it on the final image
                    var offset = 0;
                    foreach (var image in images)
                    {
                        g.DrawImage(image,
                            new Rectangle(offset, 0, image.Width, image.Height));
                        offset += image.Width;
                    }
                }

                return finalImage;
            }
            catch (Exception)
            {
                if (finalImage != null)
                    finalImage.Dispose();
                //throw ex;
                throw;
            }
            finally
            {
                //clean up memory
                foreach (var image in images)
                {
                    image.Dispose();
                }
            }
        }

        /// <summary>
        ///     取得 GhostScript Assembly 版本＆資訊
        /// </summary>
        /// <returns></returns>
        private static GhostscriptVersionInfo GetGhostscriptVersionInfo()
        {
            var path = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);

            var vesion = new GhostscriptVersionInfo(new Version(0, 0, 0), path + @"\gsdll32.dll", string.Empty,
                GhostscriptLicense.GPL);

            return vesion;
        }
    }
}
