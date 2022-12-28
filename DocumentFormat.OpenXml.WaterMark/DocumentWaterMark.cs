using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DocumentFormat.OpenXml.WaterMark
{
    public static class DocumentWaterMark
    {
        public static void CreateWaterMark(WordprocessingDocument package, string waterMark)
        {
            var documentPart = package?.MainDocumentPart;
            if (documentPart != null)
            {
                documentPart.DeleteParts(documentPart.HeaderParts);
                HeaderPart header = documentPart.AddNewPart<HeaderPart>();
                DocumentWaterMark.GenerateHeaderPartContent(header);

                string id = documentPart.GetIdOfPart(header);
                ImagePart imagePart = header.AddNewPart<ImagePart>("image/jpeg", "rId999");
                DocumentWaterMark.CreateWaterMarkImage(imagePart, waterMark);

                var sections = documentPart.Document.Body.Elements<SectionProperties>();
                foreach (var section in sections)
                {
                    section.RemoveAllChildren<HeaderReference>();
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = id });
                }
            }
        }

        private static void CreateWaterMarkImage(ImagePart imagePart, string text)
        {
            var date = DateTime.Now.ToString("yyyy-MM-dd");
            using var font = new System.Drawing.Font("Arial", 32, FontStyle.Bold);
            using var brush = new SolidBrush(System.Drawing.Color.Black);
            using var format = new StringFormat(StringFormatFlags.NoClip);
            using var bitmap = new Bitmap(800, 1000);
            var g = Graphics.FromImage(bitmap);
            g.TextRenderingHint = TextRenderingHint.AntiAlias;

            g.ResetTransform();
            g.TranslateTransform(-150, 100);
            g.RotateTransform(-30);
            g.Clear(System.Drawing.Color.Transparent);

            g.DrawString(text, font, brush, new RectangleF(0, 300, 200, 200), format);
            g.DrawString(date, font, brush, new RectangleF(500, 300, 300, 200), format);
            g.DrawString(text, font, brush, new RectangleF(0, 600, 200, 200), format);
            g.DrawString(date, font, brush, new RectangleF(500, 600, 300, 200), format);
            g.Dispose();

            var fileName = Guid.NewGuid().ToString() + ".jpeg";
            bitmap.Save(fileName);

            var fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            byte[] byteArray = new byte[fileStream.Length];
            fileStream.Read(byteArray, 0, byteArray.Length);
            fileStream.Close();

            var memory = new MemoryStream(byteArray);
            imagePart?.FeedData(memory);
            memory.Close();

            //移除文件
            File.Delete(fileName);
        }

        private static void GenerateHeaderPartContent(HeaderPart headerPart)
        {
            var shape = new Shape()
            {
                Id = "WordPictureWatermark75517470",
                Style = "position:absolute;left:0;text-align:left;margin-left:0;margin-top:0;width:415.2pt;height:456.15pt;z-index:-251656192;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin",
                OptionalString = "_x0000_s2051",
                AllowInCell = false,
                Type = "#_x0000_t75"
            };

            ImageData imageData1 = new ImageData()
            {
                Gain = "19661f",
                BlackLevel = "22938f",
                Title = "??",
                RelationshipId = "rId999"
            };

            var header = new Header();
            var paragraph = new Paragraph();
            var run = new Run();
            var picture = new Picture();
            shape.Append(imageData1);
            picture.Append(shape);
            run.Append(picture);
            paragraph.Append(run);
            header.Append(paragraph);
            headerPart.Header = header;
        }

    }
}