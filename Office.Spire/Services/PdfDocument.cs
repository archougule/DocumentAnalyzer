using Office.SpireOffice.Interfaces;
using Spire.Pdf;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Office.SpireOffice.Extensions;
using Spire.Pdf.Graphics;
using Microsoft.Extensions.Logging;
using Office.SpireOffice.Enums;
//using Deloitte.GCP.AI.Entities;

namespace Office.SpireOffice.Services
{
    public class PdfDocument : IPdfDocument
    {
        #region Fields

        private Spire.Pdf.PdfDocument _document;
        private readonly ILogger<IDocumentGenerator> _logger;

        #endregion

        #region Constructors

        public PdfDocument(Spire.Pdf.PdfDocument document, ILogger<IDocumentGenerator> logger)
        {
            _logger = logger;
            _document = document;
        }

        #endregion

        #region Public Methods

        public void Dispose()
        {
            _logger.LogInformation("Dispose start");


            try
            {
                if (_document != null)
                    _document.Dispose();
                _document = null;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message + ex.StackTrace);
            }

            _logger.LogInformation("Dispose");

        }
        public void DeleteAllWatermarks()
        {
        }

        public void DeleteImage(Image image)
        {
        }

        public List<Image> ExtractImages(Dictionary<string,string> piiData, Image img = null, bool fullreduction = false)
        {
            //var images = new List<Image>();
            //foreach (PdfPageBase page in _document.Pages)
            //{
            //    var pageImages = page.ExtractImages();
            //    if (pageImages?.Any() ?? false)
            //    {
            //        foreach (var image in pageImages)
            //        {
            //          //  var clone = Image.FromStream(new MemoryStream(image. .Copy().Bytes));
            //            var clone = image.Clone() as Image;
            //            images.Add(clone);
            //        }
            //    }
            //}
            //return images;
            return null;
        }

        public string ExtractText(bool fullreduction = false)
        {
            _logger.LogInformation("ExtractText pdf!!");
            StringBuilder content = new StringBuilder();
            foreach (PdfPageBase page in _document.Pages)
            {
                content.Append(page.ExtractText());
            }
            var text = content.ToString();
            return text;
        }

        public void ReplaceImage(Image oldImage, Image newImage, List<string> hashes)
        {
            ////	IImageData newImageData = _document.Images.Append(newImage);
            //var oldImageHashCode = oldImage.ComputeHashCode();
            //foreach (PdfPageBase page in _document.Pages)
            //{

            //    var pageImages = page.ExtractImages();
            //    if (pageImages?.Any() ?? false)
            //    {
            //        for (int i = 0; i < pageImages.Length; i++)
            //        {
            //            var hashCode = pageImages[i].ComputeHashCode();
            //            if (String.Equals(oldImageHashCode, hashCode, StringComparison.OrdinalIgnoreCase))
            //            {
            //                page.ReplaceImage(i, PdfImage.FromImage(newImage));
            //            }
            //        }
            //    }
            //}
        }

        public void ReplaceText(Dictionary<string, string> replacementData, bool fullreduction, string replacementRegexPattern)
        {
            _logger.LogInformation("ReplaceText pdf!!");
            var text = ExtractText();
            foreach (var keyValue in replacementData)
            {
                if (!String.IsNullOrEmpty(replacementRegexPattern))
                {
                    var pattern = String.Format(replacementRegexPattern, keyValue.Key.Replace("(", String.Empty).Replace(")", String.Empty));
                    text = Regex.Replace(text, pattern, keyValue.Value);
                }
                else
                {
                    text = text.Replace(keyValue.Key, keyValue.Value);
                }
            }
        }

        public byte[] ToArray(DocumentFormats? format = null)
        {
            _logger.LogInformation("ToArray pdf!!");
            using (var stream = new MemoryStream())
            {
                _document.SaveToStream(stream);
                return stream.ToArray();
            }
        }

        public void Save(string filePath, int? format = null)
        {
            _logger.LogInformation("Save pdf!!");
            if (format != null)
                _document.SaveToFile(filePath, FileFormat.DOCX);
            else
                _document.SaveToFile(filePath, FileFormat.PDF);
        }

        public IDocument GenerateDocument(Stream stream, DocumentFormats? format = null)
        {
            try{
            _logger.LogInformation("GenerateDocument pdf!!");
            var document = new Spire.Doc.Document();
            document.LoadFromStream(stream, Spire.Doc.FileFormat.Auto);
            return new WordDocument(document,_logger);
            }
            catch (Exception ex)
            {
                Dispose();
                throw ex;
            }
        }

        public IDocument ConvertDocument(Stream stream)
        {
            try{
            _logger.LogInformation("ConvertDocument pdf!!");
            _document.LoadFromStream(stream);
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            _document.SaveToFile(Path.Combine(tempPath, "present.docx"), Spire.Pdf.FileFormat.DOCX);
            Dispose();
            var documentData = File.ReadAllBytes(Path.Combine(tempPath, "present.docx"));
            using (var stream2 = new MemoryStream(documentData))
            {
                var gd = GenerateDocument(stream2);
                File.Delete(Path.Combine(tempPath, "present.docx"));
                return gd;
            }
            }
            catch (Exception ex)
            {
                Dispose();
                throw ex;
            }
        }

        //public List<Image> ExtractImages(Image img = null, List<PiiDetectionResult> piiData = null, bool fullreduction = false)
        //{
        //    throw new NotImplementedException();
        //}

        #endregion
    }
}
