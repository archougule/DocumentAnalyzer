//using Deloitte.GCP.AI.Entities;
using Office.SpireOffice.Enums;
using Office.SpireOffice.Extensions;
using Office.SpireOffice.Interfaces;
using Microsoft.Extensions.Logging;
using SkiaSharp;
using Spire.Xls;
using Spire.Xls.Core;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Office.SpireOffice.Services
{
    public class ExcelDocument : IExcelDocument
    {
        private const string Pattern = @"[#$]*[0-9]+(\s?\.?\,?\/?)*[0-9]*(Mb)*(k)*(Km)*(B)*(%)*";
        #region Fields

        private Spire.Xls.Workbook _document;
        private readonly ILogger<IDocumentGenerator> _logger;

        #endregion

        #region Constructors

        public ExcelDocument(Spire.Xls.Workbook document, ILogger<IDocumentGenerator> logger)
        {
            _logger= logger;
            _document = document;
        }

        #endregion

        #region Public Methods

        public void DeleteAllWatermarks()
        {
            foreach (Worksheet sheet in _document.Worksheets)
            {
                if (sheet.PageSetup.CenterFooterImage != null)
                {
                    sheet.PageSetup.CenterFooter = String.Empty;
                }
                if (sheet.PageSetup.CenterHeaderImage != null)
                {
                    sheet.PageSetup.CenterHeader = String.Empty;
                }
                if (sheet.PageSetup.LeftFooterImage != null)
                {
                    sheet.PageSetup.LeftFooter = String.Empty;
                }
                if (sheet.PageSetup.LeftHeaderImage != null)
                {
                    sheet.PageSetup.LeftHeader = String.Empty;
                }
                if (sheet.PageSetup.RightFooterImage != null)
                {
                    sheet.PageSetup.RightFooter = String.Empty;
                }
                if (sheet.PageSetup.RightHeaderImage != null)
                {
                    sheet.PageSetup.RightHeader = String.Empty;
                }
            }
        }

        public void DeleteImage(Image image)
        {
            var imageHashCode = image.ComputeHashCode();
            foreach (var sheet in _document.Worksheets)
            {
                for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        var clone = Image.FromStream(new MemoryStream(sheet.Pictures[i].Picture.Bytes));
                        if (String.Equals(clone.ComputeHashCode(), imageHashCode, StringComparison.OrdinalIgnoreCase))
                        {
                            sheet.Pictures[i].Remove();
                        }
                        clone.Dispose();
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex.Message+ex.StackTrace);
                    }
                }
            }
        }

        public List<Image> ExtractImages(Dictionary<string,string> piiData, Image img = null, bool fullreduction = false)
        {           
            DeleteMetadata();

            var images = new List<Image>();
            foreach (var sheet in _document?.Worksheets)
            {
                foreach (ExcelPicture picture in sheet?.Pictures)
                {
                    if (fullreduction && picture?.Picture != null)
                    {
                        try
                        {
                            Image clone = ConvertImage(picture.Picture.Copy());
                            
                            images.Add(clone);
                        }
                        catch (Exception ex)
                        {
                            _logger.LogError(ex.Message + ex.StackTrace);
                        }
                    }
                }
            }
            _logger.LogInformation("ExtractImages");
            return images;
        }

        private static Image ConvertImage(SKBitmap bitmap)
        {
            // create an image COPY
            SKImage image = SKImage.FromBitmap(bitmap);
            // encode the image (defaults to PNG)
            SKData encoded = image.Encode();
            image.Dispose();
            // get a stream over the encoded data
            Stream stream = encoded.AsStream();
            stream.Seek(0, SeekOrigin.Begin);
            var clone = Image.FromStream(stream);
            stream.Dispose();
            return clone;
        }

        static void Replace(Regex pattern, CellRange range)
        {
            if (pattern.IsMatch(range.FormulaValue?.ToString() ?? "") || pattern.IsMatch(range.Value2?.ToString() ?? ""))
            {
                if (range.HasFormula)
                {
                    Object value = range.FormulaValue;
                    range.Clear(ExcelClearOptions.ClearContent);
                    var text = value.ToString();
                    text = pattern.Replace(text, "X");
                    range.Value2 = text;
                }
                else if (!string.IsNullOrWhiteSpace(range.Value?.ToString()))
                {
                    //replace text
                    var text = range.Value;
                    text = pattern.Replace(text, "X");
                    range.Value = text;
                }
            }

        }

        public void AnonimyzeCells()
        {
            Regex regex = new Regex(Pattern, RegexOptions.None);
            //Loop through worksheets
            foreach (Worksheet sheet in _document.Worksheets)
            {
                if (sheet.IsEmpty || sheet.Rows?.Length>1000000) continue;
                List<CellRange> listRanges = new List<CellRange>();
                foreach (var row in sheet.Rows)
                {
                    if (row.IsBlank)
                    {
                        continue;
                    }
                    foreach (var cell in row.CellList)
                    {
                        if (cell.IsBlank)
                        {
                            continue;
                        }
                        Replace(regex, cell);
                    }
                }
            }
        }

        public string ExtractText(bool fullreduction = false)
        {
            _logger.LogInformation("ExtractText start");
            var text = new StringBuilder();
            try
            {            
                foreach (var sheet in _document.Worksheets)
                {
                    try
                    {
                        using (var stream = new MemoryStream())
                        {
                            sheet.SaveToStream(stream, "; ");
                            stream.Position = 0;
                            using (var reader = new StreamReader(stream))
                            {
                                string sheetText = reader.ReadToEnd();
                                text.Append(sheetText).AppendLine();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex.Message + ex.StackTrace);
                    }
                }
            }
            catch(Exception ex) {
                Dispose();
                throw ex;
            }
            _logger.LogInformation("ExtractText");
            return text.ToString();
        }

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
                    throw ex;
                }
            
            _logger.LogInformation("Dispose");

        }

        public void AnonimizeCell(bool fullreduction)
        {
            _logger.LogInformation("AnonimizeCell");


                    try
                    {
                        if (fullreduction)
                            AnonimyzeCells();
                     }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex.Message + ex.StackTrace);
                        throw ex;
                    }
            _logger.LogInformation("AnonimizeCell finish");

        }

        public void ReplaceImage(Image oldImage, Image newImage, List<string> hashes)
        {
           // var oldImageHashCode = oldImage.ComputeHashCode();
            foreach (var sheet in _document.Worksheets)
            {
                for (int i = sheet.Pictures.Count - 1; i >= 0; i--)
                {
                    try
                    {
                        using (Image clone = ConvertImage(sheet.Pictures[i].Picture))
                        {
                            //var clone = Image.FromStream(new MemoryStream(sheet.Pictures[i].Picture.Bytes));
                            if (hashes.Contains(clone.ComputeHashCode()))
                            {
                                ExcelPicture picture = sheet.Pictures[i] as ExcelPicture;
                                MemoryStream ms = new MemoryStream();
                                newImage.Save(ms, newImage.RawFormat);
                                picture.Picture = SKBitmap.Decode(ms.ToArray());
                                ms.Dispose();
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        _logger.LogError(ex.Message + ex.StackTrace);
                    }
                }
            }
            _logger.LogInformation("ReplaceImage");
        }

        public void ReplaceText(Dictionary<string, string> replacementData, bool fullreduction, string replacementRegexPattern)
        {
            _logger.LogInformation("ReplaceText start");
            for (int i = 0; i < replacementData.Count; i++)
            {
                try
                {
                    var oldValue = replacementData.ElementAt(i).Key;
                    var newValue = replacementData.ElementAt(i).Value;
                    var cells = _document.FindAllString(oldValue, false, false);
                    if (cells == null) continue;
                    if (String.IsNullOrEmpty(replacementRegexPattern))
                    {
                        foreach (var cell in cells)
                        {
                            if (cell.IsBlank || String.IsNullOrWhiteSpace(cell.Text))
                            {
                                continue;
                            }
                            cell.Text = cell.Text.Replace(oldValue, newValue);
                        }
                    }
                    else
                    {
                        var pattern = String.Format(replacementRegexPattern, oldValue);
                        foreach (var cell in cells)
                        {
                            if (cell.IsBlank || String.IsNullOrWhiteSpace(cell.Text))
                            {
                                continue;
                            }
                            cell.Text = Regex.Replace(cell.Text, pattern, newValue);
                        }
                    }
                }
                catch (Exception ex)
                {
                    _logger.LogError(ex.Message + ex.StackTrace);
                }
            }
        }

        public byte[] ToArray(DocumentFormats? format = null)
        {
            {
                _logger.LogInformation("ToArray strat");
                string tempPath = "";
                if (format.HasValue && format.Value == DocumentFormats.XLS)
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.XLS");
                    _document.SaveToFile(tempPath, Spire.Xls.ExcelVersion.Version97to2003);                
                }else
                if (format.HasValue && format.Value == DocumentFormats.XLSM)
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.XLSm");
                    _document.SaveToFile(tempPath, FileFormat.Xlsm);
                }
                else
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.xlsx");
                    _document.SaveToFile(tempPath);
                }
                //_document.SaveToStream(stream);
                //stream.Position = 0;
                _logger.LogInformation("ToArray end1");
                var documentData = File.ReadAllBytes(tempPath);
                _logger.LogInformation("ToArray end2");
                File.Delete(tempPath);
                _logger.LogInformation("ToArray end3");
                return documentData;
            }
            
        }

        public void Save(string filePath, int? format = null)
        {
            _logger.LogInformation("Save strat");
            _document.SaveToFile(filePath);
            _logger.LogInformation("Save finish");
        }

        private void DeleteMetadata()
        {
            _logger.LogInformation("DeleteMetadata finish");
           // _document.DocumentProperties.Author = "";
        //    _document.DocumentProperties.Company = "";
            _document.DocumentProperties.Keywords = "";
            _document.DocumentProperties.Comments = "";
            _document.DocumentProperties.Category = "";
            _document.DocumentProperties.Title = "";
            _document.DocumentProperties.Subject = "";
    //        _document.DocumentProperties.ApplicationName = "";
            _document.DocumentProperties.Category = "";
    //        _document.DocumentProperties.LastAuthor = "";
            _document.DocumentProperties.Manager = "";
       //     _document.DocumentProperties.RevisionNumber = "";
            _document.DocumentProperties.Subject = "";

            //var cprop = _document.CustomDocumentProperties;
            //for (int i = 0; i < cprop.Count; i++)
            //{
            //    cprop.Remove(cprop[i].Name);

            //}
            _logger.LogInformation("DeleteMetadata finish2");
        }

        public IDocument GenerateDocument(Stream stream, DocumentFormats? format = null)
        {
            try
            {
            _logger.LogInformation("GenerateDocument start");
            _document.LoadFromStream(stream);
                _logger.LogInformation("GenerateDocument finish");
                return new ExcelDocument(_document, _logger);
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
            _logger.LogInformation("ConvertDocument start");
            _document.LoadFromStream(stream);
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            _document.SaveToFile(Path.Combine(tempPath, "present.xlsx"));
            Dispose();
            _document = new Spire.Xls.Workbook();
            var documentData = File.ReadAllBytes(Path.Combine(tempPath, "present.xlsx"));
            using (var stream2 = new MemoryStream(documentData))
            {
                var gd = GenerateDocument(stream2);
                File.Delete(Path.Combine(tempPath, "present.xlsx"));
                _logger.LogInformation("ConvertDocument");
                return gd;
            }
        }
            catch(Exception ex) {
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
