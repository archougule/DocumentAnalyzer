using Office.SpireOffice.Interfaces;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using Office.SpireOffice.Extensions;
using System.Linq;
using System.Text.RegularExpressions;
using System.Reflection;
using Spire.Doc.Collections;
using System.Drawing.Imaging;
using SkiaSharp;
using Microsoft.Extensions.Logging;
using Office.SpireOffice.Enums;
//using Deloitte.GCP.AI.Entities;

namespace Office.SpireOffice.Services
{
    public class WordDocument :  IWordDocument
    {
        private const string DigitPattern = @"[#$]*[0-9]+(\s?\.?\,?\/?)*[0-9]*(Mb)*(k)*(Km)*(B)*(%)*";
        #region Fields

        private Spire.Doc.Document _document;
        private readonly ILogger<IDocumentGenerator> _logger;

        #endregion

        #region Constructors

        public WordDocument(Spire.Doc.Document document, ILogger<IDocumentGenerator> logger)
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
                if(_document!=null)
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
            _document.Watermark = null;
        }

        public void DeleteImage(Image image)
        {
            var imageHashCode = image.ComputeHashCode();
            foreach (Section section in _document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    var documentObjects = paragraph.ChildObjects.OfType<DocumentObject>().ToArray();
                    foreach (DocumentObject docObj in documentObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObj as DocPicture;
                            var clone = Image.FromStream(new MemoryStream(picture.ImageBytes));
                            var hashCode = clone?.ComputeHashCode();
                            if (String.Equals(hashCode, imageHashCode, StringComparison.OrdinalIgnoreCase))
                            {
                                paragraph.ChildObjects.Remove(docObj);
                            }
                        }
                    }
                }
            }
            _logger.LogInformation("DeleteImage end");
        }

        public static T GetPrivateProperty<T>(object obj, string propertyName)
        {
            return (T)obj.GetType()
                          .GetProperty(propertyName, BindingFlags.Instance | BindingFlags.NonPublic)
                          .GetValue(obj);
        }
        public List<Image> ExtractImages(Dictionary<string,string> piiData, Image img, bool fullreduction = false)
        {
            _logger.LogInformation("ExtractImages");
            DeleteMetadata();
            var images = new List<Image>();
            foreach (Section section in _document.Sections)
            {
                if (fullreduction)
                    AnonimazeSection(section);

                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    var removeIndeces = new List<int>();
                    var removeForm = new List<int>();
                    var i = 0;
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObject as DocPicture;
                            try
                            {
                             //   var clone = picture.Image.Clone() as Image;
                                var clone = Image.FromStream(new MemoryStream(picture.ImageBytes));
                                images.Add(clone);
                            }
                            catch (Exception ex)
                            {
                                MemoryStream ms = new MemoryStream();
                                img.Save(ms, img.RawFormat);
                                picture.ReplaceImage(ms.ToArray(), true);
                                _logger.LogError("clone end" + ex.Message + ex.StackTrace);
                                // picture.LoadImage(img);
                            }
                        }

                        if (fullreduction)
                        {
                            if (docObject.DocumentObjectType == DocumentObjectType.OfficeMath)
                                removeForm.Add(i);


                            if (docObject.DocumentObjectType == DocumentObjectType.Shape)
                            {
                                if (GetPrivateProperty<bool>(docObject, "IsChart"))
                                {

                                    removeIndeces.Add(i);
                                }
                            }

                            if (docObject.DocumentObjectType == DocumentObjectType.TextBox && docObject.ChildObjects is BodyRegionCollection)
                            {
                                foreach (var item in (docObject.ChildObjects as BodyRegionCollection))
                                {
                                    if (item is Paragraph)
                                    {
                                        var text = (item as Paragraph).Text;
                                        (item as Paragraph).Text = Regex.Replace(text, DigitPattern, "X");
                                    }
                                }
                            }

                            if (docObject.DocumentObjectType == DocumentObjectType.TextRange)
                            {
                                var text = (docObject as TextRange).Text;
                                (docObject as TextRange).Text = Regex.Replace(text, DigitPattern, "X");
                            }

                            if ((docObject.DocumentObjectType == DocumentObjectType.OleObject) && (docObject as DocOleObject).ObjectType.StartsWith("Excel.Sheet"))
                            {
                                var ole = (docObject as DocOleObject);
                                using (MemoryStream st = new MemoryStream(ole.NativeData))
                                {
                                    try
                                    {
                                        UpdateXLS(ole, st);
                                    }
                                    catch(Exception  ex)
                                    {
                                        _logger.LogError("AnonimazeTable end" + ex.Message + ex.StackTrace);
                                        removeIndeces.Add(i);
                                    }
                                }
                            }
                        }
                        i = i + 1;
                    }

                    foreach (var item in removeIndeces.OrderByDescending(v => v))
                    {
                        DocPicture docPicture = new DocPicture(paragraph.Document);
                        MemoryStream ms = new MemoryStream();
                        img.Save(ms, img.RawFormat);
                        docPicture.ReplaceImage(ms.ToArray(), true);
                       // docPicture.LoadImage(img);

                        paragraph.ChildObjects.RemoveAt(item);
                        paragraph.ChildObjects.Insert(item, docPicture);
                    }
                    foreach (var item in removeForm.OrderByDescending(v => v))
                    {
                        TextRange range = new TextRange(paragraph.Document);
                        range.Text = string.Format("{{Here was formula}}");
                        paragraph.ChildObjects.RemoveAt(item);
                        paragraph.ChildObjects.Insert(item, range);
                    }
                }
            }
            _logger.LogInformation("ExtractImages start");
            return images;
        }

        private void AnonimazeSection(Section section)
        {
            foreach (Table tab in section.Tables)
                AnonimazeTable(tab);
        }

        private void DeleteMetadata()
        {
            _logger.LogInformation("DeleteMetadata");
            //Set the build-in Properties.
            _document.BuiltinDocumentProperties.Title = "";
            _document.BuiltinDocumentProperties.Author = "";
            _document.BuiltinDocumentProperties.Company = "";
            _document.BuiltinDocumentProperties.Keywords = "";
            _document.BuiltinDocumentProperties.Comments = "";
        }

        private void AnonimazeTable(Table tab)
        {
            foreach (TableRow row in tab.Rows)
            {
                foreach (TableCell cell in row.Cells)
                {

                    foreach (var parItem in cell.ChildObjects)
                    {
                        if (parItem is Section section)
                        {
                            AnonimazeSection(parItem as Section);
                        }
                        else
                            try
                            {
                                if (parItem is Paragraph par)
                                {
                                    foreach (var docObject in par.ChildObjects)
                                    {
                                        if (docObject is TextRange)
                                            (docObject as TextRange).Text = Regex.Replace((docObject as TextRange).Text, DigitPattern, "X");
                                    }
                                }
                                else
                                {
                                    if (parItem is Table)
                                    {
                                        AnonimazeTable(parItem as Table);
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                _logger.LogError("AnonimazeTable end"+ ex.Message+ ex.StackTrace);
                            }
                    }
                }
            }
        }

        public string ExtractText(bool fullreduction = false)
        {
            _logger.LogInformation("ExtractText begin");
            var text = _document.GetText();
            _logger.LogInformation("ExtractText end");
            return text;
        }

        public void ReplaceImage(Image oldImage, Image newImage, List<string> hashes)
        {
            //  var oldImageHashCode = oldImage.ComputeHashCode();
            _logger.LogInformation("ReplaceImage start");
            foreach (Section section in _document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject docObj in paragraph.ChildObjects)
                    {
                        if (docObj.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            try
                            {
                                DocPicture picture = docObj as DocPicture;
                                var clone = Image.FromStream(new MemoryStream(picture.ImageBytes));
                                if (hashes.Contains(clone?.ComputeHashCode()))
                                {
                                    MemoryStream ms = new MemoryStream();
                                    newImage.Save(ms, newImage.RawFormat);
                                    picture.ReplaceImage(ms.ToArray(), true);
                                }
                            }
                            catch(Exception ex) {
                                _logger.LogError("REplace"+ex.Message+ex.StackTrace);
                            }
                        }
                    }
                }
            }
            _logger.LogInformation("ReplaceImage end");
        }

        public void ReplaceText(Dictionary<string, string> replacementData, bool fullreduction, string replacementRegexPattern)
        {
            try
            {
                _logger.LogInformation("ReplaceText start");
                foreach (var keyValue in replacementData)
                {
                    _document.Replace(keyValue.Key, keyValue.Value, false, true);

                }
               
                if (fullreduction)
                    _document.Replace(new Regex(DigitPattern), "X");
                _logger.LogInformation("ReplaceText end");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex.Message + ex.StackTrace);
            }
        }


        public byte[] ToArray(DocumentFormats? format = null)
        {
          //  using (var stream = new MemoryStream())
            {
                _logger.LogInformation("ToArray start");
                var tempPath = "";
                if (format.HasValue && format.Value == DocumentFormats.PDF)
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.pdf");
                    _document.SaveToFile(tempPath, Spire.Doc.FileFormat.PDF);
                }
                else
                if (format.HasValue && format.Value == DocumentFormats.DOC)
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.doc");
                    _document.SaveToFile(tempPath, Spire.Doc.FileFormat.Doc);
                }
                else 
                if (format.HasValue && format.Value == DocumentFormats.RTF)
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.rtf");
                    _document.SaveToFile(tempPath, Spire.Doc.FileFormat.Rtf);
                }
                else
                {
                    tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.docx");
                    _document.SaveToFile(tempPath, Spire.Doc.FileFormat.Docx);
                }
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
            _logger.LogInformation("Save start");
            if (format != null)
                _document.SaveToFile(filePath, FileFormat.PDF);
            else
                _document.SaveToFile(filePath);
            _logger.LogInformation("Save end");
        }

        private void UpdateXLS(DocOleObject shape, Stream stream)
        {
            _logger.LogInformation("UpdateXLS");
            var excelDocument = new Spire.Xls.Workbook();
            excelDocument.LoadFromStream(stream);
            var ex = new ExcelDocument(excelDocument, _logger);
            ex.ExtractText(true);
            shape.SetNativeData(ex.ToArray());

            var newIm = excelDocument.Worksheets[0].ToImage(1, 1, excelDocument.Worksheets[0].Rows.Count(), excelDocument.Worksheets[0].Columns.Count());
            shape.OlePicture.LoadImage(newIm);
            _logger.LogInformation("UpdateXLS end");
        }

        public IDocument GenerateDocument(Stream stream, DocumentFormats? format = null)
        {
            try{
            _logger.LogInformation("GenerateDocument start");
                if(format.HasValue && format.Value == DocumentFormats.RTF)
                    _document.LoadFromStream(stream, Spire.Doc.FileFormat.Rtf);
                else
                if (format.HasValue && format.Value == DocumentFormats.DOC)
                    _document.LoadFromStream(stream, Spire.Doc.FileFormat.Doc);
                else
                    _document.LoadFromStream(stream, Spire.Doc.FileFormat.Auto);
                _logger.LogInformation("GenerateDocument end");
            return new WordDocument(_document, _logger);
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
            _document.LoadFromStream(stream, Spire.Doc.FileFormat.Auto);
            var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            _document.SaveToFile(Path.Combine(tempPath, "present.docx"), Spire.Doc.FileFormat.Docx);
            Dispose();
            _document = new Spire.Doc.Document();
            var documentData = File.ReadAllBytes(Path.Combine(tempPath, "present.docx"));
            using (var stream2 = new MemoryStream(documentData))
            {
                var gd = GenerateDocument(stream2);
                File.Delete(Path.Combine(tempPath, "present.docx"));
                _logger.LogInformation("ConvertDocument end");
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
