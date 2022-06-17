//using Deloitte.GCP.AI;
using Office.SpireOffice.Enums;
using Office.SpireOffice.Interfaces;
using Microsoft.Extensions.Logging;
using Spire.Pdf.Graphics;
using System;
using System.Drawing;
using System.IO;
using System.Net;

namespace Office.SpireOffice.Services
{
	public class DocumentGenerator : IDocumentGenerator
    {
        private readonly ILogger<IDocumentGenerator> _logger;
        //private readonly IPiiExtractionService _extractionService;

        public DocumentGenerator(ILogger<IDocumentGenerator> logger
            /*IPiiExtractionService extractionService*/)
        {
            _logger = logger;
            //_extractionService = extractionService; 
        }

        public IDocument Generate(string filePath, DocumentFormats format)
        {
            using (var stream = new MemoryStream(File.ReadAllBytes(filePath)))
            {
                stream.Position = 0;
                return Generate(stream, format);
            }
        }

        public IDocument DownloadAndGenerate(string url, DocumentFormats format)
        {
            using (var webClient = new WebClient())
            {
                var data = webClient.DownloadData(url);
                return Generate(data, format);
            }
        }

        public IDocument Generate(Stream stream, DocumentFormats format)
        {
            switch (format)
            {
                //convertors		
                case DocumentFormats.PDF:
                    {
                        var document = new Spire.Pdf.PdfDocument();
                        var doc = new PdfDocument(document, _logger);
                        return doc.ConvertDocument(stream);
                    }
                case DocumentFormats.PPT:
                    {
                        var document = new Spire.Presentation.Presentation();
                        var mydocWrapper = new PowerPointDocument(document, _logger);
                        return mydocWrapper.ConvertDocument(stream);
                    }
                //generators
                case DocumentFormats.XLS:
                case DocumentFormats.XLSM:
                case DocumentFormats.XLSX:
                    {
                        var document = new Spire.Xls.Workbook();
                        var doc = new ExcelDocument(document, _logger);
                        return doc.GenerateDocument(stream);
                    }
                case DocumentFormats.DOC:
                case DocumentFormats.RTF:
                case DocumentFormats.DOCX:
                    {
                        var document = new Spire.Doc.Document();
                        var doc = new WordDocument(document, _logger);
                        return doc.GenerateDocument(stream, format);
                    }
                case DocumentFormats.PPTX:
                    {
                        var document = new Spire.Presentation.Presentation();
                        var doc = new PowerPointDocument(document, _logger);
                        return doc.GenerateDocument(stream);
                    }
                default:
                    {
                        throw new Exception($"{format} is unsupported file format.");
                    }
            }
        }

        public IDocument Generate(byte[] data, DocumentFormats format)
        {
            var stream = new MemoryStream(data);

            return Generate(stream, format);


        }

        public IDocument GeneratePdfFromText(string text)
        {
            var document = new Spire.Pdf.PdfDocument();
            var section = document.Sections.Add();
            var page = section.Pages.Add();
            var format = new PdfStringFormat() { LineSpacing = 20f };
            var textWidget = new PdfTextWidget(text, new PdfFont(PdfFontFamily.Helvetica, 11), PdfBrushes.Black) { StringFormat = format };
            var textLayout = new PdfTextLayout() { Break = PdfLayoutBreakType.FitPage, Layout = PdfLayoutType.Paginate };
            var bounds = new RectangleF(new PointF(0, 0), page.Canvas.ClientSize);
            textWidget.Draw(page, bounds, textLayout);

            var pdfDocument = new PdfDocument(document, _logger);
            return pdfDocument;
        }
    }
}
