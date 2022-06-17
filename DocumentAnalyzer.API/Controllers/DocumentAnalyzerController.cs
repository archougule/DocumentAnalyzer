using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Office.SpireOffice;
using System.Net.Http;
using System.Threading;
using Microsoft.Extensions.Logging;
using System.IO;
using Office.SpireOffice.Enums;
using Office.SpireOffice.Interfaces;
using System.Text.RegularExpressions;
using System.Text;
using System.Globalization;
using System.Reflection;
using System.Drawing;
using Office.SpireOffice.Services;
using Office.SpireOffice.Extensions;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace DocumentAnalyzer.API.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DocumentAnalyzerController : ControllerBase
    {
        private const string _textReplacementPattern = @"(?<![\w\d]){0}(?![\w\d])";
        private readonly IDocumentGenerator _documentGenerator;
        public DocumentAnalyzerController(IDocumentGenerator documentGenerator)
        {
            _documentGenerator = documentGenerator;
        }
        // POST api/<DocumentAnalyzerController>
        [HttpPost]
        public void Post(PostBody postBody)
        {
            var documentData = System.IO.File.ReadAllBytes(postBody.filePath);

            var documentFormat = GetDocumentFormat(postBody.filePath);
            var document = _documentGenerator.Generate(documentData, documentFormat);
            var documentOriginalText = document.ExtractText(true);

            Dictionary<string, string> piiDetectionResult = new Dictionary<string, string>();
            piiDetectionResult.Add("[Organization Name]", "IBM");
            piiDetectionResult.Add("[Person]", "John Miller");
            piiDetectionResult.Add("[Location]", "Asia");
            piiDetectionResult.Add("[Email Addresss]", "John.Miller@gmail.com");
            piiDetectionResult.Add("[Phone Number]", "91324343");
            piiDetectionResult.Add("[Date]", "01/01/2022");

            document = AnonymizeDocument(document, piiDetectionResult, true);
            var filePath = postBody.filePath.Insert(postBody.filePath.LastIndexOf('.'), "_edited");
            document.Save(filePath);
        }

        private DocumentFormats GetDocumentFormat(string path)
        {
            var fileName = path.Contains("/") ? path.Substring(path.LastIndexOf("/") + 1) : path;
            var fileExtension = Path.GetExtension(fileName).ToLower();
            var result = fileExtension.TrimStart('.');
            return GetEnumFromString<DocumentFormats>(result);
        }
        private static T GetEnumFromString<T>(string value) where T : struct
        {
            if (Enum.TryParse<T>(value, true, out T enumValue) == false)
            {
                throw new ArgumentException($"The value: {value} does not match a valid enum name or description.");
            }
            return enumValue;
        }
        private IDocument AnonymizeDocument(IDocument document, Dictionary<string, string> piiData, bool fullreduction)
        {
            if (document is IPowerPointDocument)
            {
                HandlePowerPointDocument((IPowerPointDocument)document);
            }
            AnonymizeText(document, piiData, fullreduction);
            AnonymizePictures(piiData,document, fullreduction);
            return document;
        }
        private void HandlePowerPointDocument(IPowerPointDocument powerPointDocument)
        {
            powerPointDocument.CleanSlideLayouts();
            powerPointDocument.CleanSlideMasters();
        }
        private static void AnonymizeText(IDocument document, Dictionary<string, string> piiData, bool fullreduction)
        {
            var replacementData = piiData.ToDictionary(x => x.Value, x => x.Key);
            var redactedKeys = new Dictionary<string, string>();
            foreach (var pair in replacementData)
            {
                var newKey = RemoveDiacritics(pair.Key).Replace(".", " ").Replace(",", " ").Replace(" - ", " ").Replace("&", " ")
                    .Replace("@", " ").Replace(":", " ").Trim(' ');
                newKey = Regex.Replace(newKey, @"\s+", " ");
                if (pair.Key != newKey)
                    redactedKeys.TryAdd(newKey, pair.Value);
            }
            document.ReplaceText(replacementData, fullreduction, _textReplacementPattern);
            document.ReplaceText(redactedKeys, fullreduction, _textReplacementPattern);
        }
        static string RemoveDiacritics(string text)
        {
            var normalizedString = text.Normalize(NormalizationForm.FormD);
            var stringBuilder = new StringBuilder(capacity: normalizedString.Length);

            for (int i = 0; i < normalizedString.Length; i++)
            {
                char c = normalizedString[i];
                var unicodeCategory = CharUnicodeInfo.GetUnicodeCategory(c);
                if (unicodeCategory != UnicodeCategory.NonSpacingMark)
                {
                    stringBuilder.Append(c);
                }
            }

            return stringBuilder
                .ToString()
                .Normalize(NormalizationForm.FormC);
        }
        private void AnonymizePictures(Dictionary<string,string> piiData, IDocument document, bool fullreduction)
        {
            var placeholder = GetImageResource(Assembly.GetExecutingAssembly(), "Resources.anonymization.png");
            var allPictures = document.ExtractImages(piiData,placeholder, fullreduction);
            var hashes = new List<string>();
            foreach (Image picture in allPictures)
            {
                hashes.Add(picture.ComputeHashCode().ToLower());
            }
            try
            {
                document.ReplaceImage(null, placeholder, hashes);
            }
            catch (Exception ex)
            {
            }

        }
        private static Image GetImageResource(Assembly assembly, string resourceName)
        {
            using (Stream stream = assembly.GetManifestResourceStream(assembly.GetName().Name + '.' + resourceName))
            {
                stream.Position = 0;
                var image = new Bitmap(stream);
                return image;
            }
        }
    }

    public class PostBody
    {
        public string filePath { get; set; }
    }
}
