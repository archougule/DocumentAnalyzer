//using Deloitte.GCP.AI.Entities;
using Office.SpireOffice.Enums;
using Office.SpireOffice.Interfaces;
using Microsoft.Extensions.Logging;
using Spire.Presentation;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

using Office.SpireOffice.Interfaces;
using Spire.Presentation;
using Spire.Presentation.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Office.SpireOffice.Extensions;
using System.Text.RegularExpressions;
using Spire.Presentation.Charts;
using Spire.Xls;
using FileFormat = Spire.Presentation.FileFormat;
using Microsoft.Extensions.Logging;
using Office.SpireOffice.Enums;
using SkiaSharp;
using System.Drawing.Imaging;
//
namespace Office.SpireOffice.Services
{
	public class PowerPointDocument :  IPowerPointDocument
	{
        private const string DigitPattern = @"[#$]*[0-9]+(\s?\.?\,?\/?)*[0-9]*(Mb)*(k)*(Km)*(B)*(%)*";
		#region Fields

		protected const string _textReplacementPattern = @"\b{0}\b";
		private Spire.Presentation.Presentation _document;
		private Image _img;
		private bool _fullreduction;
		private string _internalObjectsText;
		private List<ITextFrameProperties> _replacementHandlers = new List<ITextFrameProperties>();

		public List<Image> images { get; private set; }

		private readonly ILogger<IDocumentGenerator> _logger;
		//private readonly IPiiExtractionService _extractionService;
		//private readonly IAnonymizationService _anonymizationService;
		#endregion

		#region Constructors

		public PowerPointDocument(Spire.Presentation.Presentation document, ILogger<IDocumentGenerator> logger)
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
			_logger.LogInformation("DeleteAllWatermarks");
			foreach (ISlide slide in _document.Slides)
			{
				slide.SlideBackground.Fill.FillType = FillFormatType.None;
			}
		}

		public void CleanSlideLayouts()
		{
			try
			{
				_logger.LogInformation("CleanSlideLayouts");
				foreach (ISlide slide in _document.Slides)
				{
					ActiveSlide layout = (ActiveSlide)slide.Layout;
					var counter = layout.Shapes.Count;
					for (int i = 0; i < counter; i++)
					{
						layout.Shapes.RemoveAt(0);
					}
				}
			}
			catch (Exception ex) {
				_logger.LogError(ex.Message+ ex.StackTrace);
			}
		}

		public void CleanSlideMasters()
		{
			try{
			_logger.LogInformation("CleanSlideMasters end");
			foreach (IMasterSlide masterSlide in _document.Masters)
			{
				var counter = masterSlide.Shapes.Count;
				for (int i = 0; i < counter; i++)
				{
					masterSlide.Shapes.RemoveAt(0);
				}
			}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

		public void DeleteImage(Image image)
		{
			try{
			_logger.LogInformation("DeleteImage");
			var imageHashCode = image.ComputeHashCode();
			foreach (ISlide slide in _document.Slides)
			{
				foreach (IShape shape in slide.Shapes.ToArray())
				{
					if (shape is SlidePicture slidePicture)
					{
							Image clone = ConvertImage(slidePicture.PictureFill.Picture.EmbedImage.Image);
							var hashCode = clone.ComputeHashCode();
							clone.Dispose();
							if (string.Equals(imageHashCode, hashCode, StringComparison.OrdinalIgnoreCase))
						{
							slide.Shapes.Remove(shape);
						}
					}
				}
			}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
			_logger.LogInformation("DeleteImage end");
		}

		public void HandleComplexObjects(IShape shape)
		{
			try
			{
				int i = 0;
				List<int> removeIndeces = new List<int>();
				if (shape is GroupShape)
				{
					GroupShape gs = shape as GroupShape;
					foreach (IShape s in gs.Shapes)
					{
						HandleComplexObjects(s);
					}
				}
				else if (shape is IAutoShape)
				{
					IAutoShape ashape = shape as IAutoShape;
					if (ashape.TextFrame != null)
					{
						_internalObjectsText += ashape.TextFrame.Text;
						_replacementHandlers.Add(ashape.TextFrame);
					}
					else
					if (shape is IChart || shape is GraphicFrame || shape is SlidePicture || shape is PictureShape || shape is ITable || shape is IOleObject)
					{
						removeIndeces.Add(i);
					}
					i = i + 1;
				}
				else {
					ProcessShape(removeIndeces, shape);
				}

				foreach (var item in removeIndeces.Distinct().ToList().OrderByDescending(v => v))
				{
					(shape as GroupShape).Shapes.RemoveAt(item);
				}

			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

		public List<Image> ExtractImages(Dictionary<string,string> piiData, Image img = null, bool fullreduction = false)
        {
			_img= img;
			_fullreduction = fullreduction;
			images = new List<Image>();
			try
			{
			_logger.LogInformation("ExtractImages");
			DeleteMetadata();
			
			foreach (ISlide slide in _document.Slides)
			{
                var removeIndeces = new List<int>();
                var removeForm = new List<int>();
                var i = 0;
				foreach (IShape shape in slide.Shapes)
					{
						ProcessShape(  removeIndeces, shape);
					}
                    //var piiExtrationResult = _extractionService.ExtractPiiFromText(_internalObjectsText, _logger);
                    foreach (var handler in _replacementHandlers)
                    {
                        AnonymizeTextInside(piiData, handler, _logger);
                    }

                    foreach (var item in removeIndeces.Distinct().ToList().OrderByDescending(v => v))
                {
                    slide.Shapes.RemoveAt(item);
				}
			}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
			_logger.LogInformation("ExtractImages end");
			return images;
		}

		private void ProcessShape(List<int> removeIndeces, IShape shape)
		{
			try
			{
				int i = 0;
				if (shape is GroupShape)
					HandleComplexObjects(shape);
				if (shape is IAutoShape autoshape)
					if (autoshape.TextFrame != null)
					{
						_internalObjectsText += autoshape.TextFrame.Text;
						_replacementHandlers.Add(autoshape.TextFrame);
					}

				if (shape is SlidePicture slidePicture)
				{
					try
					{
						var image = slidePicture.PictureFill.Picture.EmbedImage.Image;
						Image clone = ConvertImage(image);
						images.Add(clone);
					}
					catch (Exception ex)
					{

						_logger.LogError(ex.Message + ex.StackTrace);

						var oleImg = _document.Images.Append(_img as IImageData);
						(shape as IOleObject).SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImg;
					}
				}

				if (shape is PictureShape pictureShape)
				{
					try
					{
						var image = pictureShape.EmbedImage.Image;
						Image clone = ConvertImage(image);
						images.Add(clone);
					}
					catch (Exception ex)
					{
						_logger.LogError(ex.Message + ex.StackTrace);
					}

				}

				if (_fullreduction)
				{
					if (shape is IChart)
						removeIndeces.Add(i);
					if (shape is ITable)
					{
						try
						{
							var tab = shape as ITable;
							foreach (TableRow row in tab.TableRows)
							{
								foreach (Cell cell in row)
								{
									if (!String.IsNullOrWhiteSpace(cell.TextFrame.Text))
									{
										_internalObjectsText += " " + cell.TextFrame.Text;
										_replacementHandlers.Add(cell.TextFrame);
									}
								}
							}
						}
						catch (Exception ex)
						{
							_logger.LogError(ex.Message + ex.StackTrace);
						}
					}
					var ole = (shape as IOleObject);
					if (shape is IOleObject && ole.ProgId.StartsWith("Excel.Sheet"))
					{

						using (MemoryStream st = new MemoryStream(ole.Data))
						{
							try
							{
								UpdateXLS(ole, st);
							}
							catch
							{
								removeIndeces.Add(i);
							}
						}
					}
				}
				i = i + 1;
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}

		}

		private void UpdateXLS(IOleObject shape, Stream stream)
        {
			try{
			_logger.LogInformation("UpdateXLS st");
			var excelDocument = new Spire.Xls.Workbook();
			excelDocument.LoadFromStream(stream);
            var ex = new ExcelDocument(excelDocument, _logger);
			ex.ExtractText(true);
            (shape).Data = ex.ToArray();
			var newIm= excelDocument.Worksheets[0].ToImage(1,1, excelDocument.Worksheets[0].Rows.Count(), excelDocument.Worksheets[0].Columns.Count());
			var oleImg=_document.Images.Append(newIm); 
			(shape).SubstituteImagePictureFillFormat.Picture.EmbedImage = oleImg;
			_logger.LogInformation("UpdateXLS end");
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

        public string ExtractText(bool fullreduction = false)
        {
			_logger.LogInformation("ExtractText st");
			var content = new StringBuilder();
			try{
			foreach (Spire.Presentation.ISlide slide in _document?.Slides)
			{
				foreach (Spire.Presentation.IShape shape in slide?.Shapes)
				{
					var autoShape = shape as Spire.Presentation.IAutoShape;
					if (autoShape != null && autoShape.TextFrame != null)
					{
						foreach (var paragraph in autoShape.TextFrame.Paragraphs)
						{
							var textParagraph = paragraph as TextParagraph;
							if (textParagraph != null)
							{
								if(fullreduction)
									textParagraph.Text = Regex.Replace(textParagraph.Text, DigitPattern, "X");
								content.AppendLine(textParagraph.Text);
							}
						}
					}
				}
			}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
			var text = content.ToString();
			_logger.LogInformation("ExtractText end");
			return text;
		}
		public void ReplaceImage(Image oldImage, Image newImage, List<string> hashes)
		{
			try{
			_logger.LogInformation("ReplaceImage");
				var stream = new System.IO.MemoryStream();
				newImage.Save(stream,newImage.RawFormat);
				stream.Position = 0;
				IImageData newImageData = _document.Images.Append(stream);
			//var oldImageHashCode = oldImage.ComputeHashCode();
			foreach (ISlide slide in _document.Slides)
			{
				foreach (IShape shape in slide.Shapes)
                {
						try
						{
							ReplaceShapeImage(newImageData, null, shape, hashes);
						}
						catch (Exception ex) { _logger.LogError(ex.Message+ex.StackTrace); }
                }
            }
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
			_logger.LogInformation("ReplaceImage end");
		}

        private void ReplaceShapeImage(IImageData newImageData, string oldImageHashCode, IShape shape, List<string> hashes)
        {
			_logger.LogInformation("ReplaceShapeImage");
			if (shape is SlidePicture slidePicture)
            {
				
				Image clone = ConvertImage(slidePicture.PictureFill.Picture.EmbedImage.Image);
				var hashCode = clone.ComputeHashCode();
				clone.Dispose();
				if (hashes.Contains(hashCode))
                {
                    ((SlidePicture)shape).PictureFill.Picture.EmbedImage = newImageData;
                }
            }
            if (shape is PictureShape pictureShape)
            {
				Image clone = ConvertImage(pictureShape.EmbedImage.Image);
				var hashCode = clone.ComputeHashCode();
				clone.Dispose();
				if (hashes.Contains(hashCode))
				{
					pictureShape.EmbedImage = newImageData;
                }
            }
			_logger.LogInformation("ReplaceShapeImage end");
		}

        public void ReplaceText(Dictionary<string, string> replacementData, bool fullreduction, string replacementRegexPattern)
        {
			try{
			_logger.LogInformation("ReplaceText");
			foreach (ISlide slide in _document?.Slides)
			{
				foreach (IShape shape in slide.Shapes)
				{
					var autoShape = shape as IAutoShape;
					if (autoShape != null && autoShape.TextFrame != null && autoShape.TextFrame.Paragraphs != null)
					{
						foreach (TextParagraph paragraph in autoShape.TextFrame.Paragraphs)
                            {
                                ReplaceInDoc(replacementData, replacementRegexPattern, paragraph.Text);
                            }
                        }

				}
			}
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

        protected static void ReplaceInDoc(Dictionary<string, string> replacementData, string replacementRegexPattern, string text)
        {
            foreach (var keyValue in replacementData)
            {
                if (!String.IsNullOrEmpty(replacementRegexPattern))
                {
                    var pattern = String.Format(replacementRegexPattern, keyValue.Key);
					if (keyValue.Key.ToLower().Contains("ticker") && !text.Contains(@"\(") && !text.Contains(@"\)"))
						text = text.Replace("(", @"\(").Replace(")", @"\)");
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
			byte[] documentData= null;
			try
			//using(var stream = new MemoryStream())
			{
				_logger.LogInformation("ToArray strat");
				var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + "present.pptx");
				_document.SaveToFile(tempPath,FileFormat.Pptx2013);
				//_document.SaveToStream(stream);
				//stream.Position = 0;
				_logger.LogInformation("ToArray end1");
				 documentData = File.ReadAllBytes(tempPath);
				_logger.LogInformation("ToArray end2");
				File.Delete(tempPath);
				_logger.LogInformation("ToArray end3");
				
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
			return documentData;

		}

		public void Save(string filePath, int? format = null)
		{
			try{
			_logger.LogInformation("Save");
			_document.SaveToFile(filePath, FileFormat.Pptx2013);
			_logger.LogInformation("Save fin");
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

		private void DeleteMetadata()
		{
			try{
			_logger.LogInformation("DeleteMetadata");
		//	_document.DocumentProperty.Author = "";
			_document.DocumentProperty.Company = "";
			_document.DocumentProperty.Keywords = "";
			_document.DocumentProperty.Comments = "";
			_document.DocumentProperty.Category = "";
			_document.DocumentProperty.Title = "";
			_document.DocumentProperty.Subject = "";
            _document.DocumentProperty.Category = "";
            _document.DocumentProperty.ContentStatus = "";
            _document.DocumentProperty.Manager = "";
            _document.DocumentProperty.Subject = "";
	//		_document.DocumentProperty.Application = "";
			}
			catch (Exception ex)
			{
				_logger.LogError(ex.Message + ex.StackTrace);
			}
		}

		public IDocument ConvertDocument(Stream stream)
		{
			IDocument gd = null;
			try{
			_logger.LogInformation("ConvertDocument st");
			_document.LoadFromStream(stream, Spire.Presentation.FileFormat.Auto);
			var tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
			_document.SaveToFile(Path.Combine(tempPath, "present.pptx"), Spire.Presentation.FileFormat.Pptx2013);
				Dispose();
				_document = new Spire.Presentation.Presentation();
				var documentData = File.ReadAllBytes(Path.Combine(tempPath, "present.pptx"));
			using (var stream2 = new MemoryStream(documentData))
			{
				gd = GenerateDocument(stream2);
				File.Delete(Path.Combine(tempPath, "present.pptx"));
				_logger.LogInformation("ConvertDocument end");
				
			}
			}
			catch (Exception ex)
			{
				Dispose();
				throw ex;
			}
			return gd;
		}

		public IDocument GenerateDocument(Stream stream, DocumentFormats? format = null)
		{
			try{
			_logger.LogInformation("GenerateDocument st");
			_document.LoadFromStream(stream, FileFormat.Auto);
			_logger.LogInformation("GenerateDocument");
			return new PowerPointDocument(_document,_logger);
			}
			catch (Exception ex)
			{
				Dispose();
				throw ex;
			}
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

		public void AnonymizeTextInside(Dictionary<string,string> piiData, ITextFrameProperties frame, ILogger log)
		{
            var replacementData = piiData.ToDictionary(x => x.Value, x => x.Key);
            log.LogInformation("Text anonymization started.");
            ReplaceInDoc(replacementData, _textReplacementPattern, frame);
            log.LogInformation("Text anonymization finished.");
        }

		protected static void ReplaceInDoc(Dictionary<string, string> replacementData, string replacementRegexPattern, ITextFrameProperties frame)
		{
			foreach (var keyValue in replacementData)
			{
				if (!String.IsNullOrEmpty(replacementRegexPattern))
				{
					var pattern = String.Format(replacementRegexPattern, keyValue.Key);
					if (keyValue.Key.ToLower().Contains("ticker") && !frame.Text.Contains(@"\(") && !frame.Text.Contains(@"\)"))
						frame.Text = frame.Text.Replace("(", @"\(").Replace(")", @"\)");
					frame.Text = Regex.Replace(frame.Text, pattern, keyValue.Value);
				}
				else
				{
					frame.Text = frame.Text.Replace(keyValue.Key, keyValue.Value);
				}
			}
			frame.Text = Regex.Replace(frame.Text, DigitPattern, "X");
		}

		#endregion
	}
}
