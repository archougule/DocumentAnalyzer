using Office.SpireOffice.Enums;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Office.SpireOffice.Interfaces
{
	public interface IDocument:IDisposable
	{
		/// <summary>
		/// Extracts text from document.
		/// </summary>
		/// <returns>Documents text.</returns>
		string ExtractText(bool fullreduction = false);

        /// <summary>
        /// Replaces text in the document.
        /// </summary>
        /// <param name="replacementData">Reference to the dictionary for replacement.</param>
        /// <param name="fullreduction"></param>
        /// <param name="replacementRegexPattern">Reference to the replacement regex pattern.</param>
		
        void ReplaceText(Dictionary<string, string> replacementData, bool fullreduction, string replacementRegexPattern = null);

		/// <summary>
		/// Extracts all images from document.
		/// </summary>
		/// <returns>Reference to the images list.</returns>
		List<Image> ExtractImages(Dictionary<string,string> piiData, Image img = null, bool fullreduction = false);

		/// <summary>
		/// GenerateDocument from stream.
		/// </summary>
		/// <returns>Document</returns>
		IDocument GenerateDocument(Stream stream, DocumentFormats? format = null);

		/// <summary>
		/// GenerateDocument from stream.
		/// </summary>
		/// <returns>Document</returns>
		IDocument ConvertDocument(Stream stream);

		/// <summary>
		/// Deletes image from document.
		/// </summary>
		/// <param name="image">Reference to the image to detele.</param>
		void DeleteImage(Image image);

		/// <summary>
		/// Replaces image in the document.
		/// </summary>
		/// <param name="oldImage">Reference to the old image.</param>
		/// <param name="newImage">Reference to the new image.</param>
		void ReplaceImage(Image oldImage, Image newImage, List<string> hashes);

		/// <summary>
		/// Deletes all watermarks in the document.
		/// </summary>
		void DeleteAllWatermarks();

		/// <summary>
		/// Converts Document to bytes array.
		/// </summary>
		/// <returns>Bytes array.</returns>
		byte[] ToArray(DocumentFormats? format = null);

		/// <summary>
		/// Saves document.
		/// </summary>
		/// <param name="filePath">File path.</param>
		void Save(string filePath, int? format = null);
		new void Dispose();
	}
}
