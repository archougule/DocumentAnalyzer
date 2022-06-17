using Office.SpireOffice.Enums;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Office.SpireOffice.Interfaces
{
	public interface IDocumentGenerator
	{
		/// <summary>
		/// Generates document from file.
		/// </summary>
		/// <param name="filePath">Path to the file.</param>
		/// <param name="format">Document format.</param>
		/// <returns>Reference to the IDocument instance.</returns>
		IDocument Generate(string filePath, DocumentFormats format);

		/// <summary>
		/// Generates document from file.
		/// </summary>
		/// <param name="url">URL to the file.</param>
		/// <param name="format">Document format.</param>
		/// <returns>Reference to the IDocument instance.</returns>
		IDocument DownloadAndGenerate(string url, DocumentFormats format);

		/// <summary>
		/// Generates document from Stream.
		/// </summary>
		/// <param name="stream">Reference to the file stream.</param>
		/// <param name="format">Document format.</param>
		/// <returns>Reference to the IDocument instance.</returns>
		IDocument Generate(Stream stream, DocumentFormats format);

		/// <summary>
		/// Generates document from bytes array.
		/// </summary>
		/// <param name="data">Reference to the bytes array.</param>
		/// <param name="format">Document format.</param>
		/// <returns>Reference to the IDocument instance.</returns>
		IDocument Generate(byte[] data, DocumentFormats format);

		/// <summary>
		/// Generates document from text.
		/// </summary>
		/// <param name="text">Documents text.</param>
		/// <returns>Reference to the IDocument instance.</returns>
		IDocument GeneratePdfFromText(string text);
	}
}
