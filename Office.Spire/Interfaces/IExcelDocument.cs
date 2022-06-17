using System;
using System.Collections.Generic;
using System.Text;

namespace Office.SpireOffice.Interfaces
{
	public interface IExcelDocument: IDocument
	{
		/// <summary>
		/// Extracts text from document.
		/// </summary>
		/// <returns>Documents text.</returns>
		void AnonimizeCell(bool fullreduction);
	}
}
