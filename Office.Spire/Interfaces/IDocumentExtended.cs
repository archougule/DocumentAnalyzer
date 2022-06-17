//using Deloitte.GCP.AI.Entities;
using Office.SpireOffice.Enums;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;

namespace Office.SpireOffice.Interfaces
{
   public interface IDocumentExtended : IDisposable
	{
		
		/// <summary>
		/// Extracts all images from document.
		/// </summary>
		/// <returns>Reference to the images list.</returns>
		//List<Image> ExtractImages(Image img = null, List<PiiDetectionResult> piiData = null, bool fullreduction = false);

		new void Dispose();
    
    }
}
