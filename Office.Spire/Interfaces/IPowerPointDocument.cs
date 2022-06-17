using System;
using System.Collections.Generic;
using System.Text;

namespace Office.SpireOffice.Interfaces
{
	public interface IPowerPointDocument: IDocument
	{
		void CleanSlideLayouts();

		void CleanSlideMasters();
	}
}
