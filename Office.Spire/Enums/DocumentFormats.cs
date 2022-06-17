using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;

namespace Office.SpireOffice.Enums
{
	public enum DocumentFormats
	{
		[Description("doc")]
		DOC = 1,
		[Description("rtf")]
		RTF = 2,
		[Description("docx")]
		DOCX = 3,

		[Description("xls")]
		XLS = 4,
		[Description("xlsx")]
		XLSX = 5,
		[Description("xlsm")]
		XLSM = 6,

		[Description("ppt")]
		PPT = 7,
		[Description("pptx")]
		PPTX = 8,

		[Description("pdf")]
		PDF = 9,

		[Description("txt")]
		TXT = 10,
	}
}
