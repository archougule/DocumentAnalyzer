using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace Office.SpireOffice.Extensions
{
	public static class ImageExtensions
	{
		public static string ComputeHashCode(this Image image)
		{
			using (var stream = new MemoryStream())
			{
				image.Save(stream, ImageFormat.Png);
				stream.Position = 0;
				using (SHA256 hashAlgorithm = SHA256.Create())
				{
					byte[] data = hashAlgorithm.ComputeHash(stream);
					var sBuilder = new StringBuilder();
					for (int i = 0; i < data.Length; i++)
					{
						sBuilder.Append(data[i].ToString("x2"));
					}
					return sBuilder.ToString();
				}
			}
		}
	}
}
