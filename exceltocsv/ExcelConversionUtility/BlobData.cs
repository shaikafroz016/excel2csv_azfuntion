using System.IO;

namespace ExcelConversionUtility
{
	public class BlobInput
	{
		public string BlobName { get; set; }
		public byte[] BlobContent { get; set; }
	}
	public class BlobOutput
	{
		public string BlobName { get; set; }
		public Stream BlobContent { get; set; }
	}
}
