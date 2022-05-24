using System.IO;
using SkiaSharp;

namespace OfficeOpenXml.Compatibility
{
    public class ImageCompat
    {
        internal static byte[] GetImageAsByteArray(SKImage image)
        {
            var ms = new MemoryStream();
            image.EncodedData.SaveTo(ms);
            return ms.ToArray();
        }
    }
}
