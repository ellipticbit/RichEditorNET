using System.Drawing.Imaging;

namespace EllipticBit.RichEditorNET.Formatting
{
    /// <summary>
    /// Specifies the image format to use when embedding images in HTML output.
    /// </summary>
    public enum HtmlImageFormat
    {
        /// <summary>Portable Network Graphics format.</summary>
        Png,
        /// <summary>JPEG format.</summary>
        Jpeg,
        /// <summary>Graphics Interchange Format.</summary>
        Gif,
        /// <summary>Bitmap format.</summary>
        Bmp,
        /// <summary>Tagged Image File Format.</summary>
        Tiff
    }

    internal static class HtmlImageFormatExtensions
    {
        internal static ImageFormat ToImageFormat(this HtmlImageFormat format) {
            switch (format) {
                case HtmlImageFormat.Jpeg: return ImageFormat.Jpeg;
                case HtmlImageFormat.Gif: return ImageFormat.Gif;
                case HtmlImageFormat.Bmp: return ImageFormat.Bmp;
                case HtmlImageFormat.Tiff: return ImageFormat.Tiff;
                default: return ImageFormat.Png;
            }
        }

        internal static string ToMimeType(this HtmlImageFormat format) {
            switch (format) {
                case HtmlImageFormat.Jpeg: return "image/jpeg";
                case HtmlImageFormat.Gif: return "image/gif";
                case HtmlImageFormat.Bmp: return "image/bmp";
                case HtmlImageFormat.Tiff: return "image/tiff";
                default: return "image/png";
            }
        }
    }
}
