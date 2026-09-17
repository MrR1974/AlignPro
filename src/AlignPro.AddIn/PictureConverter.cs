using System.Drawing;
using System.Windows.Forms;

namespace AlignPro.AddIn
{
    /// <summary>
    /// Converts a <see cref="Image"/> into the <c>IPictureDisp</c> the ribbon's <c>getImage</c>
    /// callback has to return.
    /// </summary>
    /// <remarks>
    /// The conversion lives on <see cref="AxHost"/> as a protected static method, so the only way to
    /// reach it is to derive from it. That is the whole reason this class exists - it is never used as
    /// a control, and is never shown.
    /// </remarks>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter() : base(string.Empty)
        {
        }

        public static stdole.IPictureDisp ToPictureDisp(Image image) =>
            (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
    }
}
