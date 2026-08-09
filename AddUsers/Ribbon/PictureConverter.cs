using System.Drawing;
using System.Windows.Forms;

namespace AddUsers
{
    /// <summary>
    /// Uses the protected AxHost conversion helpers to turn a GDI+ Image into
    /// the stdole.IPictureDisp the ribbon loadImage callback must return.
    /// </summary>
    internal sealed class PictureConverter : AxHost
    {
        private PictureConverter()
            : base("59EE46BA-677D-4D20-BF10-8D8067CB8B33")
        {
        }

        public static stdole.IPictureDisp ImageToPictureDisp(Image image)
        {
            return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}
