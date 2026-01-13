using System.Windows.Forms;
using System.Drawing;
using GoContactSyncMod.Properties;

namespace GoContactSyncMod
{
    public class GoContactSyncModForm : Form
    {
        protected void LoadStateFromSettings()
        {
            var cn = GetType().Name;

            var s = (Size)Settings.Default[cn + "_Window_Size"];

            if (s.Width > 0 && s.Height > 0)
            {
                Size = s;
            }

            var wa = Screen.FromControl(this).WorkingArea;

            if (Width > wa.Width)
            {
                Width = wa.Width;
            }

            if (Height > wa.Height)
            {
                Height = wa.Height;
            }

            var p = (Point)Settings.Default[cn + "_Window_Location"];

            if (p.X >= 0 && p.Y >= 0 && p.X <= wa.Width && p.Y <= wa.Height)
            {
                Location = p;
            }
        }

        protected void SaveStateToSettings()
        {
            var cn = GetType().Name;

            Settings.Default[cn + "_Window_Size"] = Size;
            Settings.Default[cn + "_Window_Location"] = Location;

            Settings.Default.Save();
        }
    }
}
