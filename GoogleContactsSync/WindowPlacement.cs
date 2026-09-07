using System;
using System.Drawing;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    internal sealed class WindowPlacement
    {
        public int Version { get; set; } = 1;
        public int X { get; set; }
        public int Y { get; set; }
        public int Width { get; set; }
        public int Height { get; set; }
        public bool Maximized { get; set; }

        internal bool IsValid => Version == 1 && Width >= 100 && Width <= 32767 &&
            Height >= 100 && Height <= 32767 && Math.Abs((long)X) < 100000 && Math.Abs((long)Y) < 100000;
        internal Rectangle Bounds => new Rectangle(X, Y, Width, Height);

        internal static WindowPlacement FromBounds(Rectangle bounds, bool maximized) => new WindowPlacement
        {
            X = bounds.X, Y = bounds.Y, Width = bounds.Width, Height = bounds.Height, Maximized = maximized
        };

        internal static Rectangle Fit(Rectangle requested, Rectangle workArea, Size minimum)
        {
            var width = Math.Min(workArea.Width, Math.Max(minimum.Width, requested.Width));
            var height = Math.Min(workArea.Height, Math.Max(minimum.Height, requested.Height));
            return new Rectangle(
                Math.Max(workArea.Left, Math.Min(requested.X, workArea.Right - width)),
                Math.Max(workArea.Top, Math.Min(requested.Y, workArea.Bottom - height)), width, height);
        }

        internal static Rectangle DefaultBounds(Rectangle workArea, Size preferred, Size minimum)
        {
            var width = Math.Min(preferred.Width, workArea.Width * 9 / 10);
            var height = Math.Min(preferred.Height, workArea.Height * 9 / 10);
            return Fit(new Rectangle(workArea.Left + (workArea.Width - width) / 2,
                workArea.Top + (workArea.Height - height) / 2, width, height), workArea, minimum);
        }
    }

    // Keep normal bounds independently of minimized/maximized native window bounds.
    // Persistence and screen lookup are injected so tests never alter user settings.
    internal sealed class WindowPlacementManager : IDisposable
    {
        private readonly Form form;
        private readonly Action<WindowPlacement> save;
        private readonly Func<Rectangle, Rectangle> workingArea;
        private readonly Size minimum;
        private WindowPlacement placement;
        private bool restoring;
        private bool shown;

        internal WindowPlacementManager(Form form, WindowPlacement saved, Action<WindowPlacement> save,
            Func<Rectangle, Rectangle> workingArea = null)
        {
            this.form = form;
            this.save = save;
            this.workingArea = workingArea ?? (bounds => Screen.FromRectangle(bounds).WorkingArea);
            minimum = form.MinimumSize;
            placement = saved?.IsValid == true ? saved : null;
            form.LocationChanged += Capture;
            form.Resize += Capture;
            form.ResizeEnd += SaveOnResizeEnd;
        }

        internal void RestoreForShow()
        {
            if (shown && form.Visible && form.WindowState != FormWindowState.Minimized)
                return;
            restoring = true;
            try
            {
                var preferred = form.RestoreBounds.Size;
                if (preferred.Width <= 0 || preferred.Height <= 0) preferred = form.Size;
                var area = workingArea(placement?.Bounds ?? new Rectangle(Cursor.Position, Size.Empty));
                var bounds = placement == null ? WindowPlacement.DefaultBounds(area, preferred, minimum)
                    : WindowPlacement.Fit(placement.Bounds, area, minimum);
                placement = WindowPlacement.FromBounds(bounds, placement?.Maximized == true);
                form.StartPosition = FormStartPosition.Manual;
                form.MinimumSize = new Size(Math.Min(minimum.Width, area.Width), Math.Min(minimum.Height, area.Height));
                form.WindowState = FormWindowState.Normal;
                form.Bounds = bounds;
                if (placement.Maximized) form.WindowState = FormWindowState.Maximized;
                shown = true;
            }
            finally { restoring = false; }
        }

        private void Capture(object sender, EventArgs args)
        {
            if (!shown || restoring || !form.Visible || form.WindowState == FormWindowState.Minimized)
                return;
            if (form.WindowState == FormWindowState.Normal)
                placement = WindowPlacement.FromBounds(form.Bounds, false);
            else if (placement != null)
                placement.Maximized = true;
        }

        internal void SaveCurrent()
        {
            Capture(null, EventArgs.Empty);
            if (placement?.IsValid == true)
                save(WindowPlacement.FromBounds(placement.Bounds, placement.Maximized));
        }

        private void SaveOnResizeEnd(object sender, EventArgs args) => SaveCurrent();

        public void Dispose()
        {
            form.LocationChanged -= Capture;
            form.Resize -= Capture;
            form.ResizeEnd -= SaveOnResizeEnd;
        }
    }
}
