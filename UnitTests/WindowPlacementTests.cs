using NUnit.Framework;
using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace GoContactSyncMod.UnitTests
{
    [TestFixture]
    [Apartment(ApartmentState.STA)]
    [NonParallelizable]
    public class WindowPlacementTests
    {
        private static readonly Rectangle Area = new Rectangle(0, 0, 1920, 1040);
        private static readonly Rectangle Normal = new Rectangle(120, 90, 900, 650);
        private static Form TestForm() => new Form
        {
            ShowInTaskbar = false, Opacity = 0, StartPosition = FormStartPosition.Manual,
            Bounds = Normal, MinimumSize = new Size(620, 519)
        };

        [Test]
        public void DefaultFitsScreenInsteadOfUsingOversizedDesignerDimensions()
        {
            var bounds = WindowPlacement.DefaultBounds(Area, new Size(4000, 2400), new Size(620, 519));
            Assert.That(bounds.Size, Is.EqualTo(new Size(1728, 936)));
            Assert.That(Area.Contains(bounds), Is.True);
        }

        [TestCase(2400, 1200, 1800, 1000)]
        [TestCase(-2000, -1000, 900, 650)]
        [TestCase(0, 0, 4000, 3000)]
        public void MissingMonitorAndOversizedWindowStayInsideCurrentWorkArea(int x, int y, int width, int height)
        {
            var fit = WindowPlacement.Fit(new Rectangle(x, y, width, height), Area, new Size(620, 519));
            Assert.That(Area.Contains(fit), Is.True);
        }

        [Test]
        public void NegativeCoordinatesOnExistingMonitorArePreserved()
        {
            var leftScreen = new Rectangle(-1920, 0, 1920, 1040);
            var saved = new Rectangle(-1800, 90, 900, 650);
            Assert.That(WindowPlacement.Fit(saved, leftScreen, new Size(620, 519)), Is.EqualTo(saved));
        }

        [Test]
        public void ScreenSmallerThanMinimumStillKeepsWindowOnScreen()
        {
            var small = new Rectangle(0, 0, 800, 480);
            Assert.That(WindowPlacement.Fit(Normal, small, new Size(900, 700)), Is.EqualTo(small));
        }

        [Test]
        public void InvalidSavedCoordinatesAndVersionsAreRejected()
        {
            Assert.That(new WindowPlacement { Width = int.MaxValue, Height = 600 }.IsValid, Is.False);
            Assert.That(new WindowPlacement { X = int.MinValue, Width = 800, Height = 600 }.IsValid, Is.False);
            Assert.That(new WindowPlacement { Version = 2, Width = 800, Height = 600 }.IsValid, Is.False);
        }

        [Test]
        public void ResizeAndMoveSurviveRestart()
        {
            WindowPlacement persisted = null;
            var changed = new Rectangle(180, 140, 980, 690);
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, WindowPlacement.FromBounds(Normal, false),
                value => persisted = value, _ => Area))
            {
                manager.RestoreForShow();
                form.Show();
                form.Bounds = changed;
                manager.SaveCurrent();
                form.Close();
            }
            Assert.That(persisted.Bounds, Is.EqualTo(changed));
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, persisted, _ => { }, _ => Area))
            {
                manager.RestoreForShow();
                Assert.That(form.Bounds, Is.EqualTo(changed));
                Assert.That(form.WindowState, Is.EqualTo(FormWindowState.Normal));
            }
        }

        [Test]
        public void MaximizedStateSurvivesTrayAndRestartWithoutLosingNormalBounds()
        {
            WindowPlacement persisted = null;
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, WindowPlacement.FromBounds(Normal, false),
                value => persisted = value, _ => Area))
            {
                manager.RestoreForShow();
                form.Show();
                form.WindowState = FormWindowState.Maximized;
                Application.DoEvents();
                manager.SaveCurrent();
                Assert.That(persisted.Maximized, Is.True);
                Assert.That(persisted.Bounds, Is.EqualTo(Normal));
                form.WindowState = FormWindowState.Minimized;
                form.Hide();
                manager.SaveCurrent();
                Assert.That(persisted.Maximized, Is.True);
                Assert.That(persisted.Bounds, Is.EqualTo(Normal));
                manager.RestoreForShow();
                form.Show();
                Assert.That(form.WindowState, Is.EqualTo(FormWindowState.Maximized));
                form.Close();
            }
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, persisted, _ => { }, _ => Area))
            {
                manager.RestoreForShow();
                Assert.That(form.WindowState, Is.EqualTo(FormWindowState.Maximized));
                form.Show();
                form.WindowState = FormWindowState.Normal;
                Application.DoEvents();
                Assert.That(form.Bounds, Is.EqualTo(Normal));
                form.Close();
            }
        }

        [Test]
        public void ActivatingVisibleWindowDoesNotResetCurrentSize()
        {
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, WindowPlacement.FromBounds(Normal, false), _ => { }, _ => Area))
            {
                manager.RestoreForShow();
                form.Show();
                var changed = new Rectangle(200, 150, 1000, 700);
                form.Bounds = changed;
                manager.RestoreForShow();
                Assert.That(form.Bounds, Is.EqualTo(changed));
                form.Close();
            }
        }

        [Test]
        public void TrayReopenRechecksMonitorWorkArea()
        {
            var area = Area;
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form,
                WindowPlacement.FromBounds(new Rectangle(900, 90, 900, 650), false), _ => { }, _ => area))
            {
                manager.RestoreForShow();
                form.Show();
                manager.SaveCurrent();
                form.WindowState = FormWindowState.Minimized;
                form.Hide();
                area = new Rectangle(0, 0, 1024, 768);
                manager.RestoreForShow();
                Assert.That(area.Contains(form.Bounds), Is.True);
                form.Close();
            }
        }

        [Test]
        public void HiddenStartupDoesNotReplaceSavedPlacementWithMinimizedBounds()
        {
            WindowPlacement persisted = null;
            using (var form = TestForm())
            using (var manager = new WindowPlacementManager(form, WindowPlacement.FromBounds(Normal, true),
                value => persisted = value, _ => Area))
            {
                form.WindowState = FormWindowState.Minimized;
                manager.SaveCurrent();
                Assert.That(persisted.Bounds, Is.EqualTo(Normal));
                Assert.That(persisted.Maximized, Is.True);
            }
        }
    }
}
