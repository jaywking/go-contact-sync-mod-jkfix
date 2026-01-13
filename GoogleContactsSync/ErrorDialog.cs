using Serilog;
using System;
using System.Diagnostics;
using System.Drawing;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class ErrorDialog : Form
    {
        public ErrorDialog()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
        }

        public async Task SetErrorText(Exception ex)
        {
            if (await VersionInformation.IsNewVersionAvailable(CancellationToken.None))
            {
                richTextBoxError.AppendText(Environment.NewLine);
                richTextBoxError.AppendText("NEW VERSION AVAILABLE - ");
                var downloadLink = new LinkLabel
                {
                    Text = "DOWNLOAD NOW",
                    AutoSize = true,
                    LinkColor = Color.FromArgb(0, 102, 204),
                    Location = richTextBoxError.GetPositionFromCharIndex(richTextBoxError.TextLength)
                };
                downloadLink.LinkClicked += OpenDowloadUrl;
                richTextBoxError.Controls.Add(downloadLink);
                richTextBoxError.AppendText(downloadLink.Text);
                richTextBoxError.AppendText(Environment.NewLine);
                richTextBoxError.AppendText(Environment.NewLine);
                AppendTextWithColor("PLEASE UPDATE TO THE LATEST VERSION!" + Environment.NewLine, Color.Firebrick);
            }

            AppendTextWithColor("FIRST CHECK IF THIS ERROR HAS ALREADY BEEN REPORTED!", Color.Firebrick);
            AppendTextWithColor(Environment.NewLine + "IF THE PROBLEM STILL EXISTS WRITE AN ERROR REPORT ", Color.Firebrick);
            var bugsLink = new LinkLabel
            {
                Text = "HERE!",
                AutoSize = true,
                LinkColor = Color.FromArgb(0, 102, 204),
                Location = richTextBoxError.GetPositionFromCharIndex(richTextBoxError.TextLength)
            };
            bugsLink.LinkClicked += OpenBugsUrl;
            richTextBoxError.Controls.Add(bugsLink);

            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText(Environment.NewLine);

            richTextBoxError.AppendText("GCSM VERSION:    " + VersionInformation.GetGCSMVersion().ToString());
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("OUTLOOK VERSION: " + VersionInformation.GetOutlookVersion(Synchronizer.OutlookApplication).ToString() + Environment.NewLine);
            richTextBoxError.AppendText("OS VERSION:      " + VersionInformation.GetWindowsVersion() + Environment.NewLine);
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("ERROR MESAGE:" + Environment.NewLine + Environment.NewLine);
            AppendTextWithColor(ex.Message + Environment.NewLine, Color.Firebrick);
            richTextBoxError.AppendText(Environment.NewLine);
            richTextBoxError.AppendText("ERROR MESAGE STACK TRACE:" + Environment.NewLine + Environment.NewLine);
            if (ex.StackTrace != null)
            {
                AppendTextWithColor(ex.StackTrace, Color.Firebrick);
            }
            else
            {
                AppendTextWithColor("NO STACK TRACE AVAILABLE", Color.Firebrick);
            }

            var message = richTextBoxError.Text.Replace("\n", "\r\n");
            //copy to clipboard
            try
            {
                var thread = new Thread(() => System.Windows.Clipboard.SetDataObject(message, true));
                thread.SetApartmentState(ApartmentState.STA); //Set the thread to STA
                thread.Start();
                thread.Join();
            }
            catch (Exception e)
            {
                Log.Debug("Message couldn't be copied to clipboard: " + e.Message);
            }
        }

        public string ErrorText => richTextBoxError.Text;

        private void AppendTextWithColor(string text, Color color)
        {
            var start = richTextBoxError.TextLength;
            richTextBoxError.AppendText(text);
            var end = richTextBoxError.TextLength;

            // Textbox may transform chars, so (end-start) != text.Length
            richTextBoxError.Select(start, end - start);
            {
                richTextBoxError.SelectionColor = color;
                // could set box.SelectionBackColor, box.SelectionFont too.
            }
            richTextBoxError.SelectionLength = 0; // clear
        }

        private void OpenDowloadUrl(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/projects/googlesyncmod/files/latest/download");
        }

        private void OpenBugsUrl(object sender, EventArgs e)
        {
            Process.Start("https://sourceforge.net/p/googlesyncmod/bugs/?source=navbar");
        }

        private void ErrorDialog_FormClosed(object sender, FormClosedEventArgs e)
        {
            Visible = false;
        }

        private void ErrorDialog_Load(object sender, EventArgs e)
        {
            TopMost = true;
        }
    }
}
