using Google.Apis.PeopleService.v1.Data;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    internal partial class ConflictResolverForm : Form
    {
        public ConflictResolverForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
        }

        private void GoogleComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (GoogleComboBox.SelectedItem != null)
            {
                GoogleItemTextBox.Text = ContactMatch.GetSummary((Person)GoogleComboBox.SelectedItem);
            }
        }

        private void ConflictResolverForm_Shown(object sender, EventArgs e)
        {
            SettingsForm.Instance.ShowBalloonToolTip(Text, messageLabel.Text, ToolTipIcon.Warning, 5000, true);

        }

        private void ConflictResolverForm_Load(object sender, EventArgs e)
        {
            TopMost = true;
        }
    }
}