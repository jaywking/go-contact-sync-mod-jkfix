using System.Drawing;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class DeleteTooManyPropertiesForm : GoContactSyncModForm
    {
        public DeleteTooManyPropertiesForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
        }

        private void DeleteTooManyPropertiesForm_Load(object sender, System.EventArgs e)
        {
            TopMost = true;

            LoadStateFromSettings();
        }

        private void DeleteTooManyPropertiesForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveStateToSettings();
        }
    }
}
