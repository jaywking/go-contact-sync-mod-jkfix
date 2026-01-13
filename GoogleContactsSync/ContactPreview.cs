using System;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    public partial class ContactPreview : UserControl
    {
        private Collection<CPField> fields;
        public Outlook.ContactItem OutlookContact { get; set; }

        public ContactPreview(Outlook.ContactItem _outlookContact)
        {
            InitializeComponent();
            OutlookContact = _outlookContact;
            InitializeFields();
        }

        private void InitializeFields()
        {
            // TODO: init all non null fields
            fields = new Collection<CPField>();

            var index = 0;
            var height = Font.Height;

            if (OutlookContact.FirstName != null)
            {
                fields.Add(new CPField("First name", OutlookContact.FirstName, new PointF(0, index * height)));
                index++;
            }
            if (OutlookContact.LastName != null)
            {
                fields.Add(new CPField("Last name", OutlookContact.LastName, new PointF(0, index * height)));
                index++;
            }
            if (OutlookContact.Email1Address != null)
            {
                fields.Add(new CPField("Email", ContactPropertiesUtils.GetOutlookEmailAddress1(OutlookContact), new PointF(0, index * height)));
                index++;
            }

            // resize to fit
            Height = (index + 1) * height;
        }

        private void ContactPreview_Paint(object sender, PaintEventArgs e)
        {
            foreach (var field in fields)
            {
                field.Draw(e, Font);
            }
        }
    }

    public class CPField
    {
        public string Name { get; set; }

        public string Value { get; set; }

        public PointF P { get; set; }


        public CPField(string nameVal, string valueVal, PointF pVal)
        {
            Name = nameVal;
            Value = valueVal;
            P = pVal;
        }

        public void Draw(PaintEventArgs e, Font font)
        {
            var str = Name + ": " + Value;
            if (e != null)
            {
                e.Graphics.DrawString(str, font, Brushes.Black, P);
            }
            else
            {
                throw new ArgumentNullException("PaintEventArgs is null");
            }
        }
    }
}
