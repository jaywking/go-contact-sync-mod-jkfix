using Google.Apis.Calendar.v3.Data;
using Google.Apis.PeopleService.v1.Data;
using System;
using System.Collections.Generic;
using Event = Google.Apis.Calendar.v3.Data.Event;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace GoContactSyncMod
{
    internal class ConflictResolver : IConflictResolver, IDisposable
    {
        private readonly ConflictResolverForm _form;

        public ConflictResolver()
        {
            _form = new ConflictResolverForm();
        }

        #region IConflictResolver Members

        public ConflictResolution Resolve(ContactMatch match, bool isNewMatch)
        {
            var name = match.ToString();

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these Outlook and Google Contacts \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                    "Both the Outlook Contact and the Google Person \"" + name +
                    "\" have been changed. Choose which you would like to keep.";
            }

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (match.OutlookContact != null)
            {
                var item = match.OutlookContact.GetOriginalItemFromOutlook();
                _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            }

            if (match.GoogleContact != null)
            {
                _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(match.GoogleContact);
            }

            return Resolve();
        }

        public ConflictResolution ResolveDuplicate(OutlookContactInfo outlookContact, List<Person> googleContacts, out Person googleContact)
        {
            var name = outlookContact.ToString();

            _form.messageLabel.Text =
                     "There are multiple Google Contacts (" + googleContacts.Count + ") matching unique properties for Outlook Contact \"" + name +
                     "\". Please choose from the combobox below the Google Person you would like to match with Outlook and if you want to keep the Google or Outlook properties of the selected contact.";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;

            var item = outlookContact.GetOriginalItemFromOutlook();

            _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            _form.GoogleComboBox.DataSource = googleContacts;
            //_form.GoogleComboBox.DisplayMember = "UniqueName";//ToDo: Define unique name
            _form.GoogleComboBox.Visible = true;
            _form.AllCheckBox.Visible = false;
            _form.skip.Text = "Keep both";

            var res = Resolve();
            googleContact = _form.GoogleComboBox.SelectedItem as Person;

            return res;
        }

        public DeleteResolution ResolveDelete(OutlookContactInfo outlookContact)
        {
            var name = outlookContact.ToString();

            _form.Text = "Google Person deleted";
            _form.messageLabel.Text =
                "Google Person \"" + name +
                "\" doesn't exist anymore. Do you want to delete it also on Outlook side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            var item = outlookContact.GetOriginalItemFromOutlook();
            _form.OutlookItemTextBox.Text = ContactMatch.GetSummary(item);
            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Person googleContact)
        {
            var name = ContactMatch.GetName(googleContact);

            _form.Text = "Outlook Contact deleted";
            _form.messageLabel.Text =
                "Outlook Contact \"" + name +
                "\" doesn't exist anymore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = ContactMatch.GetSummary(googleContact);
            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        private ConflictResolution Resolve()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.Ignore:
                    // skip
                    return _form.AllCheckBox.Checked ? ConflictResolution.SkipAlways : ConflictResolution.Skip;
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.GoogleWinsAlways : ConflictResolution.GoogleWins;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? ConflictResolution.OutlookWinsAlways : ConflictResolution.OutlookWins;
                default:
                    return ConflictResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedOutlook()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteGoogleAlways : DeleteResolution.DeleteGoogle;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepGoogleAlways : DeleteResolution.KeepGoogle;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        private DeleteResolution ResolveDeletedGoogle()
        {
            switch (SettingsForm.Instance.ShowConflictDialog(_form))
            {
                case System.Windows.Forms.DialogResult.No:
                    // google wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.DeleteOutlookAlways : DeleteResolution.DeleteOutlook;
                case System.Windows.Forms.DialogResult.Yes:
                    // outlook wins
                    return _form.AllCheckBox.Checked ? DeleteResolution.KeepOutlookAlways : DeleteResolution.KeepOutlook;
                default:
                    return DeleteResolution.Cancel;
            }
        }

        public ConflictResolution Resolve(Outlook.AppointmentItem oa, Google.Apis.Calendar.v3.Data.Event ga, bool isNewMatch)
        {
            var name = string.Empty;

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (oa != null)
            {
                name = oa.ToLogString();
                _form.OutlookItemTextBox.Text += oa.Body;
            }

            if (ga != null)
            {
                name = ga.ToLogString();
                _form.GoogleItemTextBox.Text += ga.Description;
            }

            if (isNewMatch)
            {
                _form.messageLabel.Text =
                    "This is the first time these appointments \"" + name +
                    "\" are synced. Choose which you would like to keep.";
                _form.skip.Text = "Keep both";
            }
            else
            {
                _form.messageLabel.Text =
                "Both the Outlook and Google Appointment \"" + name +
                "\" have been changed. Choose which you would like to keep.";
            }

            return Resolve();
        }

        public ConflictResolution Resolve(string message, Outlook.AppointmentItem oa, Event ga, bool keepOutlook, bool keepGoogle)
        {
            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text = string.Empty;
            if (oa != null)
            {
                _form.OutlookItemTextBox.Text += oa.Body;
            }

            if (ga != null)
            {
                _form.GoogleItemTextBox.Text += ga.Description;
            }

            _form.keepGoogle.Enabled = keepGoogle;
            _form.keepOutlook.Enabled = keepOutlook;
            _form.AllCheckBox.Visible = true;
            _form.messageLabel.Text = message;

            return Resolve();
        }

        public ConflictResolution Resolve(string message, Outlook.AppointmentItem oa, Google.Apis.Calendar.v3.Data.Event ga, Synchronizer sync)
        {
            return Resolve(message, oa, ga, true, false);
        }

        public ConflictResolution Resolve(string message, Google.Apis.Calendar.v3.Data.Event ga, Outlook.AppointmentItem oa, Synchronizer sync)
        {
            return Resolve(message, oa, ga, false, true);
        }

        public DeleteResolution ResolveDelete(Outlook.AppointmentItem oa)
        {
            _form.Text = "Google appointment deleted";
            _form.messageLabel.Text =
                "Google appointment \"" + oa.ToLogString() +
                "\" doesn't exist anymore. Do you want to delete it also on Outlook side?";

            _form.GoogleItemTextBox.Text = string.Empty;
            _form.OutlookItemTextBox.Text += oa.Body;
            _form.keepOutlook.Text = "Keep Outlook";
            _form.keepGoogle.Text = "Delete Outlook";
            _form.skip.Enabled = false;

            return ResolveDeletedGoogle();
        }

        public DeleteResolution ResolveDelete(Event ga)
        {
            _form.Text = "Outlook appointment deleted";
            _form.messageLabel.Text =
                "Outlook appointment \"" + ga.ToLogString() +
                "\" doesn't exist anymore. Do you want to delete it also on Google side?";

            _form.OutlookItemTextBox.Text = string.Empty;
            _form.GoogleItemTextBox.Text += ga.Description;
            _form.keepOutlook.Text = "Keep Google";
            _form.keepGoogle.Text = "Delete Google";
            _form.skip.Enabled = false;

            return ResolveDeletedOutlook();
        }

        public void Dispose()
        {
            _form.Dispose();
        }

        #endregion
    }
}
