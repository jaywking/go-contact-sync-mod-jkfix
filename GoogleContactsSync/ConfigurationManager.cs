using Microsoft.Win32;
using System;
using System.Drawing;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    public partial class ConfigurationManagerForm : Form
    {
        public string CurrentSyncProfile { get; internal set; }
        internal Synchronizer Synchronizer { get; set; }

        public ConfigurationManagerForm()
        {
            /* Cannot set Font in designer as there is automatic sorting and Font will be set after AutoScaleDimensions
             * This will prevent application to work correctly with high DPI systems. */
            Font = new Font("Verdana", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);

            InitializeComponent();
        }

        public static string AddProfile()
        {
            var vReturn = "";
            using (var AddEditProfile = new AddEditProfileForm("New profile", null))
            {
                if (AddEditProfile.ShowDialog(SettingsForm.Instance) == DialogResult.OK)
                {
                    if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                    {
                        MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "New profile");
                    }
                    else
                    {
                        Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName);
                        vReturn = AddEditProfile.ProfileName;
                    }
                }
            }
            return vReturn;
        }

        private void FillListProfiles()
        {
            var regKeyAppRoot = Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey);

            lbProfiles.Items.Clear();

            foreach (var subKeyName in regKeyAppRoot.GetSubKeyNames())
            {
                lbProfiles.Items.Add(subKeyName);
            }
        }

        //copy all the values
        private static void CopyKey(RegistryKey parent, string keyNameSource, string keyNameDestination)
        {
            var destination = parent.CreateSubKey(keyNameDestination);
            var source = parent.OpenSubKey(keyNameSource);

            foreach (var valueName in source.GetValueNames())
            {
                var objValue = source.GetValue(valueName);
                var valKind = source.GetValueKind(valueName);
                destination.SetValue(valueName, objValue, valKind);
            }
        }

        private void BtClose_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void BtAdd_Click(object sender, EventArgs e)
        {
            AddProfile();
            FillListProfiles();
        }

        private void BtEdit_Click(object sender, EventArgs e)
        {
            if (1 == lbProfiles.CheckedItems.Count)
            {
                using (var AddEditProfile = new AddEditProfileForm("Edit profile", lbProfiles.CheckedItems[0].ToString()))
                {
                    if (AddEditProfile.ShowDialog(SettingsForm.Instance) == DialogResult.OK)
                    {
                        if (null != Registry.CurrentUser.OpenSubKey(SettingsForm.AppRootKey + '\\' + AddEditProfile.ProfileName))
                        {
                            MessageBox.Show("Profile " + AddEditProfile.ProfileName + " exists, try again. ", "Edit profile");
                        }
                        else
                        {
                            CopyKey(Registry.CurrentUser.CreateSubKey(SettingsForm.AppRootKey), lbProfiles.CheckedItems[0].ToString(), AddEditProfile.ProfileName);
                            Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + lbProfiles.CheckedItems[0].ToString());
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Please, select one profile for editing", "Edit profile");
            }

            FillListProfiles();
        }

        private void BtDel_Click(object sender, EventArgs e)
        {
            if (0 >= lbProfiles.CheckedItems.Count)
            {
                MessageBox.Show("You have not selected any profile. Deletion imposible.", "Delete profile");
            }
            else if (DialogResult.Yes == MessageBox.Show("Do you sure to delete selection ?", "Delete profile",
                                                  MessageBoxButtons.YesNo, MessageBoxIcon.Question))
            {
                foreach (var itemChecked in lbProfiles.CheckedItems)
                {
                    var p = itemChecked.ToString();
                    Registry.CurrentUser.DeleteSubKeyTree(SettingsForm.AppRootKey + '\\' + p);
                    if (CurrentSyncProfile == p)
                    {
                        SettingsForm.RevokeAuthentication();
                        if (Synchronizer != null)
                        {
                            Synchronizer.LogoffGoogle();
                        }
                    }
                }
            }

            FillListProfiles();
        }

        private void ConfigurationManagerForm_Load(object sender, EventArgs e)
        {
            FillListProfiles();
        }
    }
}
