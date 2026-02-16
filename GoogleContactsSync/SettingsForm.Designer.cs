using System;

namespace GoContactSyncMod
{
    partial class SettingsForm : IDisposable
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Dispose();
            }

            if (sync != null)
            {
                sync.Dispose();
            }
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsForm));
            this.UserName = new System.Windows.Forms.TextBox();
            this.googleEmailLabel = new System.Windows.Forms.Label();
            this.syncButton = new System.Windows.Forms.Button();
            this.syncOptionBox = new System.Windows.Forms.CheckedListBox();
            this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.systemTrayMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.autoSyncInterval = new System.Windows.Forms.NumericUpDown();
            this.autoSyncCheckBox = new System.Windows.Forms.CheckBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btSyncOnlyIdle = new System.Windows.Forms.CheckBox();
            this.reportSyncResultCheckBox = new System.Windows.Forms.CheckBox();
            this.runAtStartupCheckBox = new System.Windows.Forms.CheckBox();
            this.nextSyncLabel = new System.Windows.Forms.Label();
            this.syncTimer = new System.Windows.Forms.Timer(this.components);
            this.syncOptionsGroupBox = new System.Windows.Forms.GroupBox();
            this.IncludePastRemindersCheckBox = new System.Windows.Forms.CheckBox();
            this.SyncPhotosCheckBox = new System.Windows.Forms.CheckBox();
            this.btMonthsFuture = new System.Windows.Forms.CheckBox();
            this.btMonthsPast = new System.Windows.Forms.CheckBox();
            this.SyncRemindersCheckBox = new System.Windows.Forms.CheckBox();
            this.btSyncAppointmentsPrivate = new System.Windows.Forms.CheckBox();
            this.btSyncAppointmentsForceRTF = new System.Windows.Forms.CheckBox();
            this.btSyncContactsForceRTF = new System.Windows.Forms.CheckBox();
            this.futureMonthInterval = new System.Windows.Forms.NumericUpDown();
            this.pastMonthInterval = new System.Windows.Forms.NumericUpDown();
            this.appointmentGoogleFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.labelTimezone = new System.Windows.Forms.Label();
            this.appointmentTimezonesComboBox = new System.Windows.Forms.ComboBox();
            this.btSyncAppointments = new System.Windows.Forms.CheckBox();
            this.appointmentFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.btSyncContacts = new System.Windows.Forms.CheckBox();
            this.btPromptDelete = new System.Windows.Forms.CheckBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btSyncDelete = new System.Windows.Forms.CheckBox();
            this.cmbSyncProfile = new System.Windows.Forms.ComboBox();
            this.contactFoldersComboBox = new System.Windows.Forms.ComboBox();
            this.syncProfileLabel = new System.Windows.Forms.Label();
            this.outlookStartModeLabel = new System.Windows.Forms.Label();
            this.outlookStartModeComboBox = new System.Windows.Forms.ComboBox();
            this.profileBindingLabel = new System.Windows.Forms.Label();
            this.lastSyncLabel = new System.Windows.Forms.Label();
            this.logGroupBox = new System.Windows.Forms.GroupBox();
            this.syncConsole = new System.Windows.Forms.RichTextBox();
            this.googleAccountGroupBox = new System.Windows.Forms.GroupBox();
            this.linkLabelRevokeAuthentication = new System.Windows.Forms.LinkLabel();
            this.chkUseFileAs = new System.Windows.Forms.CheckBox();
            this.proxySettingsLinkLabel = new System.Windows.Forms.LinkLabel();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.resetMatchesLinkLabel = new System.Windows.Forms.LinkLabel();
            this.Donate = new System.Windows.Forms.PictureBox();
            this.pictureBoxExit = new System.Windows.Forms.PictureBox();
            this.settingsGroupBox = new System.Windows.Forms.GroupBox();
            this.actionsTableLayout = new System.Windows.Forms.TableLayoutPanel();
            this.cancelButton = new System.Windows.Forms.Button();
            this.hideButton = new System.Windows.Forms.Button();
            this.MainPanel = new System.Windows.Forms.Panel();
            this.MainSplitter = new System.Windows.Forms.Splitter();
            this.iconTimer = new System.Windows.Forms.Timer(this.components);
            this.NewVersionLinkLabel = new System.Windows.Forms.LinkLabel();
            this.systemTrayMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).BeginInit();
            this.groupBox1.SuspendLayout();
            this.syncOptionsGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.futureMonthInterval)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pastMonthInterval)).BeginInit();
            this.logGroupBox.SuspendLayout();
            this.googleAccountGroupBox.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Donate)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExit)).BeginInit();
            this.settingsGroupBox.SuspendLayout();
            this.actionsTableLayout.SuspendLayout();
            this.MainPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(86, 18);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(417, 20);
            this.UserName.TabIndex = 0;
            this.toolTip.SetToolTip(this.UserName, "Type in your Google Mail User Name (example: username@gmail.com)");
            this.UserName.Validating += new System.ComponentModel.CancelEventHandler(this.UserName_Validating);
            this.UserName.Validated += new System.EventHandler(this.UserName_Validated);
            // 
            // googleEmailLabel
            // 
            this.googleEmailLabel.AutoSize = true;
            this.googleEmailLabel.Location = new System.Drawing.Point(6, 21);
            this.googleEmailLabel.Name = "googleEmailLabel";
            this.googleEmailLabel.Size = new System.Drawing.Size(39, 13);
            this.googleEmailLabel.TabIndex = 0;
            this.googleEmailLabel.Text = "E-&Mail:";
            // 
            // syncButton
            // 
            this.syncButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.syncButton.Location = new System.Drawing.Point(3, 3);
            this.syncButton.Name = "syncButton";
            this.syncButton.Size = new System.Drawing.Size(66, 25);
            this.syncButton.TabIndex = 0;
            this.syncButton.Text = "S&ync";
            this.syncButton.UseVisualStyleBackColor = true;
            this.syncButton.Click += new System.EventHandler(this.SyncButton_Click);
            // 
            // syncOptionBox
            // 
            this.syncOptionBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.syncOptionBox.BackColor = System.Drawing.SystemColors.Control;
            this.syncOptionBox.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.syncOptionBox.CheckOnClick = true;
            this.syncOptionBox.FormattingEnabled = true;
            this.syncOptionBox.IntegralHeight = false;
            this.syncOptionBox.Location = new System.Drawing.Point(6, 184);
            this.syncOptionBox.MinimumSize = new System.Drawing.Size(146, 0);
            this.syncOptionBox.Name = "syncOptionBox";
            this.syncOptionBox.Size = new System.Drawing.Size(154, 117);
            this.syncOptionBox.TabIndex = 8;
            this.toolTip.SetToolTip(this.syncOptionBox, resources.GetString("syncOptionBox.ToolTip"));
            this.syncOptionBox.SelectedIndexChanged += new System.EventHandler(this.SyncOptionBox_SelectedIndexChanged);
            // 
            // notifyIcon
            // 
            this.notifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Warning;
            this.notifyIcon.ContextMenuStrip = this.systemTrayMenu;
            this.notifyIcon.Icon = global::GoContactSyncMod.Properties.Resources.sync;
            this.notifyIcon.Text = "GO Contact Sync Mod";
            this.notifyIcon.Visible = true;
            this.notifyIcon.BalloonTipClicked += new System.EventHandler(this.ToolStripMenuItem1_Click);
            this.notifyIcon.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.NotifyIcon_MouseDoubleClick);
            // 
            // systemTrayMenu
            // 
            this.systemTrayMenu.ImageScalingSize = new System.Drawing.Size(24, 24);
            this.systemTrayMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem4,
            this.toolStripSeparator2,
            this.toolStripMenuItem1,
            this.toolStripMenuItem3,
            this.toolStripSeparator1,
            this.toolStripMenuItem5,
            this.toolStripMenuItem2});
            this.systemTrayMenu.Name = "systemTrayMenu";
            this.systemTrayMenu.Size = new System.Drawing.Size(108, 126);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem4.Text = "Sync";
            this.toolStripMenuItem4.Click += new System.EventHandler(this.ToolStripMenuItem4_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(104, 6);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem1.Text = "Show";
            this.toolStripMenuItem1.Click += new System.EventHandler(this.ToolStripMenuItem1_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem3.Text = "Hide";
            this.toolStripMenuItem3.Click += new System.EventHandler(this.ToolStripMenuItem3_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(104, 6);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem5.Text = "About";
            this.toolStripMenuItem5.Click += new System.EventHandler(this.ToolStripMenuItem5_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(107, 22);
            this.toolStripMenuItem2.Text = "Exit";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.ToolStripMenuItem2_Click);
            // 
            // autoSyncInterval
            // 
            this.autoSyncInterval.Location = new System.Drawing.Point(93, 91);
            this.autoSyncInterval.Maximum = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.autoSyncInterval.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            0});
            this.autoSyncInterval.Name = "autoSyncInterval";
            this.autoSyncInterval.Size = new System.Drawing.Size(42, 20);
            this.autoSyncInterval.TabIndex = 3;
            this.autoSyncInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.autoSyncInterval.Value = new decimal(new int[] {
            2,
            0,
            0,
            0});
            this.autoSyncInterval.ValueChanged += new System.EventHandler(this.AutoSyncInterval_ValueChanged);
            this.autoSyncInterval.Enter += new System.EventHandler(this.AutoSyncInterval_Enter);
            this.autoSyncInterval.Leave += new System.EventHandler(this.AutoSyncInterval_Leave);
            // 
            // autoSyncCheckBox
            // 
            this.autoSyncCheckBox.AutoSize = true;
            this.autoSyncCheckBox.Location = new System.Drawing.Point(12, 42);
            this.autoSyncCheckBox.Name = "autoSyncCheckBox";
            this.autoSyncCheckBox.Size = new System.Drawing.Size(75, 17);
            this.autoSyncCheckBox.TabIndex = 1;
            this.autoSyncCheckBox.Text = "&Auto Sync";
            this.autoSyncCheckBox.UseVisualStyleBackColor = true;
            this.autoSyncCheckBox.CheckedChanged += new System.EventHandler(this.AutoSyncCheckBox_CheckedChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 93);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(72, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Sync &Interval:";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(141, 93);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(33, 13);
            this.label4.TabIndex = 4;
            this.label4.Text = "hours";
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.btSyncOnlyIdle);
            this.groupBox1.Controls.Add(this.reportSyncResultCheckBox);
            this.groupBox1.Controls.Add(this.runAtStartupCheckBox);
            this.groupBox1.Controls.Add(this.nextSyncLabel);
            this.groupBox1.Controls.Add(this.autoSyncInterval);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.autoSyncCheckBox);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(5, 422);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(512, 138);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Automation";
            // 
            // btSyncOnlyIdle
            // 
            this.btSyncOnlyIdle.AutoSize = true;
            this.btSyncOnlyIdle.Checked = true;
            this.btSyncOnlyIdle.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncOnlyIdle.Location = new System.Drawing.Point(180, 94);
            this.btSyncOnlyIdle.Name = "btSyncOnlyIdle";
            this.btSyncOnlyIdle.Size = new System.Drawing.Size(120, 17);
            this.btSyncOnlyIdle.TabIndex = 29;
            this.btSyncOnlyIdle.Text = "Only sync when idle";
            this.toolTip.SetToolTip(this.btSyncOnlyIdle, "This specifies whether sync is only performed if PC is idle");
            this.btSyncOnlyIdle.UseVisualStyleBackColor = true;
            // 
            // reportSyncResultCheckBox
            // 
            this.reportSyncResultCheckBox.AutoSize = true;
            this.reportSyncResultCheckBox.Location = new System.Drawing.Point(12, 65);
            this.reportSyncResultCheckBox.Name = "reportSyncResultCheckBox";
            this.reportSyncResultCheckBox.Size = new System.Drawing.Size(190, 17);
            this.reportSyncResultCheckBox.TabIndex = 2;
            this.reportSyncResultCheckBox.Text = "Re&port Sync Result in System Tray";
            this.reportSyncResultCheckBox.UseVisualStyleBackColor = true;
            // 
            // runAtStartupCheckBox
            // 
            this.runAtStartupCheckBox.AutoSize = true;
            this.runAtStartupCheckBox.Location = new System.Drawing.Point(12, 21);
            this.runAtStartupCheckBox.Name = "runAtStartupCheckBox";
            this.runAtStartupCheckBox.Size = new System.Drawing.Size(213, 17);
            this.runAtStartupCheckBox.TabIndex = 0;
            this.runAtStartupCheckBox.Text = "&Launch application at Windows start up";
            this.runAtStartupCheckBox.UseVisualStyleBackColor = true;
            this.runAtStartupCheckBox.CheckedChanged += new System.EventHandler(this.RunAtStartupCheckBox_CheckedChanged);
            // 
            // nextSyncLabel
            // 
            this.nextSyncLabel.AutoSize = true;
            this.nextSyncLabel.Location = new System.Drawing.Point(9, 117);
            this.nextSyncLabel.Name = "nextSyncLabel";
            this.nextSyncLabel.Size = new System.Drawing.Size(67, 13);
            this.nextSyncLabel.TabIndex = 5;
            this.nextSyncLabel.Text = "Next Sync in";
            // 
            // syncTimer
            // 
            this.syncTimer.Interval = 1000;
            this.syncTimer.Tick += new System.EventHandler(this.SyncTimer_Tick);
            // 
            // syncOptionsGroupBox
            // 
            this.syncOptionsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.syncOptionsGroupBox.Controls.Add(this.IncludePastRemindersCheckBox);
            this.syncOptionsGroupBox.Controls.Add(this.SyncPhotosCheckBox);
            this.syncOptionsGroupBox.Controls.Add(this.btMonthsFuture);
            this.syncOptionsGroupBox.Controls.Add(this.btMonthsPast);
            this.syncOptionsGroupBox.Controls.Add(this.SyncRemindersCheckBox);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncAppointmentsPrivate);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncAppointmentsForceRTF);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncContactsForceRTF);
            this.syncOptionsGroupBox.Controls.Add(this.futureMonthInterval);
            this.syncOptionsGroupBox.Controls.Add(this.pastMonthInterval);
            this.syncOptionsGroupBox.Controls.Add(this.appointmentGoogleFoldersComboBox);
            this.syncOptionsGroupBox.Controls.Add(this.labelTimezone);
            this.syncOptionsGroupBox.Controls.Add(this.appointmentTimezonesComboBox);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncAppointments);
            this.syncOptionsGroupBox.Controls.Add(this.appointmentFoldersComboBox);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncContacts);
            this.syncOptionsGroupBox.Controls.Add(this.btPromptDelete);
            this.syncOptionsGroupBox.Controls.Add(this.panel1);
            this.syncOptionsGroupBox.Controls.Add(this.btSyncDelete);
            this.syncOptionsGroupBox.Controls.Add(this.cmbSyncProfile);
            this.syncOptionsGroupBox.Controls.Add(this.contactFoldersComboBox);
            this.syncOptionsGroupBox.Controls.Add(this.syncProfileLabel);
            this.syncOptionsGroupBox.Controls.Add(this.outlookStartModeLabel);
            this.syncOptionsGroupBox.Controls.Add(this.outlookStartModeComboBox);
            this.syncOptionsGroupBox.Controls.Add(this.profileBindingLabel);
            this.syncOptionsGroupBox.Controls.Add(this.syncOptionBox);
            this.syncOptionsGroupBox.Location = new System.Drawing.Point(5, 115);
            this.syncOptionsGroupBox.Name = "syncOptionsGroupBox";
            this.syncOptionsGroupBox.Size = new System.Drawing.Size(512, 307);
            this.syncOptionsGroupBox.TabIndex = 1;
            this.syncOptionsGroupBox.TabStop = false;
            this.syncOptionsGroupBox.Text = "Sync Options";
            // 
            // IncludePastRemindersCheckBox
            // 
            this.IncludePastRemindersCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.IncludePastRemindersCheckBox.AutoSize = true;
            this.IncludePastRemindersCheckBox.Location = new System.Drawing.Point(362, 278);
            this.IncludePastRemindersCheckBox.Name = "IncludePastRemindersCheckBox";
            this.IncludePastRemindersCheckBox.Size = new System.Drawing.Size(138, 17);
            this.IncludePastRemindersCheckBox.TabIndex = 34;
            this.IncludePastRemindersCheckBox.Text = "Include Past Reminders";
            this.toolTip.SetToolTip(this.IncludePastRemindersCheckBox, "This specifies whether past appointment reminders are also synchronized (switch o" +
        "ff to avoid ghost reminders");
            this.IncludePastRemindersCheckBox.UseVisualStyleBackColor = true;
            // 
            // SyncPhotosCheckBox
            // 
            this.SyncPhotosCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SyncPhotosCheckBox.AutoSize = true;
            this.SyncPhotosCheckBox.Checked = true;
            this.SyncPhotosCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SyncPhotosCheckBox.Location = new System.Drawing.Point(189, 212);
            this.SyncPhotosCheckBox.Name = "SyncPhotosCheckBox";
            this.SyncPhotosCheckBox.Size = new System.Drawing.Size(126, 17);
            this.SyncPhotosCheckBox.TabIndex = 33;
            this.SyncPhotosCheckBox.Text = "Sync Contact Photos";
            this.toolTip.SetToolTip(this.SyncPhotosCheckBox, "This specifies whether contact photos shall be synchronized between Outlook and G" +
        "oogle");
            this.SyncPhotosCheckBox.UseVisualStyleBackColor = true;
            // 
            // btMonthsFuture
            // 
            this.btMonthsFuture.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btMonthsFuture.AutoSize = true;
            this.btMonthsFuture.Checked = true;
            this.btMonthsFuture.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btMonthsFuture.Location = new System.Drawing.Point(343, 234);
            this.btMonthsFuture.Name = "btMonthsFuture";
            this.btMonthsFuture.Size = new System.Drawing.Size(100, 17);
            this.btMonthsFuture.TabIndex = 32;
            this.btMonthsFuture.Text = "Months (Future)";
            this.toolTip.SetToolTip(this.btMonthsFuture, resources.GetString("btMonthsFuture.ToolTip"));
            this.btMonthsFuture.UseVisualStyleBackColor = true;
            this.btMonthsFuture.CheckedChanged += new System.EventHandler(this.btMonthsFuture_CheckedChanged);
            // 
            // btMonthsPast
            // 
            this.btMonthsPast.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btMonthsPast.AutoSize = true;
            this.btMonthsPast.Checked = true;
            this.btMonthsPast.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btMonthsPast.Location = new System.Drawing.Point(343, 212);
            this.btMonthsPast.Name = "btMonthsPast";
            this.btMonthsPast.Size = new System.Drawing.Size(91, 17);
            this.btMonthsPast.TabIndex = 31;
            this.btMonthsPast.Text = "Months (Past)";
            this.toolTip.SetToolTip(this.btMonthsPast, "How many months into the past (uncheck if all)");
            this.btMonthsPast.UseVisualStyleBackColor = true;
            this.btMonthsPast.CheckedChanged += new System.EventHandler(this.btMonthsPast_CheckedChanged);
            // 
            // SyncRemindersCheckBox
            // 
            this.SyncRemindersCheckBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.SyncRemindersCheckBox.AutoSize = true;
            this.SyncRemindersCheckBox.Checked = true;
            this.SyncRemindersCheckBox.CheckState = System.Windows.Forms.CheckState.Checked;
            this.SyncRemindersCheckBox.Location = new System.Drawing.Point(343, 258);
            this.SyncRemindersCheckBox.Name = "SyncRemindersCheckBox";
            this.SyncRemindersCheckBox.Size = new System.Drawing.Size(103, 17);
            this.SyncRemindersCheckBox.TabIndex = 30;
            this.SyncRemindersCheckBox.Text = "Sync Reminders";
            this.toolTip.SetToolTip(this.SyncRemindersCheckBox, "This specifies whether appointment reminders are synchronized to target side (swi" +
        "tch off to avoid duplicate reminders from different Outlook folders)");
            this.SyncRemindersCheckBox.UseVisualStyleBackColor = true;
            this.SyncRemindersCheckBox.CheckedChanged += new System.EventHandler(this.SyncReminders_CheckedChanged);
            // 
            // btSyncAppointmentsPrivate
            // 
            this.btSyncAppointmentsPrivate.AutoSize = true;
            this.btSyncAppointmentsPrivate.Location = new System.Drawing.Point(137, 92);
            this.btSyncAppointmentsPrivate.Name = "btSyncAppointmentsPrivate";
            this.btSyncAppointmentsPrivate.Size = new System.Drawing.Size(189, 17);
            this.btSyncAppointmentsPrivate.TabIndex = 19;
            this.btSyncAppointmentsPrivate.Text = "Appointments private on other side";
            this.toolTip.SetToolTip(this.btSyncAppointmentsPrivate, resources.GetString("btSyncAppointmentsPrivate.ToolTip"));
            this.btSyncAppointmentsPrivate.UseVisualStyleBackColor = true;
            // 
            // btSyncAppointmentsForceRTF
            // 
            this.btSyncAppointmentsForceRTF.AutoSize = true;
            this.btSyncAppointmentsForceRTF.Checked = true;
            this.btSyncAppointmentsForceRTF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncAppointmentsForceRTF.Location = new System.Drawing.Point(9, 146);
            this.btSyncAppointmentsForceRTF.Name = "btSyncAppointmentsForceRTF";
            this.btSyncAppointmentsForceRTF.Size = new System.Drawing.Size(127, 17);
            this.btSyncAppointmentsForceRTF.TabIndex = 18;
            this.btSyncAppointmentsForceRTF.Text = "Force RTF Appointm.";
            this.toolTip.SetToolTip(this.btSyncAppointmentsForceRTF, "This specifies whether appointment notes are updated with RTF, even if it destroy" +
        "s the RTF format.");
            this.btSyncAppointmentsForceRTF.UseVisualStyleBackColor = true;
            // 
            // btSyncContactsForceRTF
            // 
            this.btSyncContactsForceRTF.AutoSize = true;
            this.btSyncContactsForceRTF.Checked = true;
            this.btSyncContactsForceRTF.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncContactsForceRTF.Location = new System.Drawing.Point(9, 92);
            this.btSyncContactsForceRTF.Name = "btSyncContactsForceRTF";
            this.btSyncContactsForceRTF.Size = new System.Drawing.Size(122, 17);
            this.btSyncContactsForceRTF.TabIndex = 17;
            this.btSyncContactsForceRTF.Text = "Force RTF Contacts";
            this.toolTip.SetToolTip(this.btSyncContactsForceRTF, "This specifies whether contact notes are updated with RTF, even if it destroys th" +
        "e RTF format.");
            this.btSyncContactsForceRTF.UseVisualStyleBackColor = true;
            // 
            // futureMonthInterval
            // 
            this.futureMonthInterval.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.futureMonthInterval.Location = new System.Drawing.Point(462, 232);
            this.futureMonthInterval.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.futureMonthInterval.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            -2147483648});
            this.futureMonthInterval.Name = "futureMonthInterval";
            this.futureMonthInterval.Size = new System.Drawing.Size(42, 20);
            this.futureMonthInterval.TabIndex = 11;
            this.futureMonthInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.futureMonthInterval.ValueChanged += new System.EventHandler(this.futureMonthInterval_ValueChanged);
            // 
            // pastMonthInterval
            // 
            this.pastMonthInterval.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.pastMonthInterval.Location = new System.Drawing.Point(462, 209);
            this.pastMonthInterval.Maximum = new decimal(new int[] {
            1000,
            0,
            0,
            0});
            this.pastMonthInterval.Name = "pastMonthInterval";
            this.pastMonthInterval.Size = new System.Drawing.Size(42, 20);
            this.pastMonthInterval.TabIndex = 10;
            this.pastMonthInterval.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.pastMonthInterval.Value = new decimal(new int[] {
            1,
            0,
            0,
            0});
            // 
            // appointmentGoogleFoldersComboBox
            // 
            this.appointmentGoogleFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.appointmentGoogleFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.appointmentGoogleFoldersComboBox.FormattingEnabled = true;
            this.appointmentGoogleFoldersComboBox.Location = new System.Drawing.Point(138, 146);
            this.appointmentGoogleFoldersComboBox.Name = "appointmentGoogleFoldersComboBox";
            this.appointmentGoogleFoldersComboBox.Size = new System.Drawing.Size(367, 21);
            this.appointmentGoogleFoldersComboBox.TabIndex = 7;
            this.toolTip.SetToolTip(this.appointmentGoogleFoldersComboBox, "Select the Google Appointments folder you want to sync");
            this.appointmentGoogleFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.AppointmentGoogleFoldersComboBox_SelectedIndexChanged);
            this.appointmentGoogleFoldersComboBox.Enter += new System.EventHandler(this.AppointmentGoogleFoldersComboBox_Enter);
            // 
            // labelTimezone
            // 
            this.labelTimezone.AutoSize = true;
            this.labelTimezone.Location = new System.Drawing.Point(193, 187);
            this.labelTimezone.Name = "labelTimezone";
            this.labelTimezone.Size = new System.Drawing.Size(56, 13);
            this.labelTimezone.TabIndex = 14;
            this.labelTimezone.Text = "Timezone:";
            // 
            // appointmentTimezonesComboBox
            // 
            this.appointmentTimezonesComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.appointmentTimezonesComboBox.FormattingEnabled = true;
            this.appointmentTimezonesComboBox.Items.AddRange(new object[] {
            "Africa/Abidjan",
            "Africa/Accra",
            "Africa/Addis_Ababa",
            "Africa/Asmera",
            "Africa/Bamako",
            "Africa/Bangui",
            "Africa/Banjul",
            "Africa/Blantyre",
            "Africa/Brazzaville",
            "Africa/Bujumbura",
            "Africa/Cairo",
            "Africa/Conakry",
            "Africa/Dakar",
            "Africa/Dar_es_Salaam",
            "Africa/Djibouti",
            "Africa/Douala",
            "Africa/El_Aaiun",
            "Africa/Freetown",
            "Africa/Gaborone",
            "Africa/Harare",
            "Africa/Johannesburg",
            "Africa/Kampala",
            "Africa/Kigali",
            "Africa/Kinshasa",
            "Africa/Lagos",
            "Africa/Libreville",
            "Africa/Lome",
            "Africa/Luanda",
            "Africa/Lubumbashi",
            "Africa/Lusaka",
            "Africa/Malabo",
            "Africa/Maputo",
            "Africa/Maseru",
            "Africa/Mbabane",
            "Africa/Mogadishu",
            "Africa/Monrovia",
            "Africa/Nairobi",
            "Africa/Ndjamena",
            "Africa/Niamey",
            "Africa/Nouakchott",
            "Africa/Ouagadougou",
            "Africa/Porto-Novo",
            "Africa/Sao_Tome",
            "Africa/Tunis",
            "America/Adak",
            "America/Anchorage",
            "America/Anguilla",
            "America/Antigua",
            "America/Argentina/San_Luis",
            "America/Aruba",
            "America/Asuncion",
            "America/Barbados",
            "America/Belize",
            "America/Bogota",
            "America/Buenos_Aires",
            "America/Caracas",
            "America/Cayenne",
            "America/Cayman",
            "America/Chicago",
            "America/Costa_Rica",
            "America/Curacao",
            "America/Denver",
            "America/Dominica",
            "America/Edmonton",
            "America/El_Salvador",
            "America/Godthab",
            "America/Goose_Bay",
            "America/Grand_Turk",
            "America/Grenada",
            "America/Guadeloupe",
            "America/Guatemala",
            "America/Guayaquil",
            "America/Guyana",
            "America/Halifax",
            "America/Havana",
            "America/Hermosillo",
            "America/Jamaica",
            "America/Juneau",
            "America/Kralendijk",
            "America/La_Paz",
            "America/Lima",
            "America/Los_Angeles",
            "America/Lower_Princes",
            "America/Manaus",
            "America/Marigot",
            "America/Martinique",
            "America/Mexico_City",
            "America/Miquelon",
            "America/Montevideo",
            "America/Montserrat",
            "America/Nassau",
            "America/New_York",
            "America/Noronha",
            "America/Panama",
            "America/Paramaribo",
            "America/Phoenix",
            "America/Port-au-Prince",
            "America/Port_of_Spain",
            "America/Puerto_Rico",
            "America/Rio_Branco",
            "America/Santiago",
            "America/Santo_Domingo",
            "America/Sao_Paulo",
            "America/Scoresbysund",
            "America/St_Johns",
            "America/St_Kitts",
            "America/St_Lucia",
            "America/St_Thomas",
            "America/St_Vincent",
            "America/Tegucigalpa",
            "America/Thule",
            "America/Tijuana",
            "America/Toronto",
            "America/Tortola",
            "America/Vancouver",
            "America/Winnipeg",
            "America/Yakutat",
            "Antarctica/Casey",
            "Antarctica/Davis",
            "Antarctica/DumontDUrville",
            "Antarctica/Macquarie",
            "Antarctica/Mawson",
            "Antarctica/McMurdo",
            "Antarctica/Palmer",
            "Antarctica/Rothera",
            "Antarctica/Syowa",
            "Antarctica/Vostok",
            "Asia/Aden",
            "Asia/Almaty",
            "Asia/Amman",
            "Asia/Anadyr",
            "Asia/Aqtau",
            "Asia/Aqtobe",
            "Asia/Ashgabat",
            "Asia/Baghdad",
            "Asia/Bahrain",
            "Asia/Baku",
            "Asia/Bangkok",
            "Asia/Beirut",
            "Asia/Bishkek",
            "Asia/Brunei",
            "Asia/Calcutta",
            "Asia/Choibalsan",
            "Asia/Chongqing",
            "Asia/Colombo",
            "Asia/Damascus",
            "Asia/Dhaka",
            "Asia/Dili",
            "Asia/Dubai",
            "Asia/Dushanbe",
            "Asia/Harbin",
            "Asia/Hong_Kong",
            "Asia/Hovd",
            "Asia/Irkutsk",
            "Asia/Jakarta",
            "Asia/Jayapura",
            "Asia/Jerusalem",
            "Asia/Kabul",
            "Asia/Kamchatka",
            "Asia/Karachi",
            "Asia/Kashgar",
            "Asia/Katmandu",
            "Asia/Krasnoyarsk",
            "Asia/Kuala_Lumpur",
            "Asia/Kuching",
            "Asia/Kuwait",
            "Asia/Macau",
            "Asia/Magadan",
            "Asia/Makassar",
            "Asia/Manila",
            "Asia/Muscat",
            "Asia/Nicosia",
            "Asia/Novosibirsk",
            "Asia/Omsk",
            "Asia/Oral",
            "Asia/Phnom_Penh",
            "Asia/Pyongyang",
            "Asia/Qatar",
            "Asia/Qyzylorda",
            "Asia/Rangoon",
            "Asia/Riyadh",
            "Asia/Saigon",
            "Asia/Sakhalin",
            "Asia/Samarkand",
            "Asia/Seoul",
            "Asia/Shanghai",
            "Asia/Singapore",
            "Asia/Taipei",
            "Asia/Tashkent",
            "Asia/Tbilisi",
            "Asia/Tehran",
            "Asia/Thimphu",
            "Asia/Tokyo",
            "Asia/Ulaanbaatar",
            "Asia/Urumqi",
            "Asia/Vientiane",
            "Asia/Vladivostok",
            "Asia/Yakutsk",
            "Asia/Yekaterinburg",
            "Asia/Yerevan",
            "Atlantic/Azores",
            "Atlantic/Bermuda",
            "Atlantic/Canary",
            "Atlantic/Cape_Verde",
            "Atlantic/Faeroe",
            "Atlantic/Reykjavik",
            "Atlantic/South_Georgia",
            "Atlantic/St_Helena",
            "Atlantic/Stanley",
            "Australia/Adelaide",
            "Australia/Eucla",
            "Australia/Lord_Howe",
            "Australia/Perth",
            "Australia/Sydney",
            "Europe/Amsterdam",
            "Europe/Andorra",
            "Europe/Athens",
            "Europe/Belgrade",
            "Europe/Berlin",
            "Europe/Bratislava",
            "Europe/Brussels",
            "Europe/Bucharest",
            "Europe/Budapest",
            "Europe/Copenhagen",
            "Europe/Dublin",
            "Europe/Gibraltar",
            "Europe/Helsinki",
            "Europe/Istanbul",
            "Europe/Ljubljana",
            "Europe/London",
            "Europe/Luxembourg",
            "Europe/Madrid",
            "Europe/Malta",
            "Europe/Mariehamn",
            "Europe/Monaco",
            "Europe/Moscow",
            "Europe/Oslo",
            "Europe/Paris",
            "Europe/Podgorica",
            "Europe/Prague",
            "Europe/Rome",
            "Europe/Samara",
            "Europe/San_Marino",
            "Europe/Sarajevo",
            "Europe/Skopje",
            "Europe/Sofia",
            "Europe/Stockholm",
            "Europe/Tirane",
            "Europe/Vaduz",
            "Europe/Vatican",
            "Europe/Vienna",
            "Europe/Volgograd",
            "Europe/Warsaw",
            "Europe/Zagreb",
            "Europe/Zurich",
            "Indian/Antananarivo",
            "Indian/Chagos",
            "Indian/Christmas",
            "Indian/Cocos",
            "Indian/Comoro",
            "Indian/Kerguelen",
            "Indian/Mahe",
            "Indian/Maldives",
            "Indian/Mauritius",
            "Indian/Mayotte",
            "Indian/Reunion",
            "Pacific/Apia",
            "Pacific/Auckland",
            "Pacific/Chatham",
            "Pacific/Easter",
            "Pacific/Efate",
            "Pacific/Enderbury",
            "Pacific/Fakaofo",
            "Pacific/Fiji",
            "Pacific/Funafuti",
            "Pacific/Galapagos",
            "Pacific/Gambier",
            "Pacific/Guadalcanal",
            "Pacific/Guam",
            "Pacific/Honolulu",
            "Pacific/Kiritimati",
            "Pacific/Kosrae",
            "Pacific/Kwajalein",
            "Pacific/Majuro",
            "Pacific/Marquesas",
            "Pacific/Nauru",
            "Pacific/Niue",
            "Pacific/Norfolk",
            "Pacific/Noumea",
            "Pacific/Palau",
            "Pacific/Pitcairn",
            "Pacific/Ponape",
            "Pacific/Port_Moresby",
            "Pacific/Rarotonga",
            "Pacific/Saipan",
            "Pacific/Tahiti",
            "Pacific/Tarawa",
            "Pacific/Tongatapu",
            "Pacific/Truk",
            "Pacific/Wake",
            "Pacific/Wallis"});
            this.appointmentTimezonesComboBox.Location = new System.Drawing.Point(255, 184);
            this.appointmentTimezonesComboBox.Name = "appointmentTimezonesComboBox";
            this.appointmentTimezonesComboBox.Size = new System.Drawing.Size(248, 21);
            this.appointmentTimezonesComboBox.Sorted = true;
            this.appointmentTimezonesComboBox.TabIndex = 9;
            this.toolTip.SetToolTip(this.appointmentTimezonesComboBox, "Select or enter Timezone (default is UTC), only for Recurrences!!!");
            this.appointmentTimezonesComboBox.TextChanged += new System.EventHandler(this.AppointmentTimezonesComboBox_TextChanged);
            // 
            // btSyncAppointments
            // 
            this.btSyncAppointments.AutoSize = true;
            this.btSyncAppointments.Location = new System.Drawing.Point(9, 119);
            this.btSyncAppointments.Name = "btSyncAppointments";
            this.btSyncAppointments.Size = new System.Drawing.Size(117, 17);
            this.btSyncAppointments.TabIndex = 5;
            this.btSyncAppointments.Text = "Sync &Appointments";
            this.toolTip.SetToolTip(this.btSyncAppointments, "This specifies whether appointments are synchronized.");
            this.btSyncAppointments.UseVisualStyleBackColor = true;
            this.btSyncAppointments.CheckedChanged += new System.EventHandler(this.BtSyncAppointments_CheckedChanged);
            // 
            // appointmentFoldersComboBox
            // 
            this.appointmentFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.appointmentFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.appointmentFoldersComboBox.FormattingEnabled = true;
            this.appointmentFoldersComboBox.Location = new System.Drawing.Point(138, 119);
            this.appointmentFoldersComboBox.Name = "appointmentFoldersComboBox";
            this.appointmentFoldersComboBox.Size = new System.Drawing.Size(367, 21);
            this.appointmentFoldersComboBox.TabIndex = 6;
            this.toolTip.SetToolTip(this.appointmentFoldersComboBox, "Select the Outlook Appointments folder you want to sync");
            this.appointmentFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.AppointmentFoldersComboBox_SelectedIndexChanged);
            // 
            // btSyncContacts
            // 
            this.btSyncContacts.AutoSize = true;
            this.btSyncContacts.Checked = true;
            this.btSyncContacts.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncContacts.Location = new System.Drawing.Point(9, 69);
            this.btSyncContacts.Name = "btSyncContacts";
            this.btSyncContacts.Size = new System.Drawing.Size(95, 17);
            this.btSyncContacts.TabIndex = 3;
            this.btSyncContacts.Text = "Sync &Contacts";
            this.toolTip.SetToolTip(this.btSyncContacts, "This specifies whether contacts are synchronized.");
            this.btSyncContacts.UseVisualStyleBackColor = true;
            this.btSyncContacts.CheckedChanged += new System.EventHandler(this.BtSyncContacts_CheckedChanged);
            // 
            // btPromptDelete
            // 
            this.btPromptDelete.AutoSize = true;
            this.btPromptDelete.Checked = true;
            this.btPromptDelete.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btPromptDelete.Location = new System.Drawing.Point(104, 44);
            this.btPromptDelete.Name = "btPromptDelete";
            this.btPromptDelete.Size = new System.Drawing.Size(101, 17);
            this.btPromptDelete.TabIndex = 2;
            this.btPromptDelete.Text = "Prompt De&letion";
            this.toolTip.SetToolTip(this.btPromptDelete, resources.GetString("btPromptDelete.ToolTip"));
            this.btPromptDelete.UseVisualStyleBackColor = true;
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(5, 179);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(497, 1);
            this.panel1.TabIndex = 8;
            // 
            // btSyncDelete
            // 
            this.btSyncDelete.AutoSize = true;
            this.btSyncDelete.Checked = true;
            this.btSyncDelete.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btSyncDelete.Location = new System.Drawing.Point(9, 44);
            this.btSyncDelete.Name = "btSyncDelete";
            this.btSyncDelete.Size = new System.Drawing.Size(92, 17);
            this.btSyncDelete.TabIndex = 1;
            this.btSyncDelete.Text = "Sync &Deletion";
            this.toolTip.SetToolTip(this.btSyncDelete, "This specifies whether deletions are\r\nsynchronized. Enabling this option\r\nmeans i" +
        "f you delete a contact from\r\nGoogle, then it will be deleted from\r\nOutlook and v" +
        "ice versa.");
            this.btSyncDelete.UseVisualStyleBackColor = true;
            this.btSyncDelete.CheckedChanged += new System.EventHandler(this.BtSyncDelete_CheckedChanged);
            // 
            // cmbSyncProfile
            // 
            this.cmbSyncProfile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cmbSyncProfile.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSyncProfile.FormattingEnabled = true;
            this.cmbSyncProfile.Location = new System.Drawing.Point(138, 19);
            this.cmbSyncProfile.Name = "cmbSyncProfile";
            this.cmbSyncProfile.Size = new System.Drawing.Size(367, 21);
            this.cmbSyncProfile.TabIndex = 0;
            this.toolTip.SetToolTip(this.cmbSyncProfile, "This is a profile name of your choice.\r\nIt must be unique in each computer\r\nand a" +
        "ccount you intend to sync with\r\nyour Google Mail account.");
            this.cmbSyncProfile.SelectedIndexChanged += new System.EventHandler(this.CmbSyncProfile_SelectedIndexChanged);
            // 
            // contactFoldersComboBox
            // 
            this.contactFoldersComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.contactFoldersComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.contactFoldersComboBox.FormattingEnabled = true;
            this.contactFoldersComboBox.Location = new System.Drawing.Point(138, 65);
            this.contactFoldersComboBox.Name = "contactFoldersComboBox";
            this.contactFoldersComboBox.Size = new System.Drawing.Size(367, 21);
            this.contactFoldersComboBox.TabIndex = 4;
            this.toolTip.SetToolTip(this.contactFoldersComboBox, "Select the Outlook Contacts folder you want to sync");
            this.contactFoldersComboBox.SelectedIndexChanged += new System.EventHandler(this.ContacFoldersComboBox_SelectedIndexChanged);
            // 
            // syncProfileLabel
            // 
            this.syncProfileLabel.AutoSize = true;
            this.syncProfileLabel.Location = new System.Drawing.Point(6, 22);
            this.syncProfileLabel.Name = "syncProfileLabel";
            this.syncProfileLabel.Size = new System.Drawing.Size(66, 13);
            this.syncProfileLabel.TabIndex = 0;
            this.syncProfileLabel.Text = "&Sync Profile:";
            // 
            // outlookStartModeLabel
            // 
            this.outlookStartModeLabel.AutoSize = true;
            this.outlookStartModeLabel.Location = new System.Drawing.Point(208, 47);
            this.outlookStartModeLabel.Name = "outlookStartModeLabel";
            this.outlookStartModeLabel.Size = new System.Drawing.Size(67, 13);
            this.outlookStartModeLabel.TabIndex = 35;
            this.outlookStartModeLabel.Text = "Outlook start:";
            // 
            // outlookStartModeComboBox
            // 
            this.outlookStartModeComboBox.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.outlookStartModeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.outlookStartModeComboBox.FormattingEnabled = true;
            this.outlookStartModeComboBox.Items.AddRange(new object[] {
            "Prompt",
            "Auto-start Outlook",
            "Skip check"});
            this.outlookStartModeComboBox.Location = new System.Drawing.Point(291, 44);
            this.outlookStartModeComboBox.Name = "outlookStartModeComboBox";
            this.outlookStartModeComboBox.Size = new System.Drawing.Size(214, 21);
            this.outlookStartModeComboBox.TabIndex = 36;
            this.toolTip.SetToolTip(this.outlookStartModeComboBox, "Prompt: ask before sync if Outlook is closed.\r\nAuto-start: launch Outlook before" +
        " sync.\r\nSkip check: continue without checking Outlook process.");
            this.outlookStartModeComboBox.SelectedIndexChanged += new System.EventHandler(this.OutlookStartModeComboBox_SelectedIndexChanged);
            // 
            // profileBindingLabel
            // 
            this.profileBindingLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.profileBindingLabel.AutoEllipsis = true;
            this.profileBindingLabel.Location = new System.Drawing.Point(6, 168);
            this.profileBindingLabel.Name = "profileBindingLabel";
            this.profileBindingLabel.Size = new System.Drawing.Size(499, 13);
            this.profileBindingLabel.TabIndex = 37;
            this.profileBindingLabel.Text = "Profile binding: (not selected)";
            this.toolTip.SetToolTip(this.profileBindingLabel, "Shows which profile is bound to the selected Outlook and Google account.");
            // 
            // lastSyncLabel
            // 
            this.lastSyncLabel.AutoEllipsis = true;
            this.lastSyncLabel.AutoSize = true;
            this.lastSyncLabel.Dock = System.Windows.Forms.DockStyle.Top;
            this.lastSyncLabel.Location = new System.Drawing.Point(3, 16);
            this.lastSyncLabel.Name = "lastSyncLabel";
            this.lastSyncLabel.Size = new System.Drawing.Size(69, 13);
            this.lastSyncLabel.TabIndex = 0;
            this.lastSyncLabel.Text = "Last Sync on";
            // 
            // logGroupBox
            // 
            this.logGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.logGroupBox.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.logGroupBox.Controls.Add(this.syncConsole);
            this.logGroupBox.Controls.Add(this.lastSyncLabel);
            this.logGroupBox.Location = new System.Drawing.Point(526, 0);
            this.logGroupBox.Name = "logGroupBox";
            this.logGroupBox.Size = new System.Drawing.Size(709, 567);
            this.logGroupBox.TabIndex = 2;
            this.logGroupBox.TabStop = false;
            this.logGroupBox.Text = "Sync Details && Log";
            // 
            // syncConsole
            // 
            this.syncConsole.BackColor = System.Drawing.SystemColors.Info;
            this.syncConsole.Dock = System.Windows.Forms.DockStyle.Fill;
            this.syncConsole.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.syncConsole.Location = new System.Drawing.Point(3, 29);
            this.syncConsole.Name = "syncConsole";
            this.syncConsole.ReadOnly = true;
            this.syncConsole.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical;
            this.syncConsole.Size = new System.Drawing.Size(703, 535);
            this.syncConsole.TabIndex = 1;
            this.toolTip.SetToolTip(this.syncConsole, "This window shows information\r\n from the last sync.");
            // 
            // googleAccountGroupBox
            // 
            this.googleAccountGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.googleAccountGroupBox.Controls.Add(this.linkLabelRevokeAuthentication);
            this.googleAccountGroupBox.Controls.Add(this.chkUseFileAs);
            this.googleAccountGroupBox.Controls.Add(this.proxySettingsLinkLabel);
            this.googleAccountGroupBox.Controls.Add(this.googleEmailLabel);
            this.googleAccountGroupBox.Controls.Add(this.UserName);
            this.googleAccountGroupBox.Location = new System.Drawing.Point(5, 20);
            this.googleAccountGroupBox.Name = "googleAccountGroupBox";
            this.googleAccountGroupBox.Size = new System.Drawing.Size(512, 92);
            this.googleAccountGroupBox.TabIndex = 0;
            this.googleAccountGroupBox.TabStop = false;
            this.googleAccountGroupBox.Text = "Google Account";
            // 
            // linkLabelRevokeAuthentication
            // 
            this.linkLabelRevokeAuthentication.AutoSize = true;
            this.linkLabelRevokeAuthentication.Location = new System.Drawing.Point(6, 42);
            this.linkLabelRevokeAuthentication.Name = "linkLabelRevokeAuthentication";
            this.linkLabelRevokeAuthentication.Size = new System.Drawing.Size(116, 13);
            this.linkLabelRevokeAuthentication.TabIndex = 1;
            this.linkLabelRevokeAuthentication.TabStop = true;
            this.linkLabelRevokeAuthentication.Text = "Revoke Authentication";
            this.toolTip.SetToolTip(this.linkLabelRevokeAuthentication, "If you want to authenticate GoogleContactsSyncMod again, or if something didn\'t w" +
        "ork corecctly with authenication.");
            this.linkLabelRevokeAuthentication.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.LinkLabelRevokeAuthentication_LinkClicked);
            // 
            // chkUseFileAs
            // 
            this.chkUseFileAs.AutoSize = true;
            this.chkUseFileAs.Checked = true;
            this.chkUseFileAs.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkUseFileAs.Location = new System.Drawing.Point(138, 69);
            this.chkUseFileAs.Name = "chkUseFileAs";
            this.chkUseFileAs.Size = new System.Drawing.Size(238, 17);
            this.chkUseFileAs.TabIndex = 7;
            this.chkUseFileAs.Text = "Use Outlook Contact\'s FileAs for Google Title";
            this.chkUseFileAs.UseVisualStyleBackColor = true;
            // 
            // proxySettingsLinkLabel
            // 
            this.proxySettingsLinkLabel.AutoSize = true;
            this.proxySettingsLinkLabel.Location = new System.Drawing.Point(6, 69);
            this.proxySettingsLinkLabel.Name = "proxySettingsLinkLabel";
            this.proxySettingsLinkLabel.Size = new System.Drawing.Size(74, 13);
            this.proxySettingsLinkLabel.TabIndex = 2;
            this.proxySettingsLinkLabel.TabStop = true;
            this.proxySettingsLinkLabel.Text = "Proxy Settings";
            this.toolTip.SetToolTip(this.proxySettingsLinkLabel, resources.GetString("proxySettingsLinkLabel.ToolTip"));
            this.proxySettingsLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ProxySettingsLinkLabel_LinkClicked);
            // 
            // resetMatchesLinkLabel
            // 
            this.resetMatchesLinkLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.resetMatchesLinkLabel.AutoSize = true;
            this.resetMatchesLinkLabel.Location = new System.Drawing.Point(45, 585);
            this.resetMatchesLinkLabel.Name = "resetMatchesLinkLabel";
            this.resetMatchesLinkLabel.Size = new System.Drawing.Size(79, 13);
            this.resetMatchesLinkLabel.TabIndex = 0;
            this.resetMatchesLinkLabel.TabStop = true;
            this.resetMatchesLinkLabel.Text = "&Reset Matches";
            this.toolTip.SetToolTip(this.resetMatchesLinkLabel, "This unlinks Outlook contacts with their\r\ncorresponding Google contatcs. If you\r\n" +
        "accidentaly delete a contact and you\r\ndont want the deletion to be synchronised," +
        "\r\nclick  this link.");
            this.resetMatchesLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.ResetMatchesLinkLabel_LinkClicked);
            // 
            // Donate
            // 
            this.Donate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.Donate.BackColor = System.Drawing.Color.Transparent;
            this.Donate.Image = ((System.Drawing.Image)(resources.GetObject("Donate.Image")));
            this.Donate.Location = new System.Drawing.Point(10, 585);
            this.Donate.Name = "Donate";
            this.Donate.Size = new System.Drawing.Size(29, 34);
            this.Donate.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.Donate.TabIndex = 4;
            this.Donate.TabStop = false;
            this.toolTip.SetToolTip(this.Donate, resources.GetString("Donate.ToolTip"));
            this.Donate.Click += new System.EventHandler(this.Donate_Click);
            this.Donate.MouseEnter += new System.EventHandler(this.Donate_MouseEnter);
            this.Donate.MouseLeave += new System.EventHandler(this.Donate_MouseLeave);
            // 
            // pictureBoxExit
            // 
            this.pictureBoxExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.pictureBoxExit.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("pictureBoxExit.BackgroundImage")));
            this.pictureBoxExit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.pictureBoxExit.Image = ((System.Drawing.Image)(resources.GetObject("pictureBoxExit.Image")));
            this.pictureBoxExit.Location = new System.Drawing.Point(1226, 588);
            this.pictureBoxExit.Margin = new System.Windows.Forms.Padding(0);
            this.pictureBoxExit.Name = "pictureBoxExit";
            this.pictureBoxExit.Size = new System.Drawing.Size(21, 25);
            this.pictureBoxExit.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pictureBoxExit.TabIndex = 5;
            this.pictureBoxExit.TabStop = false;
            this.toolTip.SetToolTip(this.pictureBoxExit, "Exit GO Contact Sync Mod");
            this.pictureBoxExit.Click += new System.EventHandler(this.PictureBoxExit_Click);
            // 
            // settingsGroupBox
            // 
            this.settingsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.settingsGroupBox.Controls.Add(this.groupBox1);
            this.settingsGroupBox.Controls.Add(this.googleAccountGroupBox);
            this.settingsGroupBox.Controls.Add(this.syncOptionsGroupBox);
            this.settingsGroupBox.Location = new System.Drawing.Point(0, 0);
            this.settingsGroupBox.MinimumSize = new System.Drawing.Size(429, 0);
            this.settingsGroupBox.Name = "settingsGroupBox";
            this.settingsGroupBox.Size = new System.Drawing.Size(522, 567);
            this.settingsGroupBox.TabIndex = 0;
            this.settingsGroupBox.TabStop = false;
            this.settingsGroupBox.Text = "Program Settings";
            // 
            // actionsTableLayout
            // 
            this.actionsTableLayout.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.actionsTableLayout.ColumnCount = 3;
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 63F));
            this.actionsTableLayout.Controls.Add(this.cancelButton, 0, 0);
            this.actionsTableLayout.Controls.Add(this.syncButton, 0, 0);
            this.actionsTableLayout.Controls.Add(this.hideButton, 2, 0);
            this.actionsTableLayout.Location = new System.Drawing.Point(1016, 585);
            this.actionsTableLayout.Name = "actionsTableLayout";
            this.actionsTableLayout.RowCount = 1;
            this.actionsTableLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.actionsTableLayout.Size = new System.Drawing.Size(207, 31);
            this.actionsTableLayout.TabIndex = 1;
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.Enabled = false;
            this.cancelButton.Location = new System.Drawing.Point(75, 3);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(66, 25);
            this.cancelButton.TabIndex = 1;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // hideButton
            // 
            this.hideButton.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.hideButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.hideButton.Location = new System.Drawing.Point(147, 3);
            this.hideButton.Name = "hideButton";
            this.hideButton.Size = new System.Drawing.Size(57, 25);
            this.hideButton.TabIndex = 2;
            this.hideButton.Text = "&Hide";
            this.hideButton.UseVisualStyleBackColor = true;
            this.hideButton.Click += new System.EventHandler(this.HideButton_Click);
            // 
            // MainPanel
            // 
            this.MainPanel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.MainPanel.Controls.Add(this.logGroupBox);
            this.MainPanel.Controls.Add(this.MainSplitter);
            this.MainPanel.Controls.Add(this.settingsGroupBox);
            this.MainPanel.Location = new System.Drawing.Point(10, 12);
            this.MainPanel.Name = "MainPanel";
            this.MainPanel.Size = new System.Drawing.Size(1238, 567);
            this.MainPanel.TabIndex = 0;
            // 
            // MainSplitter
            // 
            this.MainSplitter.Location = new System.Drawing.Point(0, 0);
            this.MainSplitter.Name = "MainSplitter";
            this.MainSplitter.Size = new System.Drawing.Size(4, 567);
            this.MainSplitter.TabIndex = 5;
            this.MainSplitter.TabStop = false;
            // 
            // iconTimer
            // 
            this.iconTimer.Interval = 150;
            this.iconTimer.Tick += new System.EventHandler(this.IconTimer_Tick);
            // 
            // NewVersionLinkLabel
            // 
            this.NewVersionLinkLabel.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.NewVersionLinkLabel.AutoSize = true;
            this.NewVersionLinkLabel.LinkColor = System.Drawing.Color.Blue;
            this.NewVersionLinkLabel.Location = new System.Drawing.Point(231, 603);
            this.NewVersionLinkLabel.Name = "NewVersionLinkLabel";
            this.NewVersionLinkLabel.Size = new System.Drawing.Size(426, 13);
            this.NewVersionLinkLabel.TabIndex = 1;
            this.NewVersionLinkLabel.TabStop = true;
            this.NewVersionLinkLabel.Text = "You are using the latest version of GCSM, click here to open GCSM homepage on sf." +
    "net!";
            this.NewVersionLinkLabel.Visible = false;
            this.NewVersionLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.NewVersionLinkLabel_LinkClicked);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.syncButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1258, 628);
            this.Controls.Add(this.NewVersionLinkLabel);
            this.Controls.Add(this.Donate);
            this.Controls.Add(this.MainPanel);
            this.Controls.Add(this.pictureBoxExit);
            this.Controls.Add(this.resetMatchesLinkLabel);
            this.Controls.Add(this.actionsTableLayout);
            this.HelpButton = true;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(620, 519);
            this.Name = "SettingsForm";
            this.Text = "GO Contact Sync Mod - Settings";
            this.WindowState = System.Windows.Forms.FormWindowState.Minimized;
            this.HelpButtonClicked += new System.ComponentModel.CancelEventHandler(this.SettingsForm_HelpButtonClicked);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SettingsForm_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsForm_FormClosed);
            this.Load += new System.EventHandler(this.SettingsForm_Load);
            this.HelpRequested += new System.Windows.Forms.HelpEventHandler(this.SettingsForm_HelpRequested);
            this.Resize += new System.EventHandler(this.SettingsForm_Resize);
            this.systemTrayMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.autoSyncInterval)).EndInit();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.syncOptionsGroupBox.ResumeLayout(false);
            this.syncOptionsGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.futureMonthInterval)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pastMonthInterval)).EndInit();
            this.logGroupBox.ResumeLayout(false);
            this.logGroupBox.PerformLayout();
            this.googleAccountGroupBox.ResumeLayout(false);
            this.googleAccountGroupBox.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.Donate)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBoxExit)).EndInit();
            this.settingsGroupBox.ResumeLayout(false);
            this.actionsTableLayout.ResumeLayout(false);
            this.MainPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.Label googleEmailLabel;
        private System.Windows.Forms.Button syncButton;
        private System.Windows.Forms.CheckedListBox syncOptionBox;
        internal System.Windows.Forms.NotifyIcon notifyIcon;
        private System.Windows.Forms.NumericUpDown autoSyncInterval;
        private System.Windows.Forms.CheckBox autoSyncCheckBox;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Timer syncTimer;
        private System.Windows.Forms.Label nextSyncLabel;
        private System.Windows.Forms.GroupBox syncOptionsGroupBox;
        private System.Windows.Forms.Label lastSyncLabel;
        private System.Windows.Forms.GroupBox logGroupBox;
        private System.Windows.Forms.RichTextBox syncConsole;
        private System.Windows.Forms.GroupBox googleAccountGroupBox;
        private System.Windows.Forms.ContextMenuStrip systemTrayMenu;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem3;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem4;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.CheckBox runAtStartupCheckBox;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem5;
        private System.Windows.Forms.Label syncProfileLabel;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.CheckBox btSyncDelete;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.GroupBox settingsGroupBox;
        private System.Windows.Forms.TableLayoutPanel actionsTableLayout;
        private System.Windows.Forms.Panel MainPanel;
        private System.Windows.Forms.Splitter MainSplitter;
        private System.Windows.Forms.LinkLabel resetMatchesLinkLabel;
        internal System.Windows.Forms.PictureBox Donate;
        private System.Windows.Forms.Button hideButton;
        private System.Windows.Forms.LinkLabel proxySettingsLinkLabel;
        private System.Windows.Forms.CheckBox reportSyncResultCheckBox;
        private System.Windows.Forms.CheckBox btSyncContacts;
        private System.Windows.Forms.ComboBox contactFoldersComboBox;
        private System.Windows.Forms.ComboBox cmbSyncProfile;
        private System.Windows.Forms.CheckBox btPromptDelete;
        private System.Windows.Forms.CheckBox chkUseFileAs;
        private System.Windows.Forms.PictureBox pictureBoxExit;
        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Timer iconTimer;
        private System.Windows.Forms.ComboBox appointmentFoldersComboBox;
        private System.Windows.Forms.CheckBox btSyncAppointments;
        private System.Windows.Forms.ComboBox appointmentTimezonesComboBox;
        private System.Windows.Forms.Label labelTimezone;
        private System.Windows.Forms.LinkLabel linkLabelRevokeAuthentication;
        private System.Windows.Forms.ComboBox appointmentGoogleFoldersComboBox;
        private System.Windows.Forms.NumericUpDown futureMonthInterval;
        private System.Windows.Forms.NumericUpDown pastMonthInterval;
        private System.Windows.Forms.LinkLabel NewVersionLinkLabel;
        private System.Windows.Forms.CheckBox btSyncAppointmentsForceRTF;
        private System.Windows.Forms.CheckBox btSyncContactsForceRTF;
        private System.Windows.Forms.CheckBox btSyncAppointmentsPrivate;
        private System.Windows.Forms.CheckBox btSyncOnlyIdle;
        private System.Windows.Forms.CheckBox SyncRemindersCheckBox;
        private System.Windows.Forms.CheckBox btMonthsFuture;
        private System.Windows.Forms.CheckBox btMonthsPast;
        private System.Windows.Forms.CheckBox SyncPhotosCheckBox;
        private System.Windows.Forms.CheckBox IncludePastRemindersCheckBox;
        private System.Windows.Forms.Label outlookStartModeLabel;
        private System.Windows.Forms.ComboBox outlookStartModeComboBox;
        private System.Windows.Forms.Label profileBindingLabel;
    }
}

