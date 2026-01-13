namespace GoContactSyncMod
{
    partial class ProxySettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ProxySettingsForm));
            this.cancelButton = new System.Windows.Forms.Button();
            this.okButton = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.UserName = new System.Windows.Forms.TextBox();
            this.Password = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbUseGlobalSettings = new System.Windows.Forms.CheckBox();
            this.Authorization = new System.Windows.Forms.CheckBox();
            this.CustomProxy = new System.Windows.Forms.RadioButton();
            this.SystemProxy = new System.Windows.Forms.RadioButton();
            this.label1 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Address = new System.Windows.Forms.TextBox();
            this.Port = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // cancelButton
            // 
            this.cancelButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.cancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.cancelButton.Location = new System.Drawing.Point(189, 226);
            this.cancelButton.Name = "cancelButton";
            this.cancelButton.Size = new System.Drawing.Size(98, 23);
            this.cancelButton.TabIndex = 1;
            this.cancelButton.Text = "&Cancel";
            this.cancelButton.UseVisualStyleBackColor = true;
            this.cancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // okButton
            // 
            this.okButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.okButton.Location = new System.Drawing.Point(85, 226);
            this.okButton.Name = "okButton";
            this.okButton.Size = new System.Drawing.Size(98, 23);
            this.okButton.TabIndex = 0;
            this.okButton.Text = "&OK";
            this.okButton.UseVisualStyleBackColor = true;
            this.okButton.Click += new System.EventHandler(this.OkButton_Click);
            // 
            // label2
            // 
            this.label2.Location = new System.Drawing.Point(7, 158);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(75, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "&User:";
            // 
            // label3
            // 
            this.label3.Location = new System.Drawing.Point(7, 181);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(75, 16);
            this.label3.TabIndex = 10;
            this.label3.Text = "&Password:";
            // 
            // UserName
            // 
            this.UserName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.UserName.Location = new System.Drawing.Point(100, 155);
            this.UserName.Name = "UserName";
            this.UserName.Size = new System.Drawing.Size(167, 20);
            this.UserName.TabIndex = 5;
            this.UserName.TextChanged += new System.EventHandler(this.Form_Changed);
            // 
            // Password
            // 
            this.Password.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Password.Location = new System.Drawing.Point(100, 181);
            this.Password.Name = "Password";
            this.Password.PasswordChar = '*';
            this.Password.Size = new System.Drawing.Size(167, 20);
            this.Password.TabIndex = 6;
            this.Password.TextChanged += new System.EventHandler(this.Form_Changed);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.cbUseGlobalSettings);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.Authorization);
            this.groupBox1.Controls.Add(this.UserName);
            this.groupBox1.Controls.Add(this.CustomProxy);
            this.groupBox1.Controls.Add(this.Password);
            this.groupBox1.Controls.Add(this.SystemProxy);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.Address);
            this.groupBox1.Controls.Add(this.Port);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(275, 209);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Proxy";
            // 
            // cbUseGlobalSettings
            // 
            this.cbUseGlobalSettings.Location = new System.Drawing.Point(0, 0);
            this.cbUseGlobalSettings.Name = "cbUseGlobalSettings";
            this.cbUseGlobalSettings.Size = new System.Drawing.Size(104, 24);
            this.cbUseGlobalSettings.TabIndex = 0;
            // 
            // Authorization
            // 
            this.Authorization.AutoSize = true;
            this.Authorization.Location = new System.Drawing.Point(10, 138);
            this.Authorization.Name = "Authorization";
            this.Authorization.Size = new System.Drawing.Size(87, 17);
            this.Authorization.TabIndex = 4;
            this.Authorization.Text = "A&uthorization";
            this.Authorization.UseVisualStyleBackColor = true;
            this.Authorization.CheckedChanged += new System.EventHandler(this.Form_Changed);
            // 
            // CustomProxy
            // 
            this.CustomProxy.AutoSize = true;
            this.CustomProxy.Location = new System.Drawing.Point(10, 58);
            this.CustomProxy.Name = "CustomProxy";
            this.CustomProxy.Size = new System.Drawing.Size(147, 17);
            this.CustomProxy.TabIndex = 1;
            this.CustomProxy.Text = "Use custom HTTP  proxy ";
            this.CustomProxy.UseVisualStyleBackColor = true;
            this.CustomProxy.CheckedChanged += new System.EventHandler(this.Form_Changed);
            // 
            // SystemProxy
            // 
            this.SystemProxy.AutoSize = true;
            this.SystemProxy.Checked = true;
            this.SystemProxy.Location = new System.Drawing.Point(10, 35);
            this.SystemProxy.Name = "SystemProxy";
            this.SystemProxy.Size = new System.Drawing.Size(88, 17);
            this.SystemProxy.TabIndex = 0;
            this.SystemProxy.TabStop = true;
            this.SystemProxy.Text = "System Proxy";
            this.SystemProxy.UseVisualStyleBackColor = true;
            this.SystemProxy.CheckedChanged += new System.EventHandler(this.Form_Changed);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(7, 89);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(48, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "&Address:";
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(7, 112);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(48, 16);
            this.label4.TabIndex = 8;
            this.label4.Text = "P&ort:";
            // 
            // Address
            // 
            this.Address.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Address.Location = new System.Drawing.Point(61, 86);
            this.Address.Name = "Address";
            this.Address.Size = new System.Drawing.Size(206, 20);
            this.Address.TabIndex = 2;
            this.Address.TextChanged += new System.EventHandler(this.Form_Changed);
            // 
            // Port
            // 
            this.Port.Location = new System.Drawing.Point(61, 112);
            this.Port.Name = "Port";
            this.Port.Size = new System.Drawing.Size(37, 20);
            this.Port.TabIndex = 3;
            this.Port.TextChanged += new System.EventHandler(this.Form_Changed);
            // 
            // ProxySettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.cancelButton;
            this.ClientSize = new System.Drawing.Size(299, 261);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.cancelButton);
            this.Controls.Add(this.okButton);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(315, 299);
            this.Name = "ProxySettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Proxy Settings";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button cancelButton;
        private System.Windows.Forms.Button okButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox UserName;
        private System.Windows.Forms.TextBox Password;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton SystemProxy;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Address;
        private System.Windows.Forms.TextBox Port;
        private System.Windows.Forms.RadioButton CustomProxy;
        private System.Windows.Forms.CheckBox Authorization;
        private System.Windows.Forms.CheckBox cbUseGlobalSettings;

    }
}