using System.ComponentModel;
using System.Windows.Forms;

namespace GoContactSyncMod
{
    partial class DeleteDuplicatedPropertiesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(DeleteDuplicatedPropertiesForm));
            this.bbOK = new System.Windows.Forms.Button();
            this.bbCancel = new System.Windows.Forms.Button();
            this.propertiesGrid = new System.Windows.Forms.DataGridView();
            this.Selected = new System.Windows.Forms.DataGridViewCheckBoxColumn();
            this.Key = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Value = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.explanationLabel = new System.Windows.Forms.Label();
            this.allCheck = new System.Windows.Forms.CheckBox();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            ((System.ComponentModel.ISupportInitialize)(this.propertiesGrid)).BeginInit();
            this.tableLayoutPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // bbOK
            // 
            this.bbOK.Anchor = System.Windows.Forms.AnchorStyles.Right;
            this.bbOK.AutoSize = true;
            this.bbOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.bbOK.Location = new System.Drawing.Point(100, 356);
            this.bbOK.Name = "bbOK";
            this.bbOK.Size = new System.Drawing.Size(75, 23);
            this.bbOK.TabIndex = 5;
            this.bbOK.Text = "OK";
            this.bbOK.UseVisualStyleBackColor = false;
            // 
            // bbCancel
            // 
            this.bbCancel.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.bbCancel.AutoEllipsis = true;
            this.bbCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.bbCancel.Location = new System.Drawing.Point(400, 395);
            this.bbCancel.Name = "bbCancel";
            this.bbCancel.Size = new System.Drawing.Size(75, 23);
            this.bbCancel.TabIndex = 6;
            this.bbCancel.Text = "Cancel";
            this.bbCancel.UseVisualStyleBackColor = true;
            // 
            // propertiesGrid
            // 
            this.propertiesGrid.AllowUserToAddRows = false;
            this.propertiesGrid.AllowUserToDeleteRows = false;
            this.propertiesGrid.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.propertiesGrid.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells;
            this.propertiesGrid.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.propertiesGrid.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Selected,
            this.Key,
            this.Value});
            this.tableLayoutPanel1.SetColumnSpan(this.propertiesGrid, 2);
            this.propertiesGrid.Dock = System.Windows.Forms.DockStyle.Fill;
            this.propertiesGrid.Location = new System.Drawing.Point(3, 83);
            this.propertiesGrid.Name = "propertiesGrid";
            this.propertiesGrid.RowHeadersVisible = false;
            this.propertiesGrid.RowHeadersWidth = 62;
            this.propertiesGrid.Size = new System.Drawing.Size(595, 227);
            this.propertiesGrid.TabIndex = 3;
            // 
            // Selected
            // 
            this.Selected.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Selected.HeaderText = "";
            this.Selected.MinimumWidth = 8;
            this.Selected.Name = "Selected";
            this.Selected.Width = 8;
            // 
            // Key
            // 
            this.Key.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Key.HeaderText = "Key";
            this.Key.MinimumWidth = 8;
            this.Key.Name = "Key";
            this.Key.ReadOnly = true;
            this.Key.Width = 54;
            // 
            // Value
            // 
            this.Value.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.Value.HeaderText = "Value";
            this.Value.MinimumWidth = 8;
            this.Value.Name = "Value";
            this.Value.ReadOnly = true;
            this.Value.Width = 63;
            // 
            // explanationLabel
            // 
            this.explanationLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.explanationLabel.AutoEllipsis = true;
            this.tableLayoutPanel1.SetColumnSpan(this.explanationLabel, 2);
            this.explanationLabel.Location = new System.Drawing.Point(3, 0);
            this.explanationLabel.Name = "explanationLabel";
            this.explanationLabel.Size = new System.Drawing.Size(606, 80);
            this.explanationLabel.TabIndex = 2;
            this.explanationLabel.Text = resources.GetString("explanationLabel.Text");
            // 
            // allCheck
            // 
            this.allCheck.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.allCheck.AutoSize = true;
            this.tableLayoutPanel1.SetColumnSpan(this.allCheck, 2);
            this.allCheck.Location = new System.Drawing.Point(3, 316);
            this.allCheck.Name = "allCheck";
            this.allCheck.Size = new System.Drawing.Size(595, 34);
            this.allCheck.TabIndex = 4;
            this.allCheck.Text = "Remove the same extended properties from next contacts.";
            this.allCheck.UseVisualStyleBackColor = true;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.explanationLabel, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.bbCancel, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.allCheck, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.bbOK, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.propertiesGrid, 0, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(14, 12);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.Padding = new System.Windows.Forms.Padding(10);
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 80F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 40F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(601, 413);
            this.tableLayoutPanel1.TabIndex = 1;
            // 
            // DeleteDuplicatedPropertiesForm
            // 
            this.AcceptButton = this.bbOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.bbCancel;
            this.ClientSize = new System.Drawing.Size(627, 437);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Font = new System.Drawing.Font("Verdana", 8.25F);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "DeleteDuplicatedPropertiesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Remove extended properties";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.DeleteDuplicatedPropertiesForm_FormClosing);
            this.Load += new System.EventHandler(this.DeleteDuplicatedPropertiesForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.propertiesGrid)).EndInit();
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.Button bbOK;
        private System.Windows.Forms.Button bbCancel;
        private System.Windows.Forms.DataGridViewCheckBoxColumn Selected;
        private System.Windows.Forms.DataGridViewTextBoxColumn Key;
        private System.Windows.Forms.DataGridViewTextBoxColumn Value;
        private System.Windows.Forms.CheckBox allCheck;
        private System.Windows.Forms.DataGridView propertiesGrid;

        public bool removeFromAll
        {
            get { return allCheck.Checked; }
        }

        public void AddExtendedProperty(bool selected, string name, string value)
        {
            propertiesGrid.Rows.Add(selected, name, value);
        }

        public void SortExtendedProperties()
        {
            propertiesGrid.Sort(propertiesGrid.Columns["Key"], ListSortDirection.Ascending);
        }

        private Label explanationLabel;
        private TableLayoutPanel tableLayoutPanel1;

        public DataGridViewRowCollection extendedPropertiesRows
        {
            get { return propertiesGrid.Rows; }
        }
    }
}