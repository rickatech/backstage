namespace ProLoop.WordAddin.Forms
{
    partial class SaveOrOpenControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle4 = new System.Windows.Forms.DataGridViewCellStyle();
            this.btnOpen = new System.Windows.Forms.Button();
            this.cboDocName = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.cboContent = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboEditor = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboMatter = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cboClient = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboOrgProject = new System.Windows.Forms.ComboBox();
            this.lblOrgsProjects = new System.Windows.Forms.Label();
            this.rbProjects = new System.Windows.Forms.RadioButton();
            this.rbOrganizations = new System.Windows.Forms.RadioButton();
            this.btnSettings = new System.Windows.Forms.Button();
            this.dgvFolders = new System.Windows.Forms.DataGridView();
            this.ColumnFolder = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ColumnFiles = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.buttonSearch = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvFolders)).BeginInit();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(39, 597);
            this.btnOpen.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(100, 28);
            this.btnOpen.TabIndex = 34;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            // 
            // cboDocName
            // 
            this.cboDocName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboDocName.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboDocName.FormattingEnabled = true;
            this.cboDocName.Location = new System.Drawing.Point(19, 549);
            this.cboDocName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboDocName.Name = "cboDocName";
            this.cboDocName.Size = new System.Drawing.Size(404, 24);
            this.cboDocName.TabIndex = 33;
            // 
            // label5
            // 
            this.label5.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(19, 529);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(117, 17);
            this.label5.TabIndex = 32;
            this.label5.Text = "Document Name:";
            // 
            // cboContent
            // 
            this.cboContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboContent.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboContent.FormattingEnabled = true;
            this.cboContent.Items.AddRange(new object[] {
            "Merger",
            "Info",
            "Error",
            "Important"});
            this.cboContent.Location = new System.Drawing.Point(19, 496);
            this.cboContent.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboContent.Name = "cboContent";
            this.cboContent.Size = new System.Drawing.Size(404, 24);
            this.cboContent.TabIndex = 31;
            this.cboContent.SelectedIndexChanged += new System.EventHandler(this.cboContent_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(19, 476);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(129, 17);
            this.label4.TabIndex = 30;
            this.label4.Text = "Document Content:";
            // 
            // cboEditor
            // 
            this.cboEditor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboEditor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboEditor.FormattingEnabled = true;
            this.cboEditor.Items.AddRange(new object[] {
            "Julie",
            "Client",
            "ZTech"});
            this.cboEditor.Location = new System.Drawing.Point(21, 443);
            this.cboEditor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboEditor.Name = "cboEditor";
            this.cboEditor.Size = new System.Drawing.Size(401, 24);
            this.cboEditor.TabIndex = 29;
            this.cboEditor.SelectedIndexChanged += new System.EventHandler(this.cboEditor_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(21, 423);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(121, 17);
            this.label3.TabIndex = 28;
            this.label3.Text = "Document Editor :";
            // 
            // cboMatter
            // 
            this.cboMatter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboMatter.Enabled = false;
            this.cboMatter.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboMatter.FormattingEnabled = true;
            this.cboMatter.Location = new System.Drawing.Point(23, 197);
            this.cboMatter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboMatter.Name = "cboMatter";
            this.cboMatter.Size = new System.Drawing.Size(400, 24);
            this.cboMatter.TabIndex = 27;
            this.cboMatter.SelectedIndexChanged += new System.EventHandler(this.cboMatter_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(23, 177);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 17);
            this.label2.TabIndex = 26;
            this.label2.Text = "Matter :";
            // 
            // cboClient
            // 
            this.cboClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboClient.Enabled = false;
            this.cboClient.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboClient.FormattingEnabled = true;
            this.cboClient.Location = new System.Drawing.Point(24, 143);
            this.cboClient.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboClient.Name = "cboClient";
            this.cboClient.Size = new System.Drawing.Size(396, 24);
            this.cboClient.TabIndex = 25;
            this.cboClient.SelectedIndexChanged += new System.EventHandler(this.cboClient_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(23, 124);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 17);
            this.label1.TabIndex = 24;
            this.label1.Text = "Client :";
            // 
            // cboOrgProject
            // 
            this.cboOrgProject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboOrgProject.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboOrgProject.FormattingEnabled = true;
            this.cboOrgProject.Location = new System.Drawing.Point(25, 92);
            this.cboOrgProject.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cboOrgProject.Name = "cboOrgProject";
            this.cboOrgProject.Size = new System.Drawing.Size(396, 24);
            this.cboOrgProject.TabIndex = 23;
            this.cboOrgProject.SelectedIndexChanged += new System.EventHandler(this.cboOrgProject_SelectedIndexChanged);
            // 
            // lblOrgsProjects
            // 
            this.lblOrgsProjects.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrgsProjects.AutoSize = true;
            this.lblOrgsProjects.Location = new System.Drawing.Point(21, 73);
            this.lblOrgsProjects.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblOrgsProjects.Name = "lblOrgsProjects";
            this.lblOrgsProjects.Size = new System.Drawing.Size(141, 17);
            this.lblOrgsProjects.TabIndex = 22;
            this.lblOrgsProjects.Text = "Oranization /Project :";
            // 
            // rbProjects
            // 
            this.rbProjects.AutoSize = true;
            this.rbProjects.Location = new System.Drawing.Point(23, 43);
            this.rbProjects.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.rbProjects.Name = "rbProjects";
            this.rbProjects.Size = new System.Drawing.Size(73, 21);
            this.rbProjects.TabIndex = 21;
            this.rbProjects.TabStop = true;
            this.rbProjects.Text = "Project";
            this.rbProjects.UseVisualStyleBackColor = true;
            this.rbProjects.CheckedChanged += new System.EventHandler(this.rbProjects_CheckedChanged);
            // 
            // rbOrganizations
            // 
            this.rbOrganizations.AutoSize = true;
            this.rbOrganizations.Location = new System.Drawing.Point(23, 15);
            this.rbOrganizations.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.rbOrganizations.Name = "rbOrganizations";
            this.rbOrganizations.Size = new System.Drawing.Size(117, 21);
            this.rbOrganizations.TabIndex = 20;
            this.rbOrganizations.TabStop = true;
            this.rbOrganizations.Text = "Organizations";
            this.rbOrganizations.UseVisualStyleBackColor = true;
            this.rbOrganizations.CheckedChanged += new System.EventHandler(this.rbOrganizations_CheckedChanged);
            // 
            // btnSettings
            // 
            this.btnSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Image = global::ProLoop.WordAddin.Properties.Resources.opSetting;
            this.btnSettings.Location = new System.Drawing.Point(369, 18);
            this.btnSettings.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(56, 37);
            this.btnSettings.TabIndex = 38;
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Visible = false;
            // 
            // dgvFolders
            // 
            this.dgvFolders.AllowUserToDeleteRows = false;
            this.dgvFolders.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(192)))), ((int)(((byte)(192)))), ((int)(((byte)(255)))));
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvFolders.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvFolders.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvFolders.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.ActiveCaption;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvFolders.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle2;
            this.dgvFolders.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvFolders.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ColumnFolder,
            this.ColumnFiles});
            this.dgvFolders.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dgvFolders.Location = new System.Drawing.Point(23, 231);
            this.dgvFolders.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dgvFolders.MultiSelect = false;
            this.dgvFolders.Name = "dgvFolders";
            this.dgvFolders.RowHeadersVisible = false;
            this.dgvFolders.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvFolders.Size = new System.Drawing.Size(401, 185);
            this.dgvFolders.TabIndex = 39;
            this.dgvFolders.SelectionChanged += new System.EventHandler(this.dgvFolders_SelectionChanged);
            // 
            // ColumnFolder
            // 
            dataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            dataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ColumnFolder.DefaultCellStyle = dataGridViewCellStyle3;
            this.ColumnFolder.HeaderText = "Folders";
            this.ColumnFolder.Name = "ColumnFolder";
            // 
            // ColumnFiles
            // 
            this.ColumnFiles.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            dataGridViewCellStyle4.SelectionBackColor = System.Drawing.Color.FromArgb(((int)(((byte)(224)))), ((int)(((byte)(224)))), ((int)(((byte)(224)))));
            dataGridViewCellStyle4.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.ColumnFiles.DefaultCellStyle = dataGridViewCellStyle4;
            this.ColumnFiles.HeaderText = "Documents";
            this.ColumnFiles.Name = "ColumnFiles";
            // 
            // buttonSearch
            // 
            this.buttonSearch.Location = new System.Drawing.Point(224, 593);
            this.buttonSearch.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.Size = new System.Drawing.Size(100, 28);
            this.buttonSearch.TabIndex = 40;
            this.buttonSearch.Text = "Search";
            this.buttonSearch.UseVisualStyleBackColor = true;
            this.buttonSearch.Click += new System.EventHandler(this.buttonSearch_Click);
            // 
            // SaveOrOpenControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.buttonSearch);
            this.Controls.Add(this.dgvFolders);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.btnOpen);
            this.Controls.Add(this.cboDocName);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.cboContent);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cboEditor);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cboMatter);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cboClient);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboOrgProject);
            this.Controls.Add(this.lblOrgsProjects);
            this.Controls.Add(this.rbProjects);
            this.Controls.Add(this.rbOrganizations);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "SaveOrOpenControl";
            this.Size = new System.Drawing.Size(469, 646);
            this.Load += new System.EventHandler(this.SaveOrOpenControl_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvFolders)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.ComboBox cboDocName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox cboContent;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboEditor;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboMatter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboClient;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboOrgProject;
        private System.Windows.Forms.Label lblOrgsProjects;
        private System.Windows.Forms.RadioButton rbProjects;
        private System.Windows.Forms.RadioButton rbOrganizations;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.DataGridView dgvFolders;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFolder;
        private System.Windows.Forms.DataGridViewTextBoxColumn ColumnFiles;
        private System.Windows.Forms.Button buttonSearch;
    }
}
