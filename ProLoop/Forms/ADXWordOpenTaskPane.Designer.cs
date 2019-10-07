namespace ProLoop.WordAddin.Forms
{
    partial class ADXWordOpenTaskPane
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;
  
        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }
  
        #region Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXWordOpenTaskPane));
            this.rbOrganizations = new System.Windows.Forms.RadioButton();
            this.rbProjects = new System.Windows.Forms.RadioButton();
            this.lblOrgsProjects = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cboEditor = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.btnOpen = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.tvwFolder = new System.Windows.Forms.TreeView();
            this.buttonSearch = new System.Windows.Forms.Button();
            this.textBoxContent = new System.Windows.Forms.TextBox();
            this.textBoxEditor = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.buttonFullDocSearch = new System.Windows.Forms.Button();
            this.cboDocName = new ProLoop.EasyCompletionComboBox();
            this.cboContent = new ProLoop.EasyCompletionComboBox();
            this.cboMatter = new ProLoop.EasyCompletionComboBox();
            this.cboClient = new ProLoop.EasyCompletionComboBox();
            this.cboOrgProject = new ProLoop.EasyCompletionComboBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbOrganizations
            // 
            this.rbOrganizations.AutoSize = true;
            this.rbOrganizations.Location = new System.Drawing.Point(11, 15);
            this.rbOrganizations.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.rbOrganizations.Name = "rbOrganizations";
            this.rbOrganizations.Size = new System.Drawing.Size(117, 21);
            this.rbOrganizations.TabIndex = 0;
            this.rbOrganizations.TabStop = true;
            this.rbOrganizations.Text = "Organizations";
            this.rbOrganizations.UseVisualStyleBackColor = true;
            this.rbOrganizations.CheckedChanged += new System.EventHandler(this.rbOrganizations_CheckedChanged);
            // 
            // rbProjects
            // 
            this.rbProjects.AutoSize = true;
            this.rbProjects.Location = new System.Drawing.Point(11, 44);
            this.rbProjects.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.rbProjects.Name = "rbProjects";
            this.rbProjects.Size = new System.Drawing.Size(73, 21);
            this.rbProjects.TabIndex = 1;
            this.rbProjects.TabStop = true;
            this.rbProjects.Text = "Project";
            this.rbProjects.UseVisualStyleBackColor = true;
            this.rbProjects.CheckedChanged += new System.EventHandler(this.radioButtonProject_CheckedChanged);
            // 
            // lblOrgsProjects
            // 
            this.lblOrgsProjects.AutoSize = true;
            this.lblOrgsProjects.Location = new System.Drawing.Point(11, 70);
            this.lblOrgsProjects.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblOrgsProjects.Name = "lblOrgsProjects";
            this.lblOrgsProjects.Size = new System.Drawing.Size(141, 17);
            this.lblOrgsProjects.TabIndex = 2;
            this.lblOrgsProjects.Text = "Oranization /Project :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 126);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(51, 17);
            this.label1.TabIndex = 4;
            this.label1.Text = "Client :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 180);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(56, 17);
            this.label2.TabIndex = 6;
            this.label2.Text = "Matter :";
            // 
            // cboEditor
            // 
            this.cboEditor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboEditor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboEditor.FormattingEnabled = true;
            this.cboEditor.Items.AddRange(new object[] {
            "Rickatech",
            "Jullia",
            "Rick Armstrong"});
            this.cboEditor.Location = new System.Drawing.Point(11, 438);
            this.cboEditor.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.cboEditor.Name = "cboEditor";
            this.cboEditor.Size = new System.Drawing.Size(399, 25);
            this.cboEditor.TabIndex = 9;
            this.cboEditor.Visible = false;
            this.cboEditor.SelectedIndexChanged += new System.EventHandler(this.cboEditor_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 418);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(121, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Document Editor :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(11, 471);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(129, 17);
            this.label4.TabIndex = 10;
            this.label4.Text = "Document Content:";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(11, 525);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(117, 17);
            this.label5.TabIndex = 12;
            this.label5.Text = "Document Name:";
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(11, 594);
            this.btnOpen.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(94, 29);
            this.btnOpen.TabIndex = 14;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label6);
            this.panel1.Location = new System.Drawing.Point(11, 236);
            this.panel1.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(400, 30);
            this.panel1.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1, 2);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(55, 17);
            this.label6.TabIndex = 0;
            this.label6.Text = "Folders";
            // 
            // tvwFolder
            // 
            this.tvwFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tvwFolder.Location = new System.Drawing.Point(11, 266);
            this.tvwFolder.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.tvwFolder.Name = "tvwFolder";
            this.tvwFolder.Size = new System.Drawing.Size(399, 146);
            this.tvwFolder.TabIndex = 18;
            this.tvwFolder.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvwFolder_BeforeExpand);
            this.tvwFolder.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwFolder_AfterSelect);
            // 
            // buttonSearch
            // 
            this.buttonSearch.Location = new System.Drawing.Point(138, 594);
            this.buttonSearch.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.buttonSearch.Name = "buttonSearch";
            this.buttonSearch.Size = new System.Drawing.Size(94, 29);
            this.buttonSearch.TabIndex = 20;
            this.buttonSearch.Text = "Search";
            this.buttonSearch.UseVisualStyleBackColor = true;
            this.buttonSearch.Click += new System.EventHandler(this.buttonSearch_Click);
            // 
            // textBoxContent
            // 
            this.textBoxContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxContent.Location = new System.Drawing.Point(11, 491);
            this.textBoxContent.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxContent.Name = "textBoxContent";
            this.textBoxContent.Size = new System.Drawing.Size(399, 23);
            this.textBoxContent.TabIndex = 21;
            // 
            // textBoxEditor
            // 
            this.textBoxEditor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxEditor.Location = new System.Drawing.Point(11, 439);
            this.textBoxEditor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.textBoxEditor.Name = "textBoxEditor";
            this.textBoxEditor.Size = new System.Drawing.Size(399, 23);
            this.textBoxEditor.TabIndex = 22;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label7.Location = new System.Drawing.Point(202, 4);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(50, 19);
            this.label7.TabIndex = 23;
            this.label7.Text = "label7";
            this.label7.Visible = false;
            // 
            // buttonFullDocSearch
            // 
            this.buttonFullDocSearch.Location = new System.Drawing.Point(266, 594);
            this.buttonFullDocSearch.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.buttonFullDocSearch.Name = "buttonFullDocSearch";
            this.buttonFullDocSearch.Size = new System.Drawing.Size(121, 29);
            this.buttonFullDocSearch.TabIndex = 24;
            this.buttonFullDocSearch.Text = "Full Doc Search";
            this.buttonFullDocSearch.UseVisualStyleBackColor = true;
            this.buttonFullDocSearch.Click += new System.EventHandler(this.buttonFullDocSearch_Click);
            // 
            // cboDocName
            // 
            this.cboDocName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboDocName.FormattingEnabled = true;
            this.cboDocName.Location = new System.Drawing.Point(11, 545);
            this.cboDocName.Margin = new System.Windows.Forms.Padding(4);
            this.cboDocName.Name = "cboDocName";
            this.cboDocName.Size = new System.Drawing.Size(399, 25);
            this.cboDocName.TabIndex = 13;
            this.cboDocName.SelectedIndexChanged += new System.EventHandler(this.cboDocName_SelectedIndexChanged);
            // 
            // cboContent
            // 
            this.cboContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboContent.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboContent.FormattingEnabled = true;
            this.cboContent.Location = new System.Drawing.Point(11, 491);
            this.cboContent.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.cboContent.Name = "cboContent";
            this.cboContent.Size = new System.Drawing.Size(399, 25);
            this.cboContent.TabIndex = 11;
            this.cboContent.Visible = false;
            // 
            // cboMatter
            // 
            this.cboMatter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboMatter.FormattingEnabled = true;
            this.cboMatter.Location = new System.Drawing.Point(14, 200);
            this.cboMatter.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.cboMatter.Name = "cboMatter";
            this.cboMatter.Size = new System.Drawing.Size(399, 25);
            this.cboMatter.TabIndex = 7;
            this.cboMatter.SelectedIndexChanged += new System.EventHandler(this.cboMatter_SelectedIndexChanged);
            // 
            // cboClient
            // 
            this.cboClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboClient.FormattingEnabled = true;
            this.cboClient.Location = new System.Drawing.Point(11, 146);
            this.cboClient.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.cboClient.Name = "cboClient";
            this.cboClient.Size = new System.Drawing.Size(399, 25);
            this.cboClient.TabIndex = 5;
            this.cboClient.SelectedIndexChanged += new System.EventHandler(this.cboClient_SelectedIndexChanged);
            // 
            // cboOrgProject
            // 
            this.cboOrgProject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboOrgProject.FormattingEnabled = true;
            this.cboOrgProject.Location = new System.Drawing.Point(11, 91);
            this.cboOrgProject.Margin = new System.Windows.Forms.Padding(12, 4, 12, 4);
            this.cboOrgProject.Name = "cboOrgProject";
            this.cboOrgProject.Size = new System.Drawing.Size(399, 25);
            this.cboOrgProject.TabIndex = 3;
            this.cboOrgProject.SelectedIndexChanged += new System.EventHandler(this.cboOrgProject_SelectedIndexChanged);
            // 
            // ADXWordOpenTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.ClientSize = new System.Drawing.Size(438, 669);
            this.Controls.Add(this.buttonFullDocSearch);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.textBoxEditor);
            this.Controls.Add(this.textBoxContent);
            this.Controls.Add(this.buttonSearch);
            this.Controls.Add(this.tvwFolder);
            this.Controls.Add(this.panel1);
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
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "ADXWordOpenTaskPane";
            this.Text = "Open";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.RadioButton rbOrganizations;
        private System.Windows.Forms.RadioButton rbProjects;
        private System.Windows.Forms.Label lblOrgsProjects;
        //private System.Windows.Forms.ComboBox cboOrgProject;
        private EasyCompletionComboBox cboOrgProject;
        //private System.Windows.Forms.ComboBox cboClient;
        private EasyCompletionComboBox cboClient;
        private System.Windows.Forms.Label label1;
        // private System.Windows.Forms.ComboBox cboMatter;
        private EasyCompletionComboBox cboMatter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboEditor;
        private System.Windows.Forms.Label label3;
        //private System.Windows.Forms.ComboBox cboContent;
        private EasyCompletionComboBox cboContent;
        private System.Windows.Forms.Label label4;
        //private System.Windows.Forms.ComboBox cboDocName;
        private EasyCompletionComboBox cboDocName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TreeView tvwFolder;
        private System.Windows.Forms.Button buttonSearch;
        private System.Windows.Forms.TextBox textBoxContent;
        private System.Windows.Forms.TextBox textBoxEditor;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button buttonFullDocSearch;
    }
}
