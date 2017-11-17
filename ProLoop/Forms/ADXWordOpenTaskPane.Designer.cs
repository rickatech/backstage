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
            this.rbOrganizations = new System.Windows.Forms.RadioButton();
            this.rbProjects = new System.Windows.Forms.RadioButton();
            this.lblOrgsProjects = new System.Windows.Forms.Label();
            this.cboOrgProject = new System.Windows.Forms.ComboBox();
            this.cboClient = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cboMatter = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cboEditor = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cboContent = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cboDocName = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnOpen = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.label6 = new System.Windows.Forms.Label();
            this.tvwFolder = new System.Windows.Forms.TreeView();
            this.btnSettings = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // rbOrganizations
            // 
            this.rbOrganizations.AutoSize = true;
            this.rbOrganizations.Location = new System.Drawing.Point(9, 12);
            this.rbOrganizations.Name = "rbOrganizations";
            this.rbOrganizations.Size = new System.Drawing.Size(89, 17);
            this.rbOrganizations.TabIndex = 0;
            this.rbOrganizations.TabStop = true;
            this.rbOrganizations.Text = "Organizations";
            this.rbOrganizations.UseVisualStyleBackColor = true;
            this.rbOrganizations.CheckedChanged += new System.EventHandler(this.rbOrganizations_CheckedChanged);
            // 
            // rbProjects
            // 
            this.rbProjects.AutoSize = true;
            this.rbProjects.Location = new System.Drawing.Point(9, 35);
            this.rbProjects.Name = "rbProjects";
            this.rbProjects.Size = new System.Drawing.Size(58, 17);
            this.rbProjects.TabIndex = 1;
            this.rbProjects.TabStop = true;
            this.rbProjects.Text = "Project";
            this.rbProjects.UseVisualStyleBackColor = true;
            this.rbProjects.CheckedChanged += new System.EventHandler(this.radioButtonProject_CheckedChanged);
            // 
            // lblOrgsProjects
            // 
            this.lblOrgsProjects.AutoSize = true;
            this.lblOrgsProjects.Location = new System.Drawing.Point(9, 56);
            this.lblOrgsProjects.Name = "lblOrgsProjects";
            this.lblOrgsProjects.Size = new System.Drawing.Size(107, 13);
            this.lblOrgsProjects.TabIndex = 2;
            this.lblOrgsProjects.Text = "Oranization /Project :";
            // 
            // cboOrgProject
            // 
            this.cboOrgProject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboOrgProject.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboOrgProject.FormattingEnabled = true;
            this.cboOrgProject.Location = new System.Drawing.Point(9, 73);
            this.cboOrgProject.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboOrgProject.Name = "cboOrgProject";
            this.cboOrgProject.Size = new System.Drawing.Size(320, 21);
            this.cboOrgProject.TabIndex = 3;
            this.cboOrgProject.SelectedIndexChanged += new System.EventHandler(this.cboOrgProject_SelectedIndexChanged);
            // 
            // cboClient
            // 
            this.cboClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboClient.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboClient.FormattingEnabled = true;
            this.cboClient.Location = new System.Drawing.Point(9, 117);
            this.cboClient.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboClient.Name = "cboClient";
            this.cboClient.Size = new System.Drawing.Size(320, 21);
            this.cboClient.TabIndex = 5;
            this.cboClient.SelectedIndexChanged += new System.EventHandler(this.cboClient_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(9, 101);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Client :";
            // 
            // cboMatter
            // 
            this.cboMatter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboMatter.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboMatter.FormattingEnabled = true;
            this.cboMatter.Location = new System.Drawing.Point(11, 160);
            this.cboMatter.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboMatter.Name = "cboMatter";
            this.cboMatter.Size = new System.Drawing.Size(320, 21);
            this.cboMatter.TabIndex = 7;
            this.cboMatter.SelectedIndexChanged += new System.EventHandler(this.cboMatter_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(9, 144);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Matter :";
            // 
            // cboEditor
            // 
            this.cboEditor.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboEditor.Enabled = false;
            this.cboEditor.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboEditor.FormattingEnabled = true;
            this.cboEditor.Location = new System.Drawing.Point(9, 350);
            this.cboEditor.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboEditor.Name = "cboEditor";
            this.cboEditor.Size = new System.Drawing.Size(320, 21);
            this.cboEditor.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(9, 334);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(92, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "Document Editor :";
            // 
            // cboContent
            // 
            this.cboContent.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboContent.Enabled = false;
            this.cboContent.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboContent.FormattingEnabled = true;
            this.cboContent.Location = new System.Drawing.Point(9, 393);
            this.cboContent.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboContent.Name = "cboContent";
            this.cboContent.Size = new System.Drawing.Size(320, 21);
            this.cboContent.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(9, 377);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(99, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Document Content:";
            // 
            // cboDocName
            // 
            this.cboDocName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboDocName.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.cboDocName.FormattingEnabled = true;
            this.cboDocName.Location = new System.Drawing.Point(9, 436);
            this.cboDocName.Name = "cboDocName";
            this.cboDocName.Size = new System.Drawing.Size(320, 21);
            this.cboDocName.TabIndex = 13;
            this.cboDocName.SelectedIndexChanged += new System.EventHandler(this.cboDocName_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(9, 420);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(90, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "Document Name:";
            // 
            // btnOpen
            // 
            this.btnOpen.Enabled = false;
            this.btnOpen.Location = new System.Drawing.Point(15, 475);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(75, 23);
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
            this.panel1.Location = new System.Drawing.Point(9, 189);
            this.panel1.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(320, 24);
            this.panel1.TabIndex = 17;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1, 2);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(41, 13);
            this.label6.TabIndex = 0;
            this.label6.Text = "Folders";
            // 
            // tvwFolder
            // 
            this.tvwFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tvwFolder.Location = new System.Drawing.Point(9, 213);
            this.tvwFolder.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.tvwFolder.Name = "tvwFolder";
            this.tvwFolder.Size = new System.Drawing.Size(320, 118);
            this.tvwFolder.TabIndex = 18;
            this.tvwFolder.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwFolder_AfterSelect);
            // 
            // btnSettings
            // 
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Image = global::ProLoop.WordAddin.Properties.Resources.opSetting;
            this.btnSettings.Location = new System.Drawing.Point(218, 15);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(42, 30);
            this.btnSettings.TabIndex = 19;
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // ADXWordOpenTaskPane
            // 
            this.ClientSize = new System.Drawing.Size(350, 535);
            this.Controls.Add(this.btnSettings);
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
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXWordOpenTaskPane";
            this.Text = "Open";
            //this.Activated += new System.EventHandler(this.ADXWordOpenTaskPane_Activated);
            //this.Load += new System.EventHandler(this.ADXWordOpenTaskPane_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.RadioButton rbOrganizations;
        private System.Windows.Forms.RadioButton rbProjects;
        private System.Windows.Forms.Label lblOrgsProjects;
        private System.Windows.Forms.ComboBox cboOrgProject;
        private System.Windows.Forms.ComboBox cboClient;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cboMatter;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cboEditor;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cboContent;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cboDocName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TreeView tvwFolder;
        private System.Windows.Forms.Button btnSettings;
    }
}
