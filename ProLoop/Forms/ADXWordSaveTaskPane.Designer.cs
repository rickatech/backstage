namespace ProLoop.WordAddin.Forms
{
    partial class ADXWordSaveTaskPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXWordSaveTaskPane));
            this.btnSettings = new System.Windows.Forms.Button();
            this.lblOrgsProjects = new System.Windows.Forms.Label();
            this.rbProjects = new System.Windows.Forms.RadioButton();
            this.rbOrganizations = new System.Windows.Forms.RadioButton();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.tvwFolder = new System.Windows.Forms.TreeView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtboxFileName = new System.Windows.Forms.TextBox();
            this.cboMatter = new ProLoop.EasyCompletionComboBox();
            this.cboClient = new ProLoop.EasyCompletionComboBox();
            this.cboOrgProject = new ProLoop.EasyCompletionComboBox();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSettings
            // 
            this.btnSettings.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Image = global::ProLoop.WordAddin.Properties.Resources.opSetting;
            this.btnSettings.Location = new System.Drawing.Point(272, 15);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(42, 30);
            this.btnSettings.TabIndex = 42;
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // lblOrgsProjects
            // 
            this.lblOrgsProjects.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrgsProjects.AutoSize = true;
            this.lblOrgsProjects.Location = new System.Drawing.Point(11, 59);
            this.lblOrgsProjects.Name = "lblOrgsProjects";
            this.lblOrgsProjects.Size = new System.Drawing.Size(107, 13);
            this.lblOrgsProjects.TabIndex = 41;
            this.lblOrgsProjects.Text = "Oranization /Project :";
            // 
            // rbProjects
            // 
            this.rbProjects.AutoSize = true;
            this.rbProjects.Location = new System.Drawing.Point(12, 35);
            this.rbProjects.Name = "rbProjects";
            this.rbProjects.Size = new System.Drawing.Size(58, 17);
            this.rbProjects.TabIndex = 40;
            this.rbProjects.TabStop = true;
            this.rbProjects.Text = "Project";
            this.rbProjects.UseVisualStyleBackColor = true;
            this.rbProjects.CheckedChanged += new System.EventHandler(this.radioButtonProject_CheckedChanged);
            // 
            // rbOrganizations
            // 
            this.rbOrganizations.AutoSize = true;
            this.rbOrganizations.Location = new System.Drawing.Point(12, 12);
            this.rbOrganizations.Name = "rbOrganizations";
            this.rbOrganizations.Size = new System.Drawing.Size(89, 17);
            this.rbOrganizations.TabIndex = 39;
            this.rbOrganizations.TabStop = true;
            this.rbOrganizations.Text = "Organizations";
            this.rbOrganizations.UseVisualStyleBackColor = true;
            this.rbOrganizations.CheckedChanged += new System.EventHandler(this.rbOrganizations_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(11, 146);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 46;
            this.label2.Text = "Matter :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(11, 103);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(39, 13);
            this.label1.TabIndex = 44;
            this.label1.Text = "Client :";
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
            this.tvwFolder.Location = new System.Drawing.Point(11, 215);
            this.tvwFolder.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.tvwFolder.Name = "tvwFolder";
            this.tvwFolder.Size = new System.Drawing.Size(320, 118);
            this.tvwFolder.TabIndex = 49;
            this.tvwFolder.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwFolder_AfterSelect);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label6);
            this.panel1.Location = new System.Drawing.Point(11, 191);
            this.panel1.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(320, 24);
            this.panel1.TabIndex = 48;
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Location = new System.Drawing.Point(75, 469);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(75, 23);
            this.btnSaveAs.TabIndex = 50;
            this.btnSaveAs.Text = "Save As";
            this.btnSaveAs.UseVisualStyleBackColor = true;
            this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(11, 346);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 13);
            this.label3.TabIndex = 51;
            this.label3.Text = "FileName";
            // 
            // txtboxFileName
            // 
            this.txtboxFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtboxFileName.Location = new System.Drawing.Point(11, 366);
            this.txtboxFileName.Name = "txtboxFileName";
            this.txtboxFileName.Size = new System.Drawing.Size(320, 20);
            this.txtboxFileName.TabIndex = 52;
            // 
            // cboMatter
            // 
            this.cboMatter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboMatter.FormattingEnabled = true;
            this.cboMatter.Location = new System.Drawing.Point(13, 162);
            this.cboMatter.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboMatter.Name = "cboMatter";
            this.cboMatter.Size = new System.Drawing.Size(320, 21);
            this.cboMatter.TabIndex = 47;
            this.cboMatter.SelectedIndexChanged += new System.EventHandler(this.cboMatter_SelectedIndexChanged);
            // 
            // cboClient
            // 
            this.cboClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboClient.FormattingEnabled = true;
            this.cboClient.Location = new System.Drawing.Point(11, 119);
            this.cboClient.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboClient.Name = "cboClient";
            this.cboClient.Size = new System.Drawing.Size(320, 21);
            this.cboClient.TabIndex = 45;
            this.cboClient.SelectedIndexChanged += new System.EventHandler(this.cboClient_SelectedIndexChanged);
            // 
            // cboOrgProject
            // 
            this.cboOrgProject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboOrgProject.FormattingEnabled = true;
            this.cboOrgProject.Location = new System.Drawing.Point(11, 75);
            this.cboOrgProject.Margin = new System.Windows.Forms.Padding(10, 3, 10, 3);
            this.cboOrgProject.Name = "cboOrgProject";
            this.cboOrgProject.Size = new System.Drawing.Size(320, 21);
            this.cboOrgProject.TabIndex = 43;
            this.cboOrgProject.SelectedIndexChanged += new System.EventHandler(this.cboOrgProject_SelectedIndexChanged);
            // 
            // ADXWordSaveTaskPane
            // 
            this.ClientSize = new System.Drawing.Size(350, 523);
            this.Controls.Add(this.txtboxFileName);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.btnSaveAs);
            this.Controls.Add(this.cboMatter);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.cboClient);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboOrgProject);
            this.Controls.Add(this.tvwFolder);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.lblOrgsProjects);
            this.Controls.Add(this.rbProjects);
            this.Controls.Add(this.rbOrganizations);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXWordSaveTaskPane";
            this.Text = "Save As";
            this.Activated += new System.EventHandler(this.ADXWordSaveTaskPane_Activated);
            this.Load += new System.EventHandler(this.ADXWordSaveTaskPane_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.Label lblOrgsProjects;
        private System.Windows.Forms.RadioButton rbProjects;
        private System.Windows.Forms.RadioButton rbOrganizations;
        private EasyCompletionComboBox cboMatter;
        private System.Windows.Forms.Label label2;
        private EasyCompletionComboBox cboClient;
        private System.Windows.Forms.Label label1;
        private EasyCompletionComboBox cboOrgProject;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TreeView tvwFolder;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtboxFileName;
    }
}
