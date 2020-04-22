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
            // lblOrgsProjects
            // 
            this.lblOrgsProjects.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblOrgsProjects.AutoSize = true;
            this.lblOrgsProjects.Location = new System.Drawing.Point(9, 55);
            this.lblOrgsProjects.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.lblOrgsProjects.Name = "lblOrgsProjects";
            this.lblOrgsProjects.Size = new System.Drawing.Size(149, 20);
            this.lblOrgsProjects.TabIndex = 41;
            this.lblOrgsProjects.Text = "Oranization /Project :";
            // 
            // rbProjects
            // 
            this.rbProjects.AutoSize = true;
            this.rbProjects.Location = new System.Drawing.Point(10, 29);
            this.rbProjects.Margin = new System.Windows.Forms.Padding(2);
            this.rbProjects.Name = "rbProjects";
            this.rbProjects.Size = new System.Drawing.Size(76, 24);
            this.rbProjects.TabIndex = 40;
            this.rbProjects.TabStop = true;
            this.rbProjects.Text = "Project";
            this.rbProjects.UseVisualStyleBackColor = true;
            this.rbProjects.CheckedChanged += new System.EventHandler(this.radioButtonProject_CheckedChanged);
            // 
            // rbOrganizations
            // 
            this.rbOrganizations.AutoSize = true;
            this.rbOrganizations.Location = new System.Drawing.Point(10, 7);
            this.rbOrganizations.Margin = new System.Windows.Forms.Padding(2);
            this.rbOrganizations.Name = "rbOrganizations";
            this.rbOrganizations.Size = new System.Drawing.Size(122, 24);
            this.rbOrganizations.TabIndex = 39;
            this.rbOrganizations.TabStop = true;
            this.rbOrganizations.Text = "Organizations";
            this.rbOrganizations.UseVisualStyleBackColor = true;
            this.rbOrganizations.CheckedChanged += new System.EventHandler(this.rbOrganizations_CheckedChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 169);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(60, 20);
            this.label2.TabIndex = 46;
            this.label2.Text = "Matter :";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 110);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(54, 20);
            this.label1.TabIndex = 44;
            this.label1.Text = "Client :";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(1, -2);
            this.label6.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 20);
            this.label6.TabIndex = 0;
            this.label6.Text = "Folders";
            // 
            // tvwFolder
            // 
            this.tvwFolder.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tvwFolder.Location = new System.Drawing.Point(13, 275);
            this.tvwFolder.Margin = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.tvwFolder.Name = "tvwFolder";
            this.tvwFolder.Size = new System.Drawing.Size(314, 99);
            this.tvwFolder.TabIndex = 49;
            this.tvwFolder.BeforeExpand += new System.Windows.Forms.TreeViewCancelEventHandler(this.tvwFolder_BeforeExpand);
            this.tvwFolder.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tvwFolder_AfterSelect);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label6);
            this.panel1.Location = new System.Drawing.Point(13, 255);
            this.panel1.Margin = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(314, 20);
            this.panel1.TabIndex = 48;
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Location = new System.Drawing.Point(68, 517);
            this.btnSaveAs.Margin = new System.Windows.Forms.Padding(2);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(101, 27);
            this.btnSaveAs.TabIndex = 50;
            this.btnSaveAs.Text = "Save As";
            this.btnSaveAs.UseVisualStyleBackColor = true;
            this.btnSaveAs.Click += new System.EventHandler(this.btnSaveAs_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 381);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 20);
            this.label3.TabIndex = 51;
            this.label3.Text = "FileName";
            // 
            // txtboxFileName
            // 
            this.txtboxFileName.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtboxFileName.Location = new System.Drawing.Point(14, 405);
            this.txtboxFileName.Margin = new System.Windows.Forms.Padding(2);
            this.txtboxFileName.Name = "txtboxFileName";
            this.txtboxFileName.Size = new System.Drawing.Size(314, 27);
            this.txtboxFileName.TabIndex = 52;
            // 
            // cboMatter
            // 
            this.cboMatter.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboMatter.FormattingEnabled = true;
            this.cboMatter.Location = new System.Drawing.Point(14, 195);
            this.cboMatter.Margin = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.cboMatter.Name = "cboMatter";
            this.cboMatter.Size = new System.Drawing.Size(314, 28);
            this.cboMatter.TabIndex = 47;
            this.cboMatter.SelectedIndexChanged += new System.EventHandler(this.cboMatter_SelectedIndexChanged);
            // 
            // cboClient
            // 
            this.cboClient.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboClient.FormattingEnabled = true;
            this.cboClient.Location = new System.Drawing.Point(17, 137);
            this.cboClient.Margin = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.cboClient.Name = "cboClient";
            this.cboClient.Size = new System.Drawing.Size(314, 28);
            this.cboClient.TabIndex = 45;
            this.cboClient.SelectedIndexChanged += new System.EventHandler(this.cboClient_SelectedIndexChanged);
            // 
            // cboOrgProject
            // 
            this.cboOrgProject.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboOrgProject.FormattingEnabled = true;
            this.cboOrgProject.Location = new System.Drawing.Point(13, 80);
            this.cboOrgProject.Margin = new System.Windows.Forms.Padding(8, 2, 8, 2);
            this.cboOrgProject.Name = "cboOrgProject";
            this.cboOrgProject.Size = new System.Drawing.Size(314, 28);
            this.cboOrgProject.TabIndex = 43;
            this.cboOrgProject.SelectedIndexChanged += new System.EventHandler(this.cboOrgProject_SelectedIndexChanged);
            // 
            // ADXWordSaveTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.ClientSize = new System.Drawing.Size(339, 688);
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
            this.Controls.Add(this.lblOrgsProjects);
            this.Controls.Add(this.rbProjects);
            this.Controls.Add(this.rbOrganizations);
            this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Margin = new System.Windows.Forms.Padding(2);
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
