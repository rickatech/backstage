namespace ProLoop.WordAddin.Forms
{
    partial class ADXWordInfoTaskPane
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXWordInfoTaskPane));
            this.label1 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonSum = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxSummary = new System.Windows.Forms.TextBox();
            this.btnSettings = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.labelDetail = new System.Windows.Forms.Label();
            this.labelMessage = new System.Windows.Forms.Label();
            this.buttonCheckin = new System.Windows.Forms.Button();
            this.buttonCheckout = new System.Windows.Forms.Button();
            this.labelCheckin = new System.Windows.Forms.Label();
            this.buttonRefresh = new System.Windows.Forms.Button();
            this.buttonSave = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.labelEditors = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.lblVersion = new System.Windows.Forms.Label();
            this.lbDocId = new System.Windows.Forms.Label();
            this.lblMatter = new System.Windows.Forms.Label();
            this.lblClient = new System.Windows.Forms.Label();
            this.lblorg = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(20, 27);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 25);
            this.label1.TabIndex = 0;
            this.label1.Text = "label1";
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.buttonSum);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.panel1.Location = new System.Drawing.Point(22, 526);
            this.panel1.Margin = new System.Windows.Forms.Padding(15, 4, 15, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(483, 50);
            this.panel1.TabIndex = 21;
            // 
            // buttonSum
            // 
            this.buttonSum.Location = new System.Drawing.Point(282, 2);
            this.buttonSum.Margin = new System.Windows.Forms.Padding(4);
            this.buttonSum.Name = "buttonSum";
            this.buttonSum.Size = new System.Drawing.Size(112, 38);
            this.buttonSum.TabIndex = 27;
            this.buttonSum.Text = "Edit";
            this.buttonSum.UseVisualStyleBackColor = true;
            this.buttonSum.Click += new System.EventHandler(this.buttonSum_Click);
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Segoe UI Semibold", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(6, 6);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(105, 28);
            this.label2.TabIndex = 0;
            this.label2.Text = "Keywords:";
            // 
            // textBoxSummary
            // 
            this.textBoxSummary.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxSummary.Location = new System.Drawing.Point(21, 578);
            this.textBoxSummary.Margin = new System.Windows.Forms.Padding(15, 4, 15, 4);
            this.textBoxSummary.Multiline = true;
            this.textBoxSummary.Name = "textBoxSummary";
            this.textBoxSummary.Size = new System.Drawing.Size(484, 76);
            this.textBoxSummary.TabIndex = 22;
            // 
            // btnSettings
            // 
            this.btnSettings.FlatAppearance.BorderSize = 0;
            this.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.btnSettings.Image = global::ProLoop.WordAddin.Properties.Resources.opSetting;
            this.btnSettings.Location = new System.Drawing.Point(363, 44);
            this.btnSettings.Margin = new System.Windows.Forms.Padding(4);
            this.btnSettings.Name = "btnSettings";
            this.btnSettings.Size = new System.Drawing.Size(63, 45);
            this.btnSettings.TabIndex = 20;
            this.btnSettings.UseVisualStyleBackColor = true;
            this.btnSettings.Click += new System.EventHandler(this.btnSettings_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(18, 54);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(64, 25);
            this.label4.TabIndex = 27;
            this.label4.Text = "label4";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.labelDetail);
            this.groupBox1.Controls.Add(this.labelMessage);
            this.groupBox1.Controls.Add(this.buttonCheckin);
            this.groupBox1.Controls.Add(this.buttonCheckout);
            this.groupBox1.Controls.Add(this.labelCheckin);
            this.groupBox1.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.groupBox1.Location = new System.Drawing.Point(24, 294);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox1.Size = new System.Drawing.Size(483, 170);
            this.groupBox1.TabIndex = 28;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Check in /Check out";
            // 
            // labelDetail
            // 
            this.labelDetail.AutoSize = true;
            this.labelDetail.Location = new System.Drawing.Point(12, 63);
            this.labelDetail.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelDetail.Name = "labelDetail";
            this.labelDetail.Size = new System.Drawing.Size(65, 28);
            this.labelDetail.TabIndex = 2;
            this.labelDetail.Text = "label5";
            this.labelDetail.Visible = false;
            // 
            // labelMessage
            // 
            this.labelMessage.AutoSize = true;
            this.labelMessage.Location = new System.Drawing.Point(134, 32);
            this.labelMessage.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelMessage.Name = "labelMessage";
            this.labelMessage.Size = new System.Drawing.Size(65, 28);
            this.labelMessage.TabIndex = 1;
            this.labelMessage.Text = "label6";
            // 
            // buttonCheckin
            // 
            this.buttonCheckin.Location = new System.Drawing.Point(352, 84);
            this.buttonCheckin.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCheckin.Name = "buttonCheckin";
            this.buttonCheckin.Size = new System.Drawing.Size(111, 38);
            this.buttonCheckin.TabIndex = 29;
            this.buttonCheckin.Text = "Check in";
            this.buttonCheckin.UseVisualStyleBackColor = true;
            this.buttonCheckin.Visible = false;
            this.buttonCheckin.Click += new System.EventHandler(this.buttonCheckin_Click);
            // 
            // buttonCheckout
            // 
            this.buttonCheckout.Location = new System.Drawing.Point(352, 30);
            this.buttonCheckout.Margin = new System.Windows.Forms.Padding(4);
            this.buttonCheckout.Name = "buttonCheckout";
            this.buttonCheckout.Size = new System.Drawing.Size(111, 38);
            this.buttonCheckout.TabIndex = 30;
            this.buttonCheckout.Text = "Checkout";
            this.buttonCheckout.UseVisualStyleBackColor = true;
            this.buttonCheckout.Visible = false;
            this.buttonCheckout.Click += new System.EventHandler(this.buttonCheckout_Click);
            // 
            // labelCheckin
            // 
            this.labelCheckin.AutoSize = true;
            this.labelCheckin.Location = new System.Drawing.Point(10, 32);
            this.labelCheckin.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelCheckin.Name = "labelCheckin";
            this.labelCheckin.Size = new System.Drawing.Size(65, 28);
            this.labelCheckin.TabIndex = 0;
            this.labelCheckin.Text = "label5";
            // 
            // buttonRefresh
            // 
            this.buttonRefresh.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonRefresh.Location = new System.Drawing.Point(28, 669);
            this.buttonRefresh.Margin = new System.Windows.Forms.Padding(4);
            this.buttonRefresh.Name = "buttonRefresh";
            this.buttonRefresh.Size = new System.Drawing.Size(114, 38);
            this.buttonRefresh.TabIndex = 31;
            this.buttonRefresh.Text = "Refresh";
            this.buttonRefresh.UseVisualStyleBackColor = true;
            this.buttonRefresh.Click += new System.EventHandler(this.buttonRefresh_Click);
            // 
            // buttonSave
            // 
            this.buttonSave.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.buttonSave.Location = new System.Drawing.Point(358, 667);
            this.buttonSave.Margin = new System.Windows.Forms.Padding(4);
            this.buttonSave.Name = "buttonSave";
            this.buttonSave.Size = new System.Drawing.Size(112, 38);
            this.buttonSave.TabIndex = 32;
            this.buttonSave.Text = "Save";
            this.buttonSave.UseVisualStyleBackColor = true;
            this.buttonSave.Click += new System.EventHandler(this.buttonSave_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(24, 478);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 28);
            this.label3.TabIndex = 33;
            this.label3.Text = "Editors:";
            // 
            // labelEditors
            // 
            this.labelEditors.AutoSize = true;
            this.labelEditors.Location = new System.Drawing.Point(116, 478);
            this.labelEditors.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.labelEditors.Name = "labelEditors";
            this.labelEditors.Size = new System.Drawing.Size(65, 28);
            this.labelEditors.TabIndex = 34;
            this.labelEditors.Text = "label5";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.lblVersion);
            this.groupBox2.Controls.Add(this.lbDocId);
            this.groupBox2.Controls.Add(this.lblMatter);
            this.groupBox2.Controls.Add(this.lblClient);
            this.groupBox2.Controls.Add(this.lblorg);
            this.groupBox2.Location = new System.Drawing.Point(24, 82);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4);
            this.groupBox2.Size = new System.Drawing.Size(483, 202);
            this.groupBox2.TabIndex = 35;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Metadata Info";
            // 
            // lblVersion
            // 
            this.lblVersion.AutoSize = true;
            this.lblVersion.Location = new System.Drawing.Point(8, 144);
            this.lblVersion.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblVersion.Name = "lblVersion";
            this.lblVersion.Size = new System.Drawing.Size(80, 28);
            this.lblVersion.TabIndex = 4;
            this.lblVersion.Text = "Version:";
            // 
            // lbDocId
            // 
            this.lbDocId.AutoSize = true;
            this.lbDocId.Location = new System.Drawing.Point(8, 112);
            this.lbDocId.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbDocId.Name = "lbDocId";
            this.lbDocId.Size = new System.Drawing.Size(75, 28);
            this.lbDocId.TabIndex = 3;
            this.lbDocId.Text = "Doc ID:";
            // 
            // lblMatter
            // 
            this.lblMatter.AutoSize = true;
            this.lblMatter.Location = new System.Drawing.Point(9, 82);
            this.lblMatter.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblMatter.Name = "lblMatter";
            this.lblMatter.Size = new System.Drawing.Size(75, 28);
            this.lblMatter.TabIndex = 2;
            this.lblMatter.Text = "Matter:";
            // 
            // lblClient
            // 
            this.lblClient.AutoSize = true;
            this.lblClient.Location = new System.Drawing.Point(9, 54);
            this.lblClient.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblClient.Name = "lblClient";
            this.lblClient.Size = new System.Drawing.Size(66, 28);
            this.lblClient.TabIndex = 1;
            this.lblClient.Text = "Client:";
            // 
            // lblorg
            // 
            this.lblorg.AutoSize = true;
            this.lblorg.Location = new System.Drawing.Point(9, 26);
            this.lblorg.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lblorg.Name = "lblorg";
            this.lblorg.Size = new System.Drawing.Size(115, 28);
            this.lblorg.TabIndex = 0;
            this.lblorg.Text = "Org/Project";
            // 
            // ADXWordInfoTaskPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(144F, 144F);
            this.ClientSize = new System.Drawing.Size(527, 736);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.labelEditors);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.buttonSave);
            this.Controls.Add(this.buttonRefresh);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBoxSummary);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnSettings);
            this.Controls.Add(this.label1);
            this.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ADXWordInfoTaskPane";
            this.Text = "Info";
            this.Activated += new System.EventHandler(this.ADXWordInfoTaskPane_Activated);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnSettings;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button buttonSum;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxSummary;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label labelMessage;
        private System.Windows.Forms.Label labelCheckin;
        private System.Windows.Forms.Label labelDetail;
        private System.Windows.Forms.Button buttonCheckin;
        private System.Windows.Forms.Button buttonCheckout;
        private System.Windows.Forms.Button buttonRefresh;
        private System.Windows.Forms.Button buttonSave;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label labelEditors;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label lblVersion;
        private System.Windows.Forms.Label lbDocId;
        private System.Windows.Forms.Label lblMatter;
        private System.Windows.Forms.Label lblClient;
        private System.Windows.Forms.Label lblorg;
    }
}
