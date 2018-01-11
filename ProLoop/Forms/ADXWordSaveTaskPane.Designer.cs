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
            this.saveOrOpenControl1 = new ProLoop.WordAddin.Forms.SaveOrOpenControl();
            this.SuspendLayout();
            // 
            // saveOrOpenControl1
            // 
            this.saveOrOpenControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.saveOrOpenControl1.Location = new System.Drawing.Point(0, 0);
            this.saveOrOpenControl1.Name = "saveOrOpenControl1";
            this.saveOrOpenControl1.Size = new System.Drawing.Size(300, 506);
            this.saveOrOpenControl1.TabIndex = 0;
            // 
            // ADXWordSaveTaskPane
            // 
            this.ClientSize = new System.Drawing.Size(300, 506);
            this.Controls.Add(this.saveOrOpenControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXWordSaveTaskPane";
            this.Text = "Save";
            this.Activated += new System.EventHandler(this.ADXWordSaveTaskPane_Activated);
            this.Load += new System.EventHandler(this.ADXWordSaveTaskPane_Load);
            this.ResumeLayout(false);

        }
        #endregion

        private SaveOrOpenControl saveOrOpenControl1;
    }
}
