namespace ProLoop.Dashboard
{
    partial class ADXWordTaskPaneOpen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ADXWordTaskPaneOpen));
            this.elementHost1 = new System.Windows.Forms.Integration.ElementHost();
            this.openView1 = new ProLoop.Views.OpenView(new ViewModels.OpenViewModel());
            this.SuspendLayout();
            // 
            // elementHost1
            // 
            this.elementHost1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.elementHost1.Location = new System.Drawing.Point(0, 0);
            this.elementHost1.Name = "elementHost1";
            this.elementHost1.Size = new System.Drawing.Size(300, 791);
            this.elementHost1.TabIndex = 0;
            this.elementHost1.Text = "elementHost1";
            this.elementHost1.Child = this.openView1;
            // 
            // ADXWordTaskPaneOpen
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(120F, 120F);
            this.ClientSize = new System.Drawing.Size(300, 791);
            this.Controls.Add(this.elementHost1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Location = new System.Drawing.Point(0, 0);
            this.Name = "ADXWordTaskPaneOpen";
            this.Text = "Open";
            this.ResumeLayout(false);

        }
        #endregion

        private System.Windows.Forms.Integration.ElementHost elementHost1;
        private Views.OpenView openView1;
    }
}
