using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;


namespace ProLoop.Dashboard
{
    public partial class ADXWordTaskPaneInfo: AddinExpress.WD.ADXWordTaskPane
    {     

        public ADXWordTaskPaneInfo()
        {
            InitializeComponent();           
            this.Load += ADXWordTaskPaneInfo_Load;
        }

        private void ADXWordTaskPaneInfo_Load(object sender, EventArgs e)
        {
            
        }
    }
}
