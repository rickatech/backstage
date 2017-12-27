using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
  
namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordAltTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        public ADXWordAltTaskPane()
        {
            InitializeComponent();
        }

        private void ADXWordAltTaskPane_Activated(object sender, EventArgs e)
        {
            WordAddin.AddinModule.CurrentInstance.Mode = Operation.Alt;
        }
    }
}
