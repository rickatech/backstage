using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using AddinExpress.MSO;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordSaveTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;
        public ADXWordSaveTaskPane()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
        }

        private void ADXWordSaveTaskPane_Load(object sender, EventArgs e)
        {
            AddinCurrentInstance.Mode = Operation.Open;
        }

        private void ADXWordSaveTaskPane_Activated(object sender, EventArgs e)
        {

        }
    }
}
