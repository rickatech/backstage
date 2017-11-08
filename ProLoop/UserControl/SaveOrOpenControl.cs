using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AddinExpress.MSO;

namespace ProLoop.WordAddin.Forms
{
    public partial class SaveOrOpenControl : UserControl
    {
        private readonly AddinModule AddinCurrentInstance;
        public SaveOrOpenControl()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
        }

        private void SaveOrOpenControl_Load(object sender, EventArgs e)
        {
            //if(AddinCurrentInstance.Mode==Operation.Open)
            //{
            //    //Enable some control
            //}
            //else
            //{
            //    //Handle Save Buttons
            //}
        }
    }
}
