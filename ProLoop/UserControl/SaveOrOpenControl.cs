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
        private BindingSource OrgsBindingSource;

        private BindingSource ProjectsBindingSource;

        private BindingSource ClientsBindingSource;

        private BindingSource MattersBindingSource;

        private BindingSource DocsBindingSource;

        private SettingsForm settingForm;

        private AutoCompleteStringCollection OrgsAutoCompleteCollection;

        private AutoCompleteStringCollection ClientsAutoCompleteCollection;

        private AutoCompleteStringCollection MattersAutoCompleteCollection;

        private AutoCompleteStringCollection ProjectsAutoCompleteCollection;

        private AutoCompleteStringCollection DocsAutoCompleteCollection;

        private Client ObjClient { get; set; } = new Client();
        private Project ObjProject { get; set; } = new Project();
        private Organization ObjOrganization { get; set; } = new Organization();
        private Matter ObjMatter { get; set; } = new Matter();

        private string FolderPath
        {
            get;
            set;
        }

        private string DocumentName
        {
            get;
            set;
        }
        public SaveOrOpenControl()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            this.OrgsBindingSource = new BindingSource();
            this.ProjectsBindingSource = new BindingSource();
            this.ClientsBindingSource = new BindingSource();
            this.MattersBindingSource = new BindingSource();
            this.DocsBindingSource = new BindingSource();
            this.OrgsAutoCompleteCollection = new AutoCompleteStringCollection();
            this.ProjectsAutoCompleteCollection = new AutoCompleteStringCollection();
            this.ClientsAutoCompleteCollection = new AutoCompleteStringCollection();
            this.MattersAutoCompleteCollection = new AutoCompleteStringCollection();
            this.DocsAutoCompleteCollection = new AutoCompleteStringCollection();
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
