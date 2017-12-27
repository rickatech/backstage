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
using ProLoop.WordAddin.Utils;

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

        private void rbOrganizations_CheckedChanged(object sender, EventArgs e)
        {
            cboClient.Enabled = true;
            var item = new ComboboxItem()
            {
                Text = "Test",
                index = 0,
                itemType="Organization"
            };
            cboOrgProject.Items.Clear();
            cboOrgProject.Items.Add(item);
            //cboOrgProject.SelectedIndex = 0;
        }

        private void rbProjects_CheckedChanged(object sender, EventArgs e)
        {
            cboClient.Enabled = false;
            cboMatter.Enabled = false;

            var item = new ComboboxItem()
            {
                Text = "Test Project",
                index = 0,
                itemType = "project"
            };
            cboOrgProject.Items.Clear();
            cboOrgProject.Items.Add(item);
            //cboOrgProject.SelectedIndex = 0;
        }

        private void cboOrgProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            var item = cboOrgProject.SelectedItem as ComboboxItem;
            if (item != null && item.itemType == "project")
            {
               
            }
            else
            {
                cboClient.Items.Clear();
                cboClient.Enabled = true;
                var combobox = new ComboboxItem()
                {
                    Text = "Client 1",
                    index = 0,
                    itemType = "client"
                };
                cboClient.Items.Add(combobox);
                combobox = new ComboboxItem()
                {
                    Text = "Client 2",
                    index = 0,
                    itemType = "client"
                };
                cboClient.Items.Add(combobox);
            }
            DummyFileUploader();
            // cboClient.SelectedIndex = 0;
        }

        private void cboClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            var item = cboClient.SelectedItem as ComboboxItem;
            cboMatter.Enabled = true;
            cboMatter.Items.Clear();
            if (item!=null && item.Text=="Client 1")
            {
               var combobox = new ComboboxItem()
                {
                    Text = "Matter 1",
                    index = 0,
                    itemType = "matter"
                };
                cboMatter.Items.Add(combobox);
                 combobox = new ComboboxItem()
                {
                    Text = "Matter 2",
                    index = 0,
                    itemType = "matter"
                };
                cboMatter.Items.Add(combobox);
            }
            else
            {
                var combobox = new ComboboxItem()
                {
                    Text = "Matter 3",
                    index = 0,
                    itemType = "matter"
                };
                cboMatter.Items.Add(combobox);
            }
            DummyFileUploader();
        }

        private void DummyFileUploader()
        {
            var dumyItem = new List<ComboboxItem>();
            if (cboOrgProject.Text == "Test" &&
                 string.IsNullOrEmpty(cboClient.Text) && string.IsNullOrEmpty(cboMatter.Text)
                 && string.IsNullOrEmpty(cboContent.Text) && string.IsNullOrEmpty(cboEditor.Text))
            {
                dumyItem.Clear();
                var item = new ComboboxItem()
                {
                    Text = "doc1_newyork_bloom.txt",
                    itemType = "C1/CM1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc2_boston_citi.txt",
                    itemType = "C1/CM2/M2F1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_telsa_nv.txt",
                    itemType = "C1/CM2/M2F1/M2F1C"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_apple.txt",
                    itemType = "C2/CM3/M3F1"
                };
                dumyItem.Add(item);
                ProcessDummyItem(dumyItem);
            }
            else if (cboOrgProject.Text == "Test" &&
                 cboClient.Text=="Client 1" && string.IsNullOrEmpty(cboMatter.Text)
                 && string.IsNullOrEmpty(cboContent.Text) && string.IsNullOrEmpty(cboEditor.Text))
            {
                dumyItem.Clear();
                var item = new ComboboxItem()
                {
                    Text = "doc1_newyork_bloom.txt",
                    itemType = "C1/CM1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc2_boston_citi.txt",
                    itemType = "C1/CM2/M2F1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_telsa_nv.txt",
                    itemType = "C1/CM2/M2F1/M2F1C"
                };
                dumyItem.Add(item);
                ProcessDummyItem(dumyItem);
            }
            else if (cboOrgProject.Text == "Test" &&
                cboClient.Text=="Client 1" && cboMatter.Text=="Matter 2"
                 && string.IsNullOrEmpty(cboContent.Text) && string.IsNullOrEmpty(cboEditor.Text))
            {
                dumyItem.Clear();

                var item = new ComboboxItem()
                {
                    Text = "doc2_boston_citi.txt",
                    itemType = "C1/CM2/M2F1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_telsa_nv.txt",
                    itemType = "C1/CM2/M2F1/M2F1C"
                };
                dumyItem.Add(item);
                ProcessDummyItem(dumyItem);
            }
            else if (cboOrgProject.Text == "Test" && cboContent.Text == "Merger")
            {
                dumyItem.Clear();
                var item = new ComboboxItem()
                {
                    Text = "doc2_boston_citi.txt",
                    itemType = "C1/CM2/M2F1"
                };
                dumyItem.Add(item);
                ProcessDummyItem(dumyItem);
            }
            else if (cboOrgProject.Text == "Test" && cboEditor.Text == "Julie")
            {
                dumyItem.Clear();
                var item = new ComboboxItem()
                {
                    Text = "doc1_newyork_bloom.txt",
                    itemType = "C1/CM1"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_telsa_nv.txt",
                    itemType = "C1/CM2/M2F1/M2F1C"
                };
                dumyItem.Add(item);
                item = new ComboboxItem()
                {
                    Text = "doc3_apple.txt",
                    itemType = "C2/CM3/M3F1"
                };
                dumyItem.Add(item);
                ProcessDummyItem(dumyItem);
            }
            else
            {
                ProcessDummyItem(dumyItem);
            }
        }
        private void ProcessDummyItem(List<ComboboxItem> items)
        {
            cboDocName.Items.Clear();
            dgvFolders.Rows.Clear();            
            dgvFolders.AllowUserToAddRows = true;
            foreach (var item in items)
            {
                DataGridViewRow row = (DataGridViewRow)dgvFolders.Rows[0].Clone();
                cboDocName.Items.Add(item);
                row.Cells[0].Value = item.itemType;
                row.Cells[1].Value = item.Text;
                dgvFolders.Rows.Add(row);
            }
            dgvFolders.AllowUserToAddRows = false;

        }

        private void cboMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            DummyFileUploader();
        }

        private void cboEditor_SelectedIndexChanged(object sender, EventArgs e)
        {
            DummyFileUploader();
        }

        private void cboContent_SelectedIndexChanged(object sender, EventArgs e)
        {
            DummyFileUploader();
        }

        private void dgvFolders_SelectionChanged(object sender, EventArgs e)
        {
            dgvFolders.ClearSelection();
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            using (var aheadSearch = new AheadSearchForm())
            {
                aheadSearch.ShowDialog();
            }
        }
    }
}
