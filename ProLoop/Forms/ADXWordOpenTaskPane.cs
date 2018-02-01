using System;
using System.Windows.Forms;
using System.Collections.Generic;
using AddinExpress.MSO;
using Serilog;
using ProLoop.WordAddin.Service;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.Linq;
using ProLoop.WordAddin.Utils;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordOpenTaskPane : AddinExpress.WD.ADXWordTaskPane
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

        private string selectedPath = string.Empty;
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

        private SearchParameter searchParameter;

        public ADXWordOpenTaskPane()
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
            searchParameter = new SearchParameter();
        }

        private void rbOrganizations_CheckedChanged(object sender, EventArgs e)
        {
            Log.Debug("rbOrganizations_CheckedChanged() -- Begin");

            // Disable the Open button
            btnOpen.Enabled = false;
            cboDocName.Text = string.Empty;
            cboClient.Text = string.Empty;
            cboMatter.Text = string.Empty;
            RadioButton rbOrgs = sender as RadioButton;
            CheckProLoopSettings(rbOrgs);
            if (rbOrganizations.Checked)
            {
                // Change the label to Organizations
                lblOrgsProjects.Text = "Organizations";
                searchParameter.OrgOrProject = "Organizations";
                // Enable Client and Matter
                cboClient.Enabled = true;
                cboMatter.Enabled = true;

                // Clear the Client, Matter, Docs..
                cboClient.DataBindings.Clear();
                cboClient.DataSource = null;
                cboClient.Items.Clear();

                cboMatter.DataBindings.Clear();
                cboMatter.DataSource = null;
                cboMatter.Items.Clear();

                tvwFolder.Nodes.Clear();

                //cboEditor.DataBindings.Clear();
                //cboEditor.DataSource = null;
                //cboEditor.Items.Clear();

                cboContent.DataBindings.Clear();
                cboContent.DataSource = null;
                cboContent.Items.Clear();

                cboDocName.DataBindings.Clear();
                cboDocName.DataSource = null;
                cboDocName.Items.Clear();

                List<Organization> orgs = APIHelper.GetOrganizations();

                OrgsBindingSource.DataSource = orgs;
                cboOrgProject.DataSource = OrgsBindingSource;
                cboOrgProject.DisplayMember = "title";
                cboOrgProject.ValueMember = "id";

                cboOrgProject.AutoCompleteMode = AutoCompleteMode.Suggest;
                cboOrgProject.AutoCompleteSource = AutoCompleteSource.CustomSource;

                OrgsAutoCompleteCollection.AddRange(orgs.Select(org => org.Title).ToArray());
                cboOrgProject.AutoCompleteCustomSource = OrgsAutoCompleteCollection;

                cboOrgProject.SelectedIndex = -1;
                cboOrgProject.Text = "Select:";
            }
        }

        private void radioButtonProject_CheckedChanged(object sender, EventArgs e)
        {
            Log.Debug("rbProjects_CheckedChanged() -- Begin");
            cboDocName.Text = string.Empty;
            cboClient.Text = string.Empty;
            cboMatter.Text = string.Empty;
            // Disable the Open button
            btnOpen.Enabled = false;

            RadioButton rbProjs = sender as RadioButton;
            CheckProLoopSettings(rbProjs);

            // Projects checkbox selected
            if (rbProjects.Checked)
            {
                // Change the label to Organizations
                lblOrgsProjects.Text = "Projects";
                searchParameter.OrgOrProject = "Projects";
                // Disable Client and Matter
                cboClient.Enabled = false;
                ObjClient = null;
                ObjMatter = null;
                cboMatter.Enabled = false;

                // Clear the Client, Matter, Docs..
                cboClient.DataBindings.Clear();
                cboClient.DataSource = null;
                cboClient.Items.Clear();

                cboMatter.DataBindings.Clear();
                cboMatter.DataSource = null;
                cboMatter.Items.Clear();

                tvwFolder.Nodes.Clear();

                //cboEditor.DataBindings.Clear();
                //cboEditor.DataSource = null;
                //cboEditor.Items.Clear();

                cboContent.DataBindings.Clear();
                cboContent.DataSource = null;
                cboContent.Items.Clear();

                cboDocName.DataBindings.Clear();
                cboDocName.DataSource = null;
                cboDocName.Items.Clear();

                List<Project> projects = APIHelper.GetProjects();

                ProjectsBindingSource.DataSource = projects;
                cboOrgProject.DataSource = ProjectsBindingSource;
                cboOrgProject.DisplayMember = "title";
                cboOrgProject.ValueMember = "id";

                cboOrgProject.AutoCompleteMode = AutoCompleteMode.Suggest;
                cboOrgProject.AutoCompleteSource = AutoCompleteSource.CustomSource;

                ProjectsAutoCompleteCollection.AddRange(projects.Select(project => project.Title).ToArray());
                cboOrgProject.AutoCompleteCustomSource = ProjectsAutoCompleteCollection;

                cboOrgProject.SelectedIndex = -1;
                cboOrgProject.SelectedText = "Select:";
            }
            Log.Debug("rbProjects_CheckedChanged() -- End");
        }

        private void CheckProLoopSettings(RadioButton radioButton)
        {
            if (string.IsNullOrEmpty(AddinCurrentInstance.ProLoopUrl))
            {
                this.AddinCurrentInstance.LoadValuesFromRegistry();
            }
            if (string.IsNullOrEmpty(AddinCurrentInstance.ProLoopUrl) ||
                string.IsNullOrEmpty(AddinCurrentInstance.ProLoopUsername) ||
                string.IsNullOrEmpty(AddinCurrentInstance.ProLoopPassword))

            {
                //MessageBox.Show("Please update the ProLoop settings (URL, Username, and Password) to proceed.", "Missing/Invalid Settings", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                var settingForm = new SettingsForm();

                if (settingForm.ShowDialog() == DialogResult.OK)
                {
                    radioButton.Checked = false;
                    return;
                }
            }            
            if (!string.IsNullOrEmpty(this.AddinCurrentInstance.ProLoopUrl))
            {
                //    Uri uri = new Uri(this.AddinCurrentInstance.ProLoopUrl);
                //    bool flag5 = this.localHttpClient != null;
                //    if (flag5)
                //    {
                //        bool flag6 = this.localHttpClient.BaseAddress != uri;
                //        if (flag6)
                //        {
                //            this.localHttpClient.BaseAddress = uri;
                //        }
                //    }
                //    else
                //    {
                //        this.localHttpClient = new HttpClient
                //        {
                //            BaseAddress = uri,
                //            Timeout = TimeSpan.FromSeconds(30.0)
                //        };
                //        this.localHttpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("Please update the ProLoop settings (URL, Username, and Password) to proceed.", "Invalid settings", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                //    this.ShowSettingsForm();
                //}           
            }
        }

        private void cboOrgProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboOrgProject_SelectedIndexChanged() -- Begin");

            // Make sure the Combo box has the focus before processing the event handler.
            // This will fire the event only when the Combo box has focus
            ComboBox cbo = (ComboBox)sender;
            if (!cbo.Focused) return;

            Context context = rbOrganizations.Checked ? Context.Orgs : Context.Projects;

            if (cboOrgProject.SelectedItem is Organization)
            {
                ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                searchParameter.OrgOrProjectName = ObjOrganization.Title;
                ObjProject = null;
            }
            else if (cboOrgProject.SelectedItem is Project)
            {
                ObjProject = cboOrgProject.SelectedItem as Project;
                ObjOrganization = null;
                searchParameter.OrgOrProjectName = ObjProject.Title;
            }
            else
            {
                ObjProject = null;
                ObjOrganization = null;
            }


            if (context == Context.Orgs)
            {
                // Clear the Client, Matter, Editor, Content, Docs..
                cboClient.DataBindings.Clear();
                cboClient.DataSource = null;
                cboClient.Items.Clear();

                cboMatter.DataBindings.Clear();
                cboMatter.DataSource = null;
                cboMatter.Items.Clear();

                tvwFolder.Nodes.Clear();

                //cboEditor.DataBindings.Clear();
                //cboEditor.DataSource = null;
                //cboEditor.Items.Clear();

                cboContent.DataBindings.Clear();
                cboContent.DataSource = null;
                cboContent.Items.Clear();

                cboDocName.DataBindings.Clear();
                cboDocName.DataSource = null;
                cboDocName.Items.Clear();

                List<Client> clients = APIHelper.GetClients(ObjOrganization.Id);
                ClientsBindingSource.DataSource = clients;
                cboClient.DataSource = ClientsBindingSource;
                cboClient.DisplayMember = "name";
                cboClient.ValueMember = "id";
                //cboClient_SelectedIndexChanged(cboClient, null);
                //cboClient.SelectedIndex = 0;

                cboClient.AutoCompleteMode = AutoCompleteMode.Suggest;
                cboClient.AutoCompleteSource = AutoCompleteSource.CustomSource;

                ClientsAutoCompleteCollection.AddRange(clients.Select(client => client.Name).ToArray());
                cboClient.AutoCompleteCustomSource = ClientsAutoCompleteCollection;

                cboClient.SelectedIndex = -1;
                cboClient.SelectedText = "Select:";
            }
            else
            {
                cboClient.DataBindings.Clear();
                cboClient.DataSource = null;
                cboClient.Items.Clear();

                cboMatter.DataBindings.Clear();
                cboMatter.DataSource = null;
                cboMatter.Items.Clear();

                tvwFolder.Nodes.Clear();

                //cboEditor.DataBindings.Clear();
                //cboEditor.DataSource = null;
                //cboEditor.Items.Clear();

                cboContent.DataBindings.Clear();
                cboContent.DataSource = null;
                cboContent.Items.Clear();

                cboDocName.DataBindings.Clear();
                cboDocName.DataSource = null;
                cboDocName.Items.Clear();

                // Populate the Folders
                ObjMatter = new Matter();
                ObjClient = new Client();

                tvwFolder.Nodes.Clear();

                List<ProLoopFolder> folders = new List<ProLoopFolder>();
                if (context == Context.Orgs)
                {
                    folders = APIHelper.GetFolders("", ObjOrganization.Title, ObjMatter.Name, ObjClient.Name);
                }
                else
                {
                    folders = APIHelper.GetFolders(ObjProject.Title, "", "", "");
                }


                if (folders.Count > 0)
                {
                    foreach (ProLoopFolder current in folders)
                    {
                        this.tvwFolder.Nodes.Add(current.Name);
                    }
                }
                else
                {
                    string folderPath = string.Empty;

                    if (ObjProject != null && !string.IsNullOrEmpty(ObjProject.Title))
                    {
                        folderPath = "/api/files/Projects/" + ObjProject.Title + "/*?token=";
                    }
                    else
                    {

                        if (ObjMatter != null && !string.IsNullOrEmpty(ObjMatter.Name))
                        {
                            folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    ObjOrganization.Title,
                                    "/",
                                    ObjClient.Name,
                                    "/",
                                    ObjMatter.Name,
                                    "/*?token="
                            });
                        }
                        else
                        {
                            folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                     ObjOrganization.Title,
                                    "/",
                                   ObjClient.Name,
                                    "/*?token="
                            });
                        }
                    }
                    List<ProLoopFile> files = APIHelper.GetFiles(folderPath);
                    DocsBindingSource.DataSource = files;

                    cboDocName.DataSource = DocsBindingSource;
                    cboDocName.DisplayMember = "name";
                    cboDocName.ValueMember = "name";
                    //cboMatter_SelectedIndexChanged(cboMatter, null);

                    cboDocName.AutoCompleteMode = AutoCompleteMode.Suggest;
                    cboDocName.AutoCompleteSource = AutoCompleteSource.CustomSource;

                    DocsAutoCompleteCollection.AddRange(files.Select(file => file.Name).ToArray());
                    cboDocName.AutoCompleteCustomSource = DocsAutoCompleteCollection;

                    cboDocName.SelectedIndex = -1;
                    cboDocName.SelectedText = "Select:";
                }
            }
            Log.Debug("cboOrgProject_SelectedIndexChanged() -- End");

        }

        private void cboClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboClient_SelectedIndexChanged() -- Begin");
            ComboBox comboBox = (ComboBox)sender;
            bool flag = !comboBox.Focused;
            if (!flag)
            {
                if (this.cboClient.SelectedItem is Client)
                {
                    ObjClient = (Client)this.cboClient.SelectedItem;
                    searchParameter.ClientName = ObjClient.Name;
                }
                else
                {
                    ObjClient = null;
                }
                this.cboMatter.DataBindings.Clear();
                this.cboMatter.DataSource = null;
                this.cboMatter.Items.Clear();
                //this.cboEditor.DataBindings.Clear();
                //this.cboEditor.DataSource = null;
                //this.cboEditor.Items.Clear();
                this.cboContent.DataBindings.Clear();
                this.cboContent.DataSource = null;
                this.cboContent.Items.Clear();
                this.cboDocName.DataBindings.Clear();
                this.cboDocName.DataSource = null;
                this.cboDocName.Items.Clear();
                List<Matter> matters = APIHelper.GetMatters(ObjOrganization.Id, ObjClient.Id);
                MattersBindingSource.DataSource = matters;
                cboMatter.DataSource = MattersBindingSource;
                cboMatter.DisplayMember = "name";
                cboMatter.ValueMember = "id";
                //cboMatter_SelectedIndexChanged(cboMatter, null);

                cboMatter.AutoCompleteMode = AutoCompleteMode.Suggest;
                cboMatter.AutoCompleteSource = AutoCompleteSource.CustomSource;

                MattersAutoCompleteCollection.AddRange(matters.Select(matter => matter.Name).ToArray());
                cboMatter.AutoCompleteCustomSource = MattersAutoCompleteCollection;

                cboMatter.SelectedIndex = -1;
                cboMatter.SelectedText = "Select:";

                // Clear the selected text
                //cboOrgProject.SelectionStart = 0;

                Log.Debug("cboClient_SelectedIndexChanged() -- End");

            }
        }

        private void cboMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboMatter_SelectedIndexChanged() -- Begin");
            ComboBox comboBox = (ComboBox)sender;
            bool flag = !comboBox.Focused;
            if (!flag)
            {
                if (cboMatter.SelectedItem is Matter)
                {
                    ObjMatter = (Matter)this.cboMatter.SelectedItem;
                    searchParameter.MatterName = ObjMatter.Name;
                }
                else
                {
                    ObjMatter = new Matter();
                }
                this.tvwFolder.Nodes.Clear();
                //this.cboEditor.DataBindings.Clear();
                //this.cboEditor.DataSource = null;
                //this.cboEditor.Items.Clear();
                this.cboContent.DataBindings.Clear();
                this.cboContent.DataSource = null;
                this.cboContent.Items.Clear();
                this.cboDocName.DataBindings.Clear();
                this.cboDocName.DataSource = null;
                this.cboDocName.Items.Clear();
                List<ProLoopFolder> folders = new List<ProLoopFolder>();
                if (ObjProject == null)
                {
                    folders = APIHelper.GetFolders("", ObjOrganization.Title, ObjMatter.Name, ObjClient.Name);
                }
                else
                {
                    folders = APIHelper.GetFolders(ObjProject.Title, "", "", "");
                }
                this.tvwFolder.Nodes.Clear();
                foreach (ProLoopFolder current in folders)
                {
                    this.tvwFolder.Nodes.Add(current.Name);
                }
                Log.Debug("cboMatter_SelectedIndexChanged() -- End");
            }
        }

        private void cboDocName_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboDocName_SelectedIndexChanged() -- Begin");

            // Make sure the Combobox has the focus before processing the event handler
            ComboBox cbo = (ComboBox)sender;
            if (!cbo.Focused) return;
          
            var document = new ProLoopFile();
            if (cboDocName.SelectedItem is ProLoopFile)
            {
                document = (ProLoopFile)cboDocName.SelectedItem;
                DocumentName = document.Name;
                searchParameter.FileName = DocumentName;
            }
            else
            {
                // Unknown format
                DocumentName = string.Empty;
            }

            if (!string.IsNullOrEmpty(DocumentName))
            {
                btnOpen.Enabled = true;
            }

            Log.Debug("cboDocName_SelectedIndexChanged() -- End");
        }

        private void tvwFolder_AfterSelect(object sender, TreeViewEventArgs e)
        {
            // Get the full path of the node (from root node to the selected node)
            Stack<TreeNode> resultNodes = new Stack<TreeNode>();
            GetNodesToRoot(e.Node, resultNodes);
            searchParameter.FolderName = e.Node.Text;
            StringBuilder nodePath = new StringBuilder();
            while (resultNodes.Count > 0)
            {
                nodePath.Append(resultNodes.Pop().Text);
                if (resultNodes.Count > 1) nodePath.Append("/");
            }
            this.FolderPath = nodePath.ToString();

            // Get the files in this path
            string folderPath = string.Empty;

            if (ObjProject != null && !string.IsNullOrEmpty(ObjProject.Title))
            {
                folderPath = string.Concat(new object[]
                {
                    "/api/files/Projects/",
                    ObjProject.Title,
                    "/",
                    nodePath,
                    "/*?token="
                });
            }
            else
            {

                if (ObjMatter != null && !string.IsNullOrEmpty(ObjMatter.Name))
                {

                    folderPath = "/api/files/Organizations/" + ObjOrganization.Title + "/" + ObjClient.Name + "/" + ObjMatter.Name + "/" + nodePath + "/*?token=";
                   
                }
                else
                {
                    folderPath = "/api/files/Organizations/" + ObjOrganization.Title + "/" + ObjClient.Name + "/" + nodePath + "/*?token=";
                }
            }
            List<ProLoopFile> files = APIHelper.GetFiles(folderPath);           

            DocsBindingSource.DataSource = files;

            cboDocName.DataSource = DocsBindingSource;
            cboDocName.DisplayMember = "name";
            cboDocName.ValueMember = "name";
            //cboMatter_SelectedIndexChanged(cboMatter, null);

            cboDocName.AutoCompleteMode = AutoCompleteMode.Suggest;
            cboDocName.AutoCompleteSource = AutoCompleteSource.CustomSource;

            DocsAutoCompleteCollection.AddRange(files.Select(file => file.Name).ToArray());
            cboDocName.AutoCompleteCustomSource = DocsAutoCompleteCollection;

            cboDocName.SelectedIndex = -1;
            cboDocName.SelectedText = "Select:";
        }
        private void GetNodesToRoot(TreeNode node, Stack<TreeNode> resultNodes)
        {
            if (node == null)
                return; // previous node was the root.           
                resultNodes.Push(node);
                GetNodesToRoot(node.Parent, resultNodes);
            
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            searchParameter = new SearchParameter();
            Log.Debug("btnOpen_Click() -- Begin");

            if (string.IsNullOrEmpty(AddinCurrentInstance.WebDAVMappedDriveLetter))
                AddinCurrentInstance.MapWebDAVFolderToDriveLetter();
            string filePath = AddinCurrentInstance.WebDAVMappedDriveLetter;
            if (!string.IsNullOrEmpty(selectedPath))
            {
                filePath =filePath+"\\"+ selectedPath;
                selectedPath = string.Empty;
            }
            else
            {
                if (!string.IsNullOrEmpty(DocumentName)) // Open the document
                {

                    //string filePath = AddinCurrentInstance.ProLoopUrl;
                    if (ObjProject != null && !string.IsNullOrEmpty(ObjProject.Title))
                    {
                        if (!string.IsNullOrEmpty(FolderPath))
                            filePath += "\\Projects\\" + ObjProject.Title + "\\" + FolderPath + "\\" + DocumentName; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        else
                            filePath += "\\Projects\\" + ObjProject.Title + "\\" + DocumentName;
                    }
                    else
                    {
                        if (ObjMatter != null && !string.IsNullOrEmpty(ObjMatter.Name))
                        {
                            if (!string.IsNullOrEmpty(FolderPath))
                                filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                            FolderPath + "\\" + DocumentName;
                            else
                                filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                            DocumentName; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(FolderPath))
                                filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + FolderPath + "\\" +
                                            DocumentName; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                            else
                                filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + DocumentName;
                            // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        }
                    }
                }
            }
            _Application WordApp = AddinCurrentInstance.WordApp;
            WordApp.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            Documents documents = null;
            Document document = null;
            try
            {
                documents = WordApp.Documents;

                // If a document is already open in the active Window, close it
                if (documents.Count > 0)
                {
                    document = WordApp.ActiveDocument;
                    document.Close();
                }

                // Open the selected document
                document = documents.Open(filePath, AddToRecentFiles: false);
            }
            catch (Exception exception)
            {
                Log.Error("Error while opening the selected document file using File->Open.");
                Log.Error(exception.Message);
            }

            //if (documents != null) Marshal.ReleaseComObject(documents);

            Log.Debug("btnOpen_Click() -- End");
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            Log.Debug("btnSettings_Click() -- Begin");

            ShowSettingsForm();

#if NETFX45
            Task.Factory.StartNew(AddinCurrentInstance.UnmapAndMapWebDavFolderToDrive);
#elif NETFX35
            ThreadPool.QueueUserWorkItem(obj => { AddinCurrentInstance.UnmapAndMapWebDavFolderToDrive(); });
#endif

            Log.Debug("btnSettings_Click() -- End");
        }
        private void ShowSettingsForm()
        {
            Log.Debug("ShowSettingsForm() -- Begin");

            if (settingForm == null)
                settingForm = new SettingsForm();

            // Read the settings (URL, Username and Password) from the registry & display the Settings form
            AddinCurrentInstance.LoadValuesFromRegistry();

            if (settingForm.ShowDialog() == DialogResult.OK)
            {
                var result = AddinCurrentInstance.SaveValuesToRegistry(AddinCurrentInstance.ProLoopUrl,
                    AddinCurrentInstance.ProLoopUsername, AddinCurrentInstance.ProLoopPassword);

                if (result)
                    MessageBox.Show("Successfully saved the values.", "Success", MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                else
                    MessageBox.Show("Error while saving the values.", "Error", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
            }

            Log.Debug("ShowSettingsForm() -- End");
        }

        private void buttonSearch_Click(object sender, EventArgs e)
        {
            searchParameter.KeyWord = textBoxContent.Text;
            using (var search = new AheadSearchForm(searchParameter))
            {
                var dialogResult = search.ShowDialog();
                if(dialogResult==DialogResult.OK)
                {
                    selectedPath =search.FilePath;
                    if(!string.IsNullOrEmpty(selectedPath))
                    {
                        
                        FillFromSearchWindow();
                        btnOpen.Enabled = true;
                    }
                }
            }
        }
        private void FillFromSearchWindow()
        {
            string[] PathCollection = selectedPath.Split('/');
            if(selectedPath.StartsWith("Organizations"))
            {
                cboOrgProject.Text = string.Empty;
                rbOrganizations.Checked = true;
                cboOrgProject.SelectedText = PathCollection[1];
                cboClient.Enabled = true;
                cboMatter.Enabled = true;

                //Handle Client
                cboClient.Text = string.Empty;
                cboClient.SelectedText = PathCollection[2];

                //Handle Matter
                cboMatter.Text = string.Empty;
                cboMatter.SelectedText = PathCollection[3];

                // Handle FolderName
                tvwFolder.Nodes.Clear();
                for (int i=4;i<PathCollection.Length-1;i++)
                {
                    this.tvwFolder.Nodes.Add(PathCollection[i]);
                }
                //Handle FileName
                cboDocName.Text = string.Empty;
                cboDocName.SelectedText = PathCollection[PathCollection.Length - 1];
            }
            else
            {
                rbProjects.Checked = true;
                cboOrgProject.Text = string.Empty;
                cboOrgProject.SelectedText = PathCollection[1];
                cboClient.Enabled = false;
                cboMatter.Enabled = false;
                // Handle FolderName
                tvwFolder.Nodes.Clear();
                for (int i = 2; i < PathCollection.Length - 1; i++)
                {
                    this.tvwFolder.Nodes.Add(PathCollection[i]);
                }
                cboDocName.Text = string.Empty;
                cboDocName.SelectedText = PathCollection[PathCollection.Length-1];
            }
            //cboContent.Text = searchParameter.KeyWord;
            textBoxContent.Text = searchParameter.KeyWord;
            cboEditor.Text = searchParameter.EditorName;
            
        }

        private void cboEditor_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchParameter.EditorName = cboEditor.Text;
        }
        //private void ADXWordOpenTaskPane_Load(object sender, EventArgs e)
        //{
        //    AddinCurrentInstance.Mode = Operation.Open;
        //}

        //private void ADXWordOpenTaskPane_Activated(object sender, EventArgs e)
        //{

        //}
    }
}