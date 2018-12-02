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
using ProLoop.WordAddin.Model;
using System.IO;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordOpenTaskPane : AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;
        private SettingsForm settingForm;

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
        ProLoopFile proloopFile { get; set; }
        private SearchParameter searchParameter;

        public ADXWordOpenTaskPane()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            searchParameter = new SearchParameter();
            rbOrganizations.Checked = Properties.Settings.Default.OpenOrganisation;
            rbProjects.Checked = Properties.Settings.Default.OpenProject;
            if (string.IsNullOrEmpty(AddinCurrentInstance.ProLoopUsername))
            {
                label7.Visible = false;
            }
            else
            {
                label7.Visible = true;
                label7.Text = $"Current user : {AddinCurrentInstance.ProLoopUsername}";
            }
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
                searchParameter.ClientName = string.Empty;
                searchParameter.OrgOrProjectName = string.Empty;
                searchParameter.MatterName = string.Empty;
                searchParameter.EditorName = string.Empty;
                searchParameter.FileName = string.Empty;
                searchParameter.KeyWord = string.Empty;
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
                cboOrgProject.DataSource = orgs;
                cboOrgProject.MatchingMethod = StringMatchingMethod.NoWildcards;
                cboOrgProject.DisplayMember = "title";
                cboOrgProject.ValueMember = "id";
                if (!string.IsNullOrEmpty(Properties.Settings.Default.OpenEntity))
                {
                    var item = orgs.FirstOrDefault(x => x.Title == Properties.Settings.Default.OpenEntity);
                    if (item != null)
                    {
                        cboOrgProject.SelectedItem = item;

                    }
                    else
                    {
                        cboOrgProject.SelectedIndex = -1;
                        cboOrgProject.Text = "Select:";
                    }
                }
                else
                {
                    cboOrgProject.SelectedIndex = -1;
                    cboOrgProject.Text = "Select:";
                }
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
                searchParameter.OrgOrProjectName = string.Empty;
                searchParameter.ClientName = string.Empty;
                searchParameter.MatterName = string.Empty;
                searchParameter.EditorName = string.Empty;
                searchParameter.FileName = string.Empty;
                searchParameter.KeyWord = string.Empty;
                // Disable Client and Matter
                cboClient.Enabled = false;
                ObjClient = null;
                ObjMatter = null;
                cboMatter.Enabled = false;

                // Clear the Client, Matter, Docs..


                tvwFolder.Nodes.Clear();

                List<Project> projects = APIHelper.GetProjects();

                //ProjectsBindingSource.DataSource = projects;
                cboOrgProject.DataSource = projects;
                cboOrgProject.MatchingMethod = StringMatchingMethod.NoWildcards;
                cboOrgProject.DisplayMember = "title";
                cboOrgProject.ValueMember = "id";

                if (!string.IsNullOrEmpty(Properties.Settings.Default.OpenEntity) && projects != null)
                {
                    var item = projects.FirstOrDefault(x => x.Title == Properties.Settings.Default.OpenEntity);
                    if (item != null)
                    {
                        cboOrgProject.SelectedItem = item;
                    }
                    else
                    {
                        cboOrgProject.SelectedIndex = -1;
                        cboOrgProject.Text = "Select:";
                    }
                }
                else
                {
                    cboOrgProject.SelectedIndex = -1;
                    cboOrgProject.Text = "Select:";
                }


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
            var cbo = (EasyCompletionComboBox)sender;
            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.OpenClient) && !cbo.Focused)
            {
                ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                cboClient.Items.Clear();
                var client = new Client() { Name = Properties.Settings.Default.OpenClient, Id = Properties.Settings.Default.OpenClientId };
                cboClient.Items.Add(client);
                cboClient.DisplayMember = "name";
                cboClient.ValueMember = "id";
                cboClient.SelectedItem = client;
                ObjProject = null;
            }
            if (!cbo.Focused) return;

            Context context = rbOrganizations.Checked ? Context.Orgs : Context.Projects;

            if (cboOrgProject.SelectedItem is Organization)
            {
                ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                searchParameter.OrgOrProjectName = ObjOrganization.Title;
                searchParameter.ClientName = string.Empty;
                searchParameter.MatterName = string.Empty;
                searchParameter.EditorName = string.Empty;
                searchParameter.FileName = string.Empty;
                searchParameter.KeyWord = string.Empty;
                ObjProject = null;
            }
            else if (cboOrgProject.SelectedItem is Project)
            {
                ObjProject = cboOrgProject.SelectedItem as Project;
                ObjOrganization = null;
                searchParameter.OrgOrProjectName = ObjProject.Title;
                searchParameter.ClientName = string.Empty;
                searchParameter.MatterName = string.Empty;
                searchParameter.EditorName = string.Empty;
                searchParameter.FileName = string.Empty;
                searchParameter.KeyWord = string.Empty;
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
                //ClientsBindingSource.DataSource = clients;
                cboClient.DataSource = clients;
                cboClient.MatchingMethod = StringMatchingMethod.NoWildcards;
                cboClient.DisplayMember = "name";
                cboClient.ValueMember = "id";
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
                string tempFolder = string.Empty;
                List<ProLoopFolder> folders = new List<ProLoopFolder>();
                if (context == Context.Orgs)
                {
                    tempFolder = APIHelper.GetFolderPath("", ObjOrganization.Title, ObjMatter.Name, ObjClient.Name);
                    folders = APIHelper.GetFolders(tempFolder);
                }
                else
                {
                    tempFolder = APIHelper.GetFolderPath(ObjProject.Title, "", "", "");
                    folders = APIHelper.GetFolders(tempFolder);
                }


                if (folders.Count > 0)
                {
                    //foreach (ProLoopFolder current in folders)
                    //{
                    //    this.tvwFolder.Nodes.Add(current.Name);
                    //}
                    ProcessFolderItem(folders, tempFolder);
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
                    // DocsBindingSource.DataSource = files;

                    cboDocName.DataSource = files;
                    cboDocName.MatchingMethod = StringMatchingMethod.NoWildcards;
                    cboDocName.DisplayMember = "name";
                    cboDocName.ValueMember = "name";

                    cboDocName.SelectedIndex = -1;
                    cboDocName.SelectedText = "Select:";
                }
            }
            Log.Debug("cboOrgProject_SelectedIndexChanged() -- End");

        }

        private void cboClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.OpenMatter) && !cboClient.Focused)
            {
                if (cboMatter.Items.Count == 0)
                {
                    ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                    ObjClient = (Client)this.cboClient.SelectedItem;
                    cboMatter.Items.Clear();
                    var matter = new Matter() { Name = Properties.Settings.Default.OpenMatter, Id = Properties.Settings.Default.OpenMatterId };
                    cboMatter.Items.Add(matter);
                    cboMatter.DisplayMember = "name";
                    cboMatter.ValueMember = "id";
                    cboMatter.SelectedItem = matter;
                }
            }

            if (!cboClient.Focused)
                return;
            Log.Debug("cboClient_SelectedIndexChanged() -- Begin");

            if (this.cboClient.SelectedItem is Client)
            {
                ObjClient = (Client)this.cboClient.SelectedItem;
                searchParameter.ClientName = ObjClient.Name;
                searchParameter.MatterName = string.Empty;
                searchParameter.EditorName = string.Empty;
                searchParameter.FileName = string.Empty;
                searchParameter.KeyWord = string.Empty;
            }
            else
            {
                ObjClient = null;
            }
            //this.cboMatter.DataBindings.Clear();
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
            if (ObjClient == null)
                return;
            List<Matter> matters = APIHelper.GetMatters(ObjOrganization.Id, ObjClient.Id);
            cboMatter.DataSource = matters;
            cboMatter.MatchingMethod = StringMatchingMethod.NoWildcards;
            cboMatter.DisplayMember = "name";
            cboMatter.ValueMember = "id";
            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.OpenMatter))
            {
                var item = matters.FirstOrDefault(x => x.Name == Properties.Settings.Default.OpenMatter);
                if (item != null)
                    cboMatter.SelectedItem = item;
                else
                {
                    cboMatter.SelectedIndex = -1;
                    cboMatter.SelectedText = "Select:";
                }
            }
            else
            {
                cboMatter.SelectedIndex = -1;
                cboMatter.SelectedText = "Select:";
            }

            Log.Debug("cboClient_SelectedIndexChanged() -- End");

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
                    searchParameter.EditorName = string.Empty;
                    searchParameter.FileName = string.Empty;
                    searchParameter.KeyWord = string.Empty;
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
                string tempFolder = string.Empty;
                if (ObjProject == null)
                {
                    tempFolder = APIHelper.GetFolderPath("", ObjOrganization.Title, ObjMatter.Name, ObjClient.Name);
                    folders = APIHelper.GetFolders(tempFolder);
                }
                else
                {
                    tempFolder = APIHelper.GetFolderPath(ObjProject.Title, "", "", "");
                    folders = APIHelper.GetFolders(tempFolder);
                }

                ProcessFolderItem(folders, tempFolder);

                Log.Debug("cboMatter_SelectedIndexChanged() -- End");
            }
        }

        private void ProcessFolderItem(List<ProLoopFolder> folders, string tempfolder)
        {
            var filesitem = folders.Where(x => x.Name.Contains(".")).ToList();
            if (filesitem != null && filesitem.Count > 0)
            {
                List<ProLoopFile> files = new List<ProLoopFile>();
                foreach (var item in filesitem)
                {
                    files.Add(new ProLoopFile() { Name = item.Name, Path = item.Path });
                }
                if (files.Count > 0)
                // DocsBindingSource.DataSource = files;
                {
                    cboDocName.DataSource = files;
                    cboDocName.DisplayMember = "name";
                    cboDocName.ValueMember = "name";
                    cboDocName.SelectedIndex = -1;
                    cboDocName.SelectedText = "Select:";
                }
            }
            //Pull only folder item
            folders = folders.Where(x => !x.Name.Contains(".")).ToList();
            if (folders != null && folders.Count > 0)
            {
                this.tvwFolder.Nodes.Clear();
                //Adding first Node               
                var root = tvwFolder.Nodes.Add(@"\");
                root.Tag = tempfolder;
                foreach (ProLoopFolder current in folders)
                {
                    var node = new TreeNode();
                    node.Text = current.Name;
                    node.Tag = current.Path;
                    root.Nodes.Add(node);
                }
            }
        }

        private void cboDocName_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboDocName_SelectedIndexChanged() -- Begin");
            // Make sure the Combobox has the focus before processing the event handler
            ComboBox cbo = (ComboBox)sender;
            if (!cbo.Focused) return;

            proloopFile = new ProLoopFile();
            if (cboDocName.SelectedItem is ProLoopFile)
            {
                proloopFile = (ProLoopFile)cboDocName.SelectedItem;
                DocumentName = proloopFile.Name;
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

            StringBuilder nodePath = new StringBuilder();
            if (!e.Node.Text.Contains(@"\"))
            {

                string[] data = e.Node.FullPath.Split('\\');
                foreach (string item in data)
                {
                    if (string.IsNullOrEmpty(item))
                        continue;
                    nodePath.Append(item).Append("/");
                }
                this.FolderPath = nodePath.ToString().TrimEnd('/');
                searchParameter.FolderName = FolderPath;
            }
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

                    folderPath = "/api/files/Organizations/" + ObjOrganization.Title + "/" + ObjClient.Name + "/" + ObjMatter.Name + "/" + nodePath + "*?token=";

                }
                else
                {
                    folderPath = "/api/files/Organizations/" + ObjOrganization.Title + "/" + ObjClient.Name + "/" + nodePath + "*?token=";
                }
            }
            List<ProLoopFile> files = APIHelper.GetFiles(folderPath);

            // DocsBindingSource.DataSource = files;

            if (files.Count > 0)
            {
                cboDocName.DataSource = files.Where(x => x.Name.Contains(".")).ToList();
            }
            cboDocName.DisplayMember = "name";
            cboDocName.ValueMember = "name";

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
                filePath = filePath + "\\" + selectedPath;
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
                document = documents.Open(filePath, AddToRecentFiles: false,ReadOnly:false);
                string fileInfo = $"{AddinCurrentInstance.ProLoopUrl}/api/filetags/{proloopFile.Path}";
                var data = MetadataHandler.GetMetaDataInfo<MetaDataInfo>(fileInfo);
                if (data != null)
                {
                    foreach (Section wordSection in document.Sections)
                    {
                        Range footerRange = wordSection.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;

                        footerRange.Font.ColorIndex = WdColorIndex.wdBlack;
                        footerRange.Bold = 1;
                        footerRange.Text = $"Doc Id:{data[0].VersionId} \t\t Version:2.0.0";
                    }
                    FileInfo docinfo = new FileInfo(filePath);
                    if (!docinfo.IsReadOnly)
                    {
                        document.Save();
                    }                 
                }
                SaveSettingChange();
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
            searchParameter.EditorName = textBoxEditor.Text;
            using (var search = new AheadSearchForm(searchParameter))
            {
                var dialogResult = search.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    selectedPath = search.FilePath;
                    if (!string.IsNullOrEmpty(selectedPath))
                    {

                        FillFromSearchWindow();
                        btnOpen.Enabled = true;
                    }
                }
            }
        }

        private void SaveSettingChange()
        {
            Properties.Settings.Default.OpenOrganisation = rbOrganizations.Checked;
            Properties.Settings.Default.OpenProject = rbProjects.Checked;
            if (rbOrganizations.Checked)
            {
                Properties.Settings.Default.OpenMatter = ((Matter)cboMatter.SelectedItem).Name;
                Properties.Settings.Default.OpenMatterId = ((Matter)cboMatter.SelectedItem).Id;
                Properties.Settings.Default.OpenClient = ((Client)cboClient.SelectedItem).Name;
                Properties.Settings.Default.OpenClientId = ((Client)cboClient.SelectedItem).Id;
            }
            Properties.Settings.Default.OpenEntity = cboOrgProject.Text;
            Properties.Settings.Default.Save();
        }
        private void FillFromSearchWindow()
        {
            string[] PathCollection = selectedPath.Split('/');
            if (selectedPath.StartsWith("Organizations"))
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
                for (int i = 4; i < PathCollection.Length - 1; i++)
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
                var root = tvwFolder.Nodes.Add(@"\");
                for (int i = 2; i < PathCollection.Length - 1; i++)
                {
                    root.Nodes.Add(PathCollection[i]);
                }
                cboDocName.Text = string.Empty;
                cboDocName.SelectedText = PathCollection[PathCollection.Length - 1];
            }
            //cboContent.Text = searchParameter.KeyWord;
            textBoxContent.Text = searchParameter.KeyWord;
            cboEditor.Text = searchParameter.EditorName;

        }

        private void cboEditor_SelectedIndexChanged(object sender, EventArgs e)
        {
            searchParameter.EditorName = cboEditor.Text;
        }

        private void tvwFolder_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            var expandNode = e.Node;
            if (expandNode != null)
            {
                expandNode.Nodes.Clear();
                var folders = APIHelper.GetFolders(expandNode.Tag as string);
                foreach (ProLoopFolder folder in folders)
                {
                    if (folder.Name.Contains("."))
                        continue;
                    TreeNode childNode = new TreeNode();
                    childNode.Text = folder.Name;
                    childNode.Tag = $"/api/files/{folder.Path}";
                    expandNode.Nodes.Add(childNode);
                }
                foreach (TreeNode node in expandNode.Nodes)
                {
                    if (node != null && node.Tag is string)
                    {
                        string folderPath = node.Tag as string;
                        if (!string.IsNullOrEmpty(folderPath))
                        {
                            folderPath = node.Tag as string;
                            var childFolders = APIHelper.GetFolders(folderPath);
                            if (childFolders != null)
                            {
                                foreach (ProLoopFolder childFolder in childFolders)
                                {
                                    if (childFolder.Name.Contains("."))
                                        continue;
                                    var treenode = new TreeNode();
                                    treenode.Text = childFolder.Name;
                                    treenode.Tag = childFolder.Path;
                                    node.Nodes.Add(treenode);
                                }
                            }
                        }
                    }
                }
            }
        }

        private void buttonFullDocSearch_Click(object sender, EventArgs e)
        {
            using (var search = new ContentSearch(searchParameter))
            {
                var dialogResult = search.ShowDialog();
                if (dialogResult == DialogResult.OK)
                {
                    selectedPath = search.FilePath;
                    if (!string.IsNullOrEmpty(selectedPath))
                    {
                        FillFromSearchWindow();
                        btnOpen.Enabled = true;
                        btnOpen.PerformClick();
                    }
                }
            }
        }
    }
}