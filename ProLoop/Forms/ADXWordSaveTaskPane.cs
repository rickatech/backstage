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
using System.Drawing;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordSaveTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;       
        private SettingsForm settingForm;

        private Client ObjClient { get; set; } = new Client();
        private Project ObjProject { get; set; } = new Project();
        private Organization ObjOrganization { get; set; } = new Organization();
        private Matter ObjMatter { get; set; } = new Matter();

        private string FolderPath = string.Empty;
        
        public ADXWordSaveTaskPane()
        {
            InitializeComponent();
            btnSaveAs.Enabled = false;
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            rbOrganizations.Checked = Properties.Settings.Default.SaveOrganization;
            rbProjects.Checked = Properties.Settings.Default.SaveProject;
           
        }

        private void ADXWordSaveTaskPane_Load(object sender, EventArgs e)
        {
            AddinCurrentInstance.Mode = Operation.Open;
        }

        private void ADXWordSaveTaskPane_Activated(object sender, EventArgs e)
        {
            txtboxFileName.Text = AddinCurrentInstance.WordApp.ActiveDocument.Name;
        }

        private void rbOrganizations_CheckedChanged(object sender, EventArgs e)
        {
            Log.Debug("rbOrganizations_CheckedChanged() -- Begin");
            FolderPath = string.Empty;
            // Disable the Open button
            btnSaveAs.Enabled = true;
            cboClient.Text = string.Empty;
            cboMatter.Text = string.Empty;
            RadioButton rbOrgs = sender as RadioButton;
            CheckProLoopSettings(rbOrgs);
            if (rbOrganizations.Checked)
            {
                // Change the label to Organizations
                lblOrgsProjects.Text = "Organizations";               
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

                List<Organization> orgs = APIHelper.GetOrganizations();
                cboOrgProject.DataSource = orgs;
                cboOrgProject.MatchingMethod = StringMatchingMethod.NoWildcards;
                cboOrgProject.DisplayMember = "title";
                cboOrgProject.ValueMember = "id";
                if (!string.IsNullOrEmpty(Properties.Settings.Default.SaveEntityName))
                {
                    var item = orgs.FirstOrDefault(x => x.Title == Properties.Settings.Default.SaveEntityName);
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
            btnSaveAs.Enabled = true;
            FolderPath = string.Empty;
            cboClient.Text = string.Empty;
            cboMatter.Text = string.Empty;
            // Disable the Open button
           

            RadioButton rbProjs = sender as RadioButton;
            CheckProLoopSettings(rbProjs);

            // Projects checkbox selected
            if (rbProjects.Checked)
            {
                // Change the label to Organizations
                lblOrgsProjects.Text = "Projects";               
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

                if (!string.IsNullOrEmpty(Properties.Settings.Default.SaveEntityName) && projects != null)
                {
                    var item = projects.FirstOrDefault(x => x.Title == Properties.Settings.Default.SaveEntityName);
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
            
        }

        private void cboOrgProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            Log.Debug("cboOrgProject_SelectedIndexChanged() -- Begin");
            FolderPath = string.Empty;
            // Make sure the Combo box has the focus before processing the event handler.
            // This will fire the event only when the Combo box has focus
            var cbo = (EasyCompletionComboBox)sender;
            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.SaveClient) && !cbo.Focused)
            {
                ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                cboClient.Items.Clear();
                var client = new Client() { Name = Properties.Settings.Default.SaveClient, Id = Properties.Settings.Default.SaveClientId };
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
                
                ObjProject = null;
            }
            else if (cboOrgProject.SelectedItem is Project)
            {
                ObjProject = cboOrgProject.SelectedItem as Project;
                ObjOrganization = null;
               
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
                    ProcessFolderItem(folders,tempFolder);
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

                   
                }
            }
            Log.Debug("cboOrgProject_SelectedIndexChanged() -- End");
        }

        private void cboClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            FolderPath = string.Empty;

            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.SaveMatter) && !cboClient.Focused)
            {
                if (cboMatter.Items.Count == 0)
                {
                    ObjOrganization = (Organization)this.cboOrgProject.SelectedItem;
                    ObjClient = (Client)this.cboClient.SelectedItem;
                    cboMatter.Items.Clear();
                    var matter = new Matter() { Name = Properties.Settings.Default.SaveMatter, Id = Properties.Settings.Default.SaveMatterId };
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

            }
            else
            {
                ObjClient = null;
            }
            //this.cboMatter.DataBindings.Clear();
            this.cboMatter.DataSource = null;
            this.cboMatter.Items.Clear();

            if (ObjClient == null)
                return;
            List<Matter> matters = APIHelper.GetMatters(ObjOrganization.Id, ObjClient.Id);
            cboMatter.DataSource = matters;
            cboMatter.MatchingMethod = StringMatchingMethod.NoWildcards;
            cboMatter.DisplayMember = "name";
            cboMatter.ValueMember = "id";

            if (rbOrganizations.Checked && !string.IsNullOrEmpty(Properties.Settings.Default.SaveMatter))
            {
                var item = matters.FirstOrDefault(x => x.Name == Properties.Settings.Default.SaveMatter);
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
            FolderPath = string.Empty;
            ComboBox comboBox = (ComboBox)sender;
            bool flag = !comboBox.Focused;
            if (!flag)
            {
                if (cboMatter.SelectedItem is Matter)
                {
                    ObjMatter = (Matter)this.cboMatter.SelectedItem;
                   
                }
                else
                {
                    ObjMatter = new Matter();
                }
                this.tvwFolder.Nodes.Clear();

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

        private void ProcessFolderItem(List<ProLoopFolder> folders,string tempfolder)
        {
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

        private void btnSaveAs_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtboxFileName.Text))
            {
                MessageBox.Show("File Name is required", "Proloop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string filePath = AddinCurrentInstance.WebDAVMappedDriveLetter;
            if (ObjProject != null && !string.IsNullOrEmpty(ObjProject.Title))
            {
                if (!string.IsNullOrEmpty(FolderPath))
                    filePath += "\\Projects\\" + ObjProject.Title + "\\" + FolderPath + "\\" + txtboxFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                else
                    filePath += "\\Projects\\" + ObjProject.Title + "\\" + txtboxFileName.Text;
            }
            else
            {
                if (ObjMatter != null && !string.IsNullOrEmpty(ObjMatter.Name))
                {
                    if (!string.IsNullOrEmpty(FolderPath))
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                    FolderPath + "\\" + txtboxFileName.Text;
                    else
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                    txtboxFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                }
                else
                {
                    if (!string.IsNullOrEmpty(FolderPath))
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + FolderPath + "\\" +
                                    txtboxFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                    else
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + txtboxFileName.Text;
                    // + "?token=" + AddinCurrentInstance.ProLoopToken;
                }
            }

            if (string.IsNullOrEmpty(System.IO.Path.GetExtension(filePath)))
            {
                filePath += ".docx";
            }
            try
            {
                var missing = System.Type.Missing;
                var AddtoRecent = false;
                if (AddinCurrentInstance.SelectedFile != null)
                {
                    if(!string.IsNullOrWhiteSpace(AddinCurrentInstance.SelectedFile.LockingUserId) && AddinCurrentInstance.SelectedFile.LockingUserId != AddinCurrentInstance.CurrentUserId.ToString())
                    {
                        var message = "This Document is Checked Out by Rick Shaul.Would you like to Save As a new Document?";
                        var resposne = AddinCurrentInstance.DisplayWaranigMessage(message);
                        if (resposne)
                        {                            
                            AddinCurrentInstance.WordApp.ActiveDocument.SaveAs(filePath, missing, missing, missing, AddtoRecent, missing,
                                missing, missing, missing, missing, missing, missing, missing, missing, missing);
                            SaveSettingChange();
                        }
                        else
                        {
                            return;
                        }
                    }
                    else
                    {
                        AddinCurrentInstance.WordApp.ActiveDocument.SaveAs(filePath, missing, missing, missing, AddtoRecent, missing,
                                            missing, missing, missing, missing, missing, missing, missing, missing, missing);
                        SaveSettingChange();
                    }
                }
                else {
                    AddinCurrentInstance.WordApp.ActiveDocument.SaveAs(filePath, missing, missing, missing, AddtoRecent, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    SaveSettingChange();
                }
               
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save document", "Proloop", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }


        private void tvwFolder_AfterSelect(object sender, TreeViewEventArgs e)
        {
            StringBuilder nodePath = new StringBuilder();
            e.Node.EnsureVisible();
            
            if (!e.Node.Text.Contains(@"\"))
            {

                string[] data = e.Node.FullPath.Split('\\');
                foreach (string item in data)
                {
                    if (string.IsNullOrEmpty(item))
                        continue;
                    nodePath.Append(item).Append("\\");
                }
                this.FolderPath = nodePath.ToString().TrimEnd('\\');               
            }
            // Get the files in this path
            string folderPath = string.Empty;
        }

        private void GetNodesToRoot(TreeNode node, Stack<TreeNode> resultNodes)
        {
            if (node == null)
                return; // previous node was the root.           
            resultNodes.Push(node);
            GetNodesToRoot(node.Parent, resultNodes);

        }

        private void SaveSettingChange()
        {
            Properties.Settings.Default.SaveOrganization = rbOrganizations.Checked;
            Properties.Settings.Default.SaveProject = rbProjects.Checked;
            if (rbOrganizations.Checked)
            {
                Properties.Settings.Default.SaveMatter = ((Matter)cboMatter.SelectedItem).Name;
                Properties.Settings.Default.SaveMatterId = ((Matter)cboMatter.SelectedItem).Id;
                Properties.Settings.Default.SaveClient = ((Client)cboClient.SelectedItem).Name;
                Properties.Settings.Default.SaveClientId = ((Client)cboClient.SelectedItem).Id;
            }
            Properties.Settings.Default.SaveEntityName = cboOrgProject.Text;
            Properties.Settings.Default.Save();
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
    }
}