using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using System.Collections.Generic;
using AddinExpress.MSO;
using ProLoop.WordAddin.Service;
using ProLoop.WordAddin.Model;
using System.Linq;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordTaskPaneSaveNew: AddinExpress.WD.ADXWordTaskPane
    {
       
        private readonly AddinModule AddinCurrentInstance;
        private string orgProject, client, matter;
        List<Organization> orgs = new List<Organization>();
        List<Project> projects = new List<Project>();

        private Client ObjClient { get; set; } = new Client();
        private Project ObjProject { get; set; } = new Project();
        private Organization ObjOrganization { get; set; } = new Organization();
        private Matter ObjMatter { get; set; } = new Matter();

        private string FolderPath = string.Empty;
        private bool isActionbyTreeView = false;
        public ADXWordTaskPaneSaveNew()
        {
            InitializeComponent();
            this.AddinCurrentInstance = ADXAddinModule.CurrentInstance as AddinModule;
            this.Load += ADXWordOpenNewTaskPane_Load;
        }

        private void treeView1_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag is null)
                return;
            else if (e.Node.Tag != null && e.Node.Tag is string)
            {
                if ((string)e.Node.Tag == "Organization")
                {
                    Cursor.Current = Cursors.WaitCursor;
                    // Application.DoEvents();
                    foreach (TreeNode item in e.Node.Nodes)
                    {
                        if (item.Tag == null)
                            continue;
                        var org = item.Tag as Organization;
                        var clients = APIHelper.GetClients(org.Id);
                        if (clients == null || clients.Count == 0)
                            continue;
                        AddTreeNodeToItem(item, clients);
                    }
                    Cursor.Current = Cursors.Default;
                }
                else if ((string)e.Node.Tag == "project")
                {
                    Cursor.Current = Cursors.WaitCursor;
                    foreach (TreeNode item in e.Node.Nodes)
                    {
                        var tempFolder = APIHelper.GetFolderPath(item.Text, "", "", "");
                        var folders = APIHelper.GetFolders(tempFolder);
                        if (folders != null)
                        {
                            AddTreeNodeToItem(item, folders);
                        }
                    }
                    Cursor.Current = Cursors.Default;
                }
            }
            else if (e.Node.Tag is Organization)
            {
                if (e.Node.Tag == null)
                    return;
                Cursor.Current = Cursors.WaitCursor;
                var org = e.Node.Tag as Organization;
                foreach (TreeNode item in e.Node.Nodes)
                {
                    if (item.Tag == null)
                        continue;
                    var client = item.Tag as Client;
                    var matters = APIHelper.GetMatters(org.Id, client.Id);
                    if (matters == null || matters.Count == 0)
                        continue;
                    AddTreeNodeToItem(item, matters);
                }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Node.Tag is Project)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (TreeNode item in e.Node.Nodes)
                {
                    var tempFolder = APIHelper.GetFolderPath(item.Text, "", "", "");
                    var folders = APIHelper.GetFolders(tempFolder);
                    if (folders != null)
                        AddTreeNodeToItem(item, folders);
                }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Node.Tag is Matter)
            {
                var currentNode = e.Node;
                if (currentNode != null)
                {
                    foreach (TreeNode node in currentNode.Nodes)
                    {
                        var folderItem = node.Tag as ProLoopFolder;
                        var folders = APIHelper.GetFolders($"/api/files/{folderItem.Path}");
                        if (folders != null)
                            AddTreeNodeToItem(node, folders);
                    }
                    //currentNode.Nodes.Clear();

                }
            }
            else if (e.Node.Tag is Client)
            {
                Cursor.Current = Cursors.WaitCursor;
                foreach (TreeNode node in e.Node.Nodes)
                {
                    var tempFolder = APIHelper.GetFolderPath("", e.Node.Parent.Text, node.Text, e.Node.Text);
                    var folders = APIHelper.GetFolders(tempFolder);
                    if (folders != null)
                        AddTreeNodeToItem(node, folders);
                }
                Cursor.Current = Cursors.Default;
            }
            else if (e.Node.Tag is ProLoopFolder)
            {
                var currentNode = e.Node;
                if (currentNode != null)
                {
                    foreach (TreeNode node in currentNode.Nodes)
                    {
                        var folderItem = node.Tag as ProLoopFolder;
                        var folders = APIHelper.GetFolders($"/api/files/{folderItem.Path}");
                        if (folders != null && folders.Count > 0)
                            AddTreeNodeToItem(node, folders);
                    }
                    //currentNode.Nodes.Clear();

                }

            }
        }

        private void AddTreeNodeToItem<T>(TreeNode node, List<T> items)
        {
            node.Nodes.Clear();
            foreach (var item in items)
            {
                string nameofItem = string.Empty;
                if (item is Organization)
                {
                    nameofItem = (item as Organization).Title;
                }
                else if (item is Client)
                {
                    nameofItem = (item as Client).Name;
                }
                else if (item is Matter)
                {
                    nameofItem = (item as Matter).Name;
                }
                else if (item is ProLoopFolder)
                {
                    if (!(item as ProLoopFolder).isDirctory)
                        continue;
                    nameofItem = (item as ProLoopFolder).Name;
                }
                else if (item is ProLoopFile)
                {
                    continue;
                }
                var treenode = new TreeNode(nameofItem);
                treenode.Tag = item;
                node.Nodes.Add(treenode);
            }
        }

        private void radioButtonOrganization_CheckedChanged(object sender, EventArgs e)
        {
            LoadOrgProject();
        }

        private void checkBoxShowPath_CheckedChanged(object sender, EventArgs e)
        {
            ShowHidePath();
        }

        private void checkBoxFileTree_CheckedChanged(object sender, EventArgs e)
        {
            splitContainer1.Panel1Collapsed = !checkBoxFileTree.Checked;
            isActionbyTreeView = false;
        }

        private void easyCompletionComboBoxOrgProj_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (radioButtonOrganization.Checked)
            {
                var selectedOrg = easyCompletionComboBoxOrgProj.SelectedItem as Organization;
                if (selectedOrg != null)
                {
                    ObjOrganization = selectedOrg;
                    ObjProject = null;
                    var clients = APIHelper.GetClients(selectedOrg.Id);
                    if (clients.Count == 0)
                    {
                        easyCompletionComboxClient.DataSource = null;
                        client = string.Empty;
                    }
                    orgProject = selectedOrg.Title;
                    easyCompletionComboxClient.DataSource = clients;
                    easyCompletionComboxClient.MatchingMethod = StringMatchingMethod.NoWildcards;
                    easyCompletionComboxClient.DisplayMember = "Name";
                    easyCompletionComboxClient.ValueMember = "id";
                    if (isActionbyTreeView)
                    {
                        var item = clients.Find(x => x.Name == client);
                        if (item != null)
                        {
                            easyCompletionComboxClient.SelectedItem = item;
                        }
                    }
                    else
                    {
                        isActionbyTreeView = false;
                    }
                }
                else
                {
                    easyCompletionComboxClient.SelectedIndex = -1;
                }
            }
            if (radioButtonProject.Checked)
            {
                client = string.Empty;
                matter = string.Empty;
                var selectedOrg = easyCompletionComboBoxOrgProj.SelectedItem as Project;
                if (selectedOrg != null)
                {
                    ObjOrganization = null;
                    orgProject = selectedOrg.Title;
                    ObjProject = selectedOrg;
                    string folderPath = "/api/files/Projects/" + orgProject + "/*?token=";                   
                    GetFiles(folderPath);
                }
            }
        }

        private void easyCompletionComboBoxMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedMatter = easyCompletionComboBoxMatter.SelectedItem as Matter;
            dataGridViewFileDetail.DataSource = null;
            if (selectedMatter != null)
            {
                matter = selectedMatter.Name;
                ObjMatter = selectedMatter;
                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    orgProject,
                                    "/",
                                    client,
                                    "/",
                                   matter,
                                    "/*?token="
                            });
                //ProcessAutoComplete(); 
             
                GetFiles(folderPath);
            }
        }
        private void LoadOrgProject()
        {
            dataGridViewFileDetail.DataSource = null;
            if (radioButtonOrganization.Checked)
            {
                easyCompletionComboxClient.Enabled = true;
                easyCompletionComboBoxMatter.Enabled = true;
                if (orgs.Count == 0)
                {
                    easyCompletionComboBoxOrgProj.DataSource = null;
                    return;
                }
                if (easyCompletionComboBoxOrgProj.DataSource == null)
                {
                    easyCompletionComboBoxOrgProj.DataSource = orgs;
                    easyCompletionComboBoxOrgProj.MatchingMethod = StringMatchingMethod.NoWildcards;
                    easyCompletionComboBoxOrgProj.DisplayMember = "title";
                    easyCompletionComboBoxOrgProj.ValueMember = "id";
                }
                if (isActionbyTreeView)
                {
                    var item = orgs.Find(x => x.Title == orgProject);
                    if (item != null)
                    {
                        easyCompletionComboBoxOrgProj.SelectedItem = item;
                    }
                }
                else
                {
                    easyCompletionComboBoxOrgProj.SelectedIndex = 0;
                }
                easyCompletionComboxClient.Enabled = true;
                easyCompletionComboBoxMatter.Enabled = true;
            }
            else if (radioButtonProject.Checked)
            {
                easyCompletionComboxClient.Enabled = false;
                easyCompletionComboBoxMatter.Enabled = false;
                if (projects.Count == 0)
                {
                    easyCompletionComboBoxOrgProj.DataSource = null;
                    return;
                }
                if (easyCompletionComboBoxOrgProj.DataSource == null)
                {
                    easyCompletionComboBoxOrgProj.DataSource = projects;
                    easyCompletionComboBoxOrgProj.MatchingMethod = StringMatchingMethod.NoWildcards;
                    easyCompletionComboBoxOrgProj.DisplayMember = "title";
                    easyCompletionComboBoxOrgProj.ValueMember = "id";
                }                
                if (isActionbyTreeView)
                {
                    var item = projects.Find(x => x.Title == orgProject);
                    if (item != null)
                    {
                        easyCompletionComboBoxOrgProj.SelectedItem = item;
                    }
                }
                else
                {
                    easyCompletionComboBoxOrgProj.SelectedIndex = 0;
                }
                easyCompletionComboxClient.DataSource = null;
                easyCompletionComboBoxMatter.DataSource = null;
                easyCompletionComboxClient.Enabled = false;
                easyCompletionComboBoxMatter.Enabled = false;
            }
        }
        private void ShowHidePath()
        {
            if (dataGridViewFileDetail.Columns.Count > 0)
                dataGridViewFileDetail.Columns[1].Visible = checkBoxShowPath.Checked;
        }

        private void easyCompletionComboxClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedClient = easyCompletionComboxClient.SelectedItem as Client;
            var selectedOrg = easyCompletionComboBoxOrgProj.SelectedItem as Organization;
            if (selectedClient != null && selectedOrg != null)
            {
                ObjClient = selectedClient;
                var matters = APIHelper.GetMatters(selectedOrg.Id, selectedClient.Id);
                if (matters.Count == 0)
                {
                    matter = string.Empty;
                    easyCompletionComboBoxMatter.DataSource = null;
                    return;
                }
                orgProject = selectedOrg.Title;
                client = selectedClient.Name;
                easyCompletionComboBoxMatter.DataSource = matters;
                easyCompletionComboBoxMatter.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxMatter.DisplayMember = "Name";
                easyCompletionComboBoxMatter.ValueMember = "id";
                if (isActionbyTreeView)
                {
                    var item = matters.Find(x => x.Name == client);
                    if (item != null)
                    {
                        easyCompletionComboBoxMatter.SelectedItem = item;
                    }
                }
                else
                {
                    isActionbyTreeView = false;
                }
            }
            else
            {
                easyCompletionComboBoxMatter.SelectedIndex = -1;
            }
        }
        private void GetFiles(string folderPath)
        {
            if (!AddinCurrentInstance.isOldViewEnable)
            {
                folderPath = folderPath.Replace("files", "dir");
            }
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var files = APIHelper.GetFiles(folderPath);
            var objectList = new List<MetaDataInfo>();
            if (files != null)
                foreach (var file in files)
                {
                    if (!file.isDirctory)
                        objectList.Add(new MetaDataInfo() { Name = file.Name, Path = file.Path, VersionId = file.VersionId });
                }
            dataGridViewFileDetail.AllowUserToAddRows = true;
            dataGridViewFileDetail.DataSource = objectList;
            ShowHidePath();
            dataGridViewFileDetail.AllowUserToAddRows = false;
            pictureBox1.Visible = false;
            pictureBox1.SendToBack();
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            TreeNode selectedNode = treeView1.SelectedNode;
            if (selectedNode == null || selectedNode.Tag == null)
                return;
            //isActionbyTreeView = true;
            if (selectedNode.Tag is Organization)
            {
                //txtClient.Enabled = true;
                //txtMatter.Enabled = true;
                orgProject = selectedNode.Text;
                if (radioButtonOrganization.Checked)
                {
                    var selectOrg = orgs.Find(x => x.Title == orgProject);
                    if (selectOrg != null)
                    {
                        easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;                        
                        ObjOrganization = selectOrg;
                    }
                }
                else
                {
                    radioButtonOrganization.Checked = true;
                    easyCompletionComboBoxOrgProj.Text = orgProject;
                }

                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    orgProject,
                                    "/*?token="
                            });
                ObjProject = null;
                GetFiles(folderPath);
            }
            else if (selectedNode.Tag is Client)
            {
                //txtClient.Enabled = true;
                //txtMatter.Enabled = true;
                orgProject = selectedNode.Parent.Text;
                client = selectedNode.Text;
                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    orgProject,
                                    "/",
                                    client,
                                    "/*?token="
                            });
                //ProcessAutoComplete();  
                //searchPath = $"Organizations/{orgProject}";
                if (radioButtonOrganization.Checked)
                {
                    var selectOrg = orgs.Find(x => x.Title == orgProject);
                    if (selectOrg != null)
                    {
                        easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;                       
                    }
                }
                else
                {
                    radioButtonOrganization.Checked = true;
                    easyCompletionComboBoxOrgProj.Text = orgProject;                   
                }
                easyCompletionComboxClient.Text = client;
                GetFiles(folderPath);
                ObjProject = null;
            }
            else if (selectedNode.Tag is Matter)
            {
                //isActionbyTreeView = true;
                orgProject = selectedNode.Parent.Parent.Text;
                client = selectedNode.Parent.Text;
                matter = selectedNode.Text;
                if (radioButtonOrganization.Checked)
                {
                    var selectOrg = orgs.Find(x => x.Title == orgProject);
                    if (selectOrg != null)
                    {
                        easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;                       
                        
                    }
                }
                else
                {
                    radioButtonOrganization.Checked = true;
                    easyCompletionComboBoxOrgProj.Text = orgProject;
                }
                easyCompletionComboxClient.Text = client;
                easyCompletionComboBoxMatter.Text = matter;

                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    orgProject,
                                    "/",
                                    client,
                                    "/",
                                   matter = selectedNode.Text,
                                    "/*?token="
                            });
                //ProcessAutoComplete();  
                //searchPath = $"Organizations/{orgProject}";
                ObjProject = null;
                GetFiles(folderPath);
            }
            else if (selectedNode.Tag is Project)
            {
                orgProject = selectedNode.Text;
                string folderPath = "/api/files/Projects/" + orgProject + "/*?token=";
                //searchPath = $"Projects/{orgProject}";
                if (radioButtonProject.Checked)
                {
                    var selectOrg = projects.Find(x => x.Title == orgProject);
                    if (selectOrg != null)
                    {
                        ObjProject = selectOrg;
                        easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;
                    }
                }
                else
                {
                    radioButtonProject.Checked = true;
                    easyCompletionComboBoxOrgProj.Text = orgProject;
                }              
                GetFiles(folderPath);
                ObjOrganization = null;
                ObjMatter = null;
                ObjClient = null;
                //ProcessAutoComplete();
            }
            else if (selectedNode.Tag is ProLoopFolder)
            {
                string folderPath = (selectedNode.Tag as ProLoopFolder).Path;
                var pathCollaction = folderPath.Split('/');
                orgProject = pathCollaction[1];
                List<string> pathList = pathCollaction.ToList();
                if (folderPath.Contains("Projects"))
                {
                    if (radioButtonProject.Checked)
                    {
                        var selectOrg = projects.Find(x => x.Title == orgProject);
                        if (selectOrg != null)
                        {
                            ObjProject = selectOrg;
                            easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;
                        }
                    }
                    else
                    {
                        easyCompletionComboBoxOrgProj.Text = orgProject;
                        radioButtonProject.Checked = true;
                    }
                    
                    foreach(string data in pathList.Skip(2))
                    {
                        FolderPath = data + "\\";
                    }                   
                    ObjOrganization = null;
                    ObjMatter = null;
                    ObjClient = null;
                }
                else
                {
                    foreach (string data in pathList.Skip(4))
                    {
                        FolderPath = data + "\\";
                    }
                    //txtClient.Enabled = true;
                    //txtMatter.Enabled = true;
                    client = pathCollaction[2];
                    matter = pathCollaction[3];
                    if (radioButtonOrganization.Checked)
                    {
                        var selectOrg = orgs.Find(x => x.Title == orgProject);
                        if (selectOrg != null)
                        {
                            easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;                            
                        }
                    }
                    else
                    {
                        radioButtonOrganization.Checked = true;
                    }
                    easyCompletionComboxClient.Text = client;
                    easyCompletionComboBoxMatter.Text = matter;
                    ObjProject = null;
                }
                //searchPath = folderPath;
                folderPath = $"/api/files/{folderPath}/*?token=";
                GetFiles(folderPath);
                FolderPath = FolderPath.TrimEnd(new char[] { '\\' });
            }
            else
            {
                client = string.Empty;
                matter = string.Empty;
                orgProject = string.Empty;
            }
        }       

        private void ADXWordOpenNewTaskPane_Load(object sender, EventArgs e)
        {
            orgs = APIHelper.GetOrganizations();
            var treeNodeOrg = new TreeNode($"Organizations ({orgs.Count})");
            treeNodeOrg.Tag = "Organization";
            treeView1.Nodes[0].Nodes.Add(treeNodeOrg);
            foreach (var org in orgs)
            {
                var node = new TreeNode(org.Title);
                node.Tag = org;
                treeView1.Nodes[0].Nodes[0].Nodes.Add(node);
            }
            projects = APIHelper.GetProjects();
            var treenodeProject = new TreeNode($"Projects ({projects.Count})");
            treenodeProject.Tag = "project";
            treeView1.Nodes[0].Nodes.Add(treenodeProject);
            foreach (var project in projects)
            {
                var node = new TreeNode(project.Title);
                node.Tag = project;
                treeView1.Nodes[0].Nodes[1].Nodes.Add(node);
            }
            splitContainer1.Panel1Collapsed = !checkBoxFileTree.Checked;
            ShowHidePath();
            radioButtonOrganization.Checked = true;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFileName.Text))
            {
                MessageBox.Show("File Name is required", "Proloop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string filePath = AddinCurrentInstance.WebDAVMappedDriveLetter;
            string localFolderPath = "/api/files";
            if (ObjProject != null && !string.IsNullOrEmpty(ObjProject.Title))
            {
                if (!string.IsNullOrEmpty(FolderPath))
                {
                    filePath += "\\Projects\\" + ObjProject.Title + "\\" + FolderPath + "\\" + txtFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                    localFolderPath += $"/Projects/{ObjProject.Title}/{FolderPath}";
                }
                else
                {
                    filePath += "\\Projects\\" + ObjProject.Title + "\\" + txtFileName.Text;
                    localFolderPath += $"/Projects/{ObjProject.Title}";
                }
            }
            else
            {
                if (ObjMatter != null && !string.IsNullOrEmpty(ObjMatter.Name))
                {
                    if (!string.IsNullOrEmpty(FolderPath))
                    {
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                    FolderPath + "\\" + txtFileName.Text;
                        localFolderPath += $"/Organizations/{ObjOrganization.Title}/{ObjClient.Name}/{ObjMatter.Name}";
                    }
                    else
                    {
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + ObjMatter.Name + "\\" +
                                    txtFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        localFolderPath += $"/Organizations/{ObjOrganization.Title}/{ObjClient.Name}/{ObjMatter.Name}";
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(FolderPath))
                    {
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + FolderPath + "\\" +
                                    txtFileName.Text; // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        localFolderPath += $"/Organizations/{ObjOrganization.Title}/{ObjClient.Name}/{FolderPath}";
                    }
                    else
                    {
                        filePath += "\\Organizations\\" + ObjOrganization.Title + "\\" + ObjClient.Name + "\\" + txtFileName.Text;
                        // + "?token=" + AddinCurrentInstance.ProLoopToken;
                        localFolderPath += $"/Organizations/{ObjOrganization.Title}/{ObjClient.Name}";
                    }
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
                    if (!string.IsNullOrWhiteSpace(AddinCurrentInstance.SelectedFile.LockingUserId) && AddinCurrentInstance.SelectedFile.LockingUserId != AddinCurrentInstance.CurrentUserId.ToString())
                    {
                        var message = "This Document is Checked Out by Rick Shaul.Would you like to Save As a new Document?";
                        var resposne = AddinCurrentInstance.DisplayWaranigMessage(message);
                        if (resposne)
                        {
                            AddinCurrentInstance.WordApp.ActiveDocument.SaveAs(filePath, missing, missing, missing, AddtoRecent, missing,
                                missing, missing, missing, missing, missing, missing, missing, missing, missing);
                            //SaveSettingChange();
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
                        //SaveSettingChange();
                    }
                }
                else
                {
                    AddinCurrentInstance.WordApp.ActiveDocument.SaveAs(filePath, missing, missing, missing, AddtoRecent, missing,
                        missing, missing, missing, missing, missing, missing, missing, missing, missing);
                    //SaveSettingChange();
                }
                localFolderPath += "/*?token";
                GetFiles(localFolderPath);
                isActionbyTreeView = false;
                FolderPath = string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to save document", "Proloop", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
