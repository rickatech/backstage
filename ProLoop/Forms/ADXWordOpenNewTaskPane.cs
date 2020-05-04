using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using ProLoop.WordAddin.Service;
using System.Collections.Generic;
using System.Net;
using Newtonsoft.Json;
using ProLoop.WordAddin.Model;
using AddinExpress.MSO;
using Microsoft.Office.Interop.Word;
using System.IO;
using ProLoop.WordAddin.Utils;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordOpenNewTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private string searchPath = string.Empty;
        private readonly AddinModule AddinCurrentInstance;
        private string orgProject, client, matter;
        List<Organization> orgs = new List<Organization>();
        List<Project> projects = new List<Project>();
        private bool isActionbyTreeView = false;
        public ADXWordOpenNewTaskPane()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            this.Load += ADXWordOpenNewTaskPane_Load;
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
                    foreach(TreeNode node in currentNode.Nodes)
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
            foreach(var item in items)
            {
                string nameofItem = string.Empty;
                if(item is Organization)
                {
                    nameofItem = (item as Organization).Title;
                }
                else if(item is Client)
                {
                    nameofItem = (item as Client).Name;
                }
                else if (item is Matter)
                {
                    nameofItem = (item as Matter).Name;
                }
                else if(item is ProLoopFolder)
                {
                    if (!(item as ProLoopFolder).isDirctory)
                        continue;
                    nameofItem = (item as ProLoopFolder).Name;
                }
                else if(item is ProLoopFile)
                {
                    continue;
                }
                var treenode = new TreeNode(nameofItem);
                treenode.Tag = item;
                node.Nodes.Add(treenode);
            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            TreeNode selectedNode = treeView1.SelectedNode;
            if (selectedNode == null || selectedNode.Tag == null)
                return;
            ClearSearchText();
            if (selectedNode.Tag is Organization)
            {
                //txtClient.Enabled = true;
                //txtMatter.Enabled = true;
                orgProject = selectedNode.Text;
                //var selectOrg = orgs.Find(x => x.Title == orgProject);
                //if (selectOrg != null)
                //{
                //    easyCompletionComboBoxOrgProj.SelectedItem = selectOrg;
                //}
                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    orgProject,
                                    "/*?token="
                            });
                //ProcessAutoComplete();  
                searchPath = $"Organizations/{orgProject}";
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
                searchPath = $"Organizations/{orgProject}";
                GetFiles(folderPath);
            }
            else if (selectedNode.Tag is Matter)
            {
                isActionbyTreeView = true;
                orgProject = selectedNode.Parent.Parent.Text;
                client = selectedNode.Parent.Text;
                matter = selectedNode.Text;
                if (radioButtonOrganization.Checked)
                {
                    easyCompletionComboBoxOrgProj.Text = orgProject;
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
                    searchPath = $"Organizations/{orgProject}";
                    GetFiles(folderPath);
                    isActionbyTreeView = false;
                }
                else
                {
                    radioButtonOrganization.Checked = true;
                }
            }
            else if (selectedNode.Tag is Project)
            {
                orgProject = selectedNode.Text;
                string folderPath = "/api/files/Projects/" + orgProject + "/*?token=";
                searchPath = $"Projects/{orgProject}";
                if (radioButtonProject.Checked)
                {
                    easyCompletionComboBoxOrgProj.Text = orgProject;
                    GetFiles(folderPath);
                    isActionbyTreeView = false;
                }
                else
                {
                    radioButtonProject.Checked = true;
                }
               
                //ProcessAutoComplete();
            }
            else if (selectedNode.Tag is ProLoopFolder)
            {
                string folderPath = (selectedNode.Tag as ProLoopFolder).Path;
                var pathCollaction = folderPath.Split('/');
                orgProject = pathCollaction[1];
                if (folderPath.Contains("Projects"))
                {
                    // txtClient.Enabled = false;
                    //txtMatter.Enabled = false;                  
                }
                else
                {
                    //txtClient.Enabled = true;
                    //txtMatter.Enabled = true;
                    client = pathCollaction[2];
                    matter = pathCollaction[3];
                }
                searchPath = folderPath;
                folderPath = $"/api/files/{folderPath}/*?token=";
                GetFiles(folderPath);
            }
            else
            {
                client = string.Empty;
                matter = string.Empty;
                orgProject = string.Empty;
            }
        }
        private void ClearSearchText()
        {
            dataGridViewFileDetail.Refresh();
            txtKeyword.Text = string.Empty;
            txtFileName.Text = string.Empty;
            txtEditor.Text = string.Empty;
            txtDocId.Text = string.Empty;
            textBoxContent.Text = string.Empty;
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
            foreach(var file in files)
            {
                if (!file.isDirctory)
                    objectList.Add(new MetaDataInfo() { Name = file.Name, Path = file.Path + "/" + file.Name, VersionId = file.VersionId });
            }
            dataGridViewFileDetail.AllowUserToAddRows = true;
            dataGridViewFileDetail.DataSource = objectList;
            ShowHidePath();
            dataGridViewFileDetail.AllowUserToAddRows = false;
            pictureBox1.Visible = false;
            pictureBox1.SendToBack();
        }
        private void ProcessAutoComplete()
        {
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var client = new WebClient();
            string url = string.Empty;

            //if (string.IsNullOrEmpty(txtDocId.Text) && string.IsNullOrEmpty(txtFileName.Text) && string.IsNullOrEmpty(txtEditor.Text)
            //    && string.IsNullOrEmpty(txtKeyword.Text) && string.IsNullOrEmpty(textBoxContent.Text))
            //{
            //    url = $"{WordAddin.AddinModule.CurrentInstance.ProLoopUrl}/api/sayt/f/?keywords={txtKeyword.Text}&editor={txtEditor.Text}&s={txtFileName.Text}&fid=223&body={textBoxContent.Text}";
            //}
            //else
            //{
            url = $"{WordAddin.AddinModule.CurrentInstance.ProLoopUrl}/api/sayt/f/{searchPath}?keywords={txtKeyword.Text}&editor={txtEditor.Text}&s={txtFileName.Text}&fid={txtDocId.Text}&body={textBoxContent.Text}";
            //}

            //Log.Debug($"Processing {url} in ProcessAutoComplete()");
            client.DownloadStringAsync(new Uri(url));
            client.DownloadStringCompleted += Client_DownloadStringCompleted;
        }

        private void Client_DownloadStringCompleted(object sender, DownloadStringCompletedEventArgs e)
        {
            try
            {
                var ObjectList = new List<MetaDataInfo>();
                if (e == null)
                {
                    dataGridViewFileDetail.DataSource = ObjectList;
                    dataGridViewFileDetail.AllowUserToAddRows = false;
                    ShowHidePath();
                    pictureBox1.Visible = false;
                    return;
                }
                var jsonstring = e.Result;
                if (!string.IsNullOrEmpty(jsonstring))
                    ObjectList = JsonConvert.DeserializeObject<List<MetaDataInfo>>(jsonstring);
                dataGridViewFileDetail.Invoke(new Action(() =>
                {
                    dataGridViewFileDetail.AllowUserToAddRows = true;
                    if (ObjectList != null)
                    {
                        if (ObjectList.Count > 10)
                        {
                            ObjectList = ObjectList.GetRange(0, 10);
                        }
                        ObjectList = ObjectList.FindAll(x => x.Path.Contains("."));
                    }
                    dataGridViewFileDetail.DataSource = ObjectList;
                    dataGridViewFileDetail.AllowUserToAddRows = false;
                    pictureBox1.Visible = false;
                    ShowHidePath();
                    pictureBox1.SendToBack();
                }));
                
            }
            catch (Exception ex)
            {
               // Log.Debug($"Failed to get response due to {ex.Message} in Client_DownloadStringCompleted()");
                pictureBox1.Visible = false;
                pictureBox1.SendToBack();
            }
        }

        private void txtDocId_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtEditor_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtFileName_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtKeyword_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void textBoxContent_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void dataGridViewFileDetail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewFileDetail.SelectedRows != null)
            {
                string selectedFile = dataGridViewFileDetail.SelectedRows[0].Cells[1].Value as string;
                AddinCurrentInstance.fileMetadata = new FileMetadataInfo()
                {
                    ClientName = client,
                    MatterName = matter,
                    ProjectName = orgProject
                };
                OpenDocument(selectedFile);
            }
        }
        private void OpenDocument(string currentFile)
        {
            if (string.IsNullOrEmpty(AddinCurrentInstance.WebDAVMappedDriveLetter))
                AddinCurrentInstance.MapWebDAVFolderToDriveLetter();
            string filePath = AddinCurrentInstance.WebDAVMappedDriveLetter;
            filePath += "\\" + currentFile.Replace("/", "\\");
            string extension = Path.GetExtension(filePath);
            
            if (AddinCurrentInstance.docExtension.Contains(extension))
            {
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
                    document = documents.Open(filePath, AddToRecentFiles: false, ReadOnly: false);
                    string fileInfo = $"{AddinCurrentInstance.ProLoopUrl}/api/filetags/{currentFile}";
                    var data = MetadataHandler.GetMetaDataInfo<MetaDataInfo>(fileInfo);
                    if (data != null)
                    {
                        if (document != null)
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
                    }

                }
                catch (Exception exception)
                {

                }
            }
            else
            {
                MessageBox.Show("You can only Open Doc or Docx file.", "Proloop", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void radioButtonProject_CheckedChanged(object sender, EventArgs e)
        {
           // LoadOrgProject();
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
                    orgProject = selectedOrg.Title;
                    string folderPath = "/api/files/Projects/" + orgProject + "/*?token=";
                    searchPath = $"Projects/{orgProject}";
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
                searchPath = $"Organizations/{orgProject}";
                GetFiles(folderPath);                
            }
        }

        private void easyCompletionComboxClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedClient = easyCompletionComboxClient.SelectedItem as Client;
            var selectedOrg = easyCompletionComboBoxOrgProj.SelectedItem as Organization;
            if (selectedClient != null && selectedOrg != null)
            {
                var matters = APIHelper.GetMatters(selectedOrg.Id, selectedClient.Id);
                if (matters.Count==0)
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
            }
            else
            {
                easyCompletionComboBoxMatter.SelectedIndex = -1;
            }
        }

        private void ShowHidePath()
        {
            if (dataGridViewFileDetail.Columns.Count > 0)
                dataGridViewFileDetail.Columns[1].Visible = checkBoxShowPath.Checked;
        }
        private void LoadOrgProject()
        {
            dataGridViewFileDetail.DataSource = null;
            if (radioButtonOrganization.Checked)
            {
                if (orgs.Count == 0)
                {
                    easyCompletionComboBoxOrgProj.DataSource = null;
                    return;
                }
                easyCompletionComboBoxOrgProj.DataSource = orgs;
                easyCompletionComboBoxOrgProj.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxOrgProj.DisplayMember = "title";
                easyCompletionComboBoxOrgProj.ValueMember = "id";
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
                if (projects.Count == 0)
                {
                    easyCompletionComboBoxOrgProj.DataSource = null;
                    return;
                }
                easyCompletionComboBoxOrgProj.DataSource = projects;
                easyCompletionComboBoxOrgProj.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxOrgProj.DisplayMember = "title";
                easyCompletionComboBoxOrgProj.ValueMember = "id";
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
    }
}
