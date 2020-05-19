using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using System.Collections.Generic;
using AddinExpress.MSO;
using ProLoop.WordAddin.Service;
using ProLoop.WordAddin.Model;
using Newtonsoft.Json;
using System.Net;
using System.IO;
using Microsoft.Office.Interop.Word;
using ProLoop.WordAddin.Utils;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordTaskPaneOpenDemand: AddinExpress.WD.ADXWordTaskPane
    {
        private const int EM_SETCUEBANNER = 5377;
        private string searchPath = string.Empty;
        private readonly AddinModule AddinCurrentInstance;
        private Project currentProject;
        private Client currentClient;
        private Matter currentMatter;
        private Organization currentOrganization;
        List<Organization> orgs = new List<Organization>();
        List<Project> projects = new List<Project>();
        private bool isActionbyTreeView = false;
        public ADXWordTaskPaneOpenDemand()
        {
            InitializeComponent();
            this.AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            this.Load += ADXWordTaskPaneOpenDemand_Load;
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);
        private void ADXWordTaskPaneOpenDemand_Load(object sender, EventArgs e)
        {
            SendMessage(txtboxEditors.Handle, EM_SETCUEBANNER, 0, "Editors");
            SendMessage(txtboxContent.Handle, EM_SETCUEBANNER, 0, "Content");
            SendMessage(txtboxTags.Handle, EM_SETCUEBANNER, 0, "Tags");
            SendMessage(txtboxDocumentId.Handle, EM_SETCUEBANNER, 0, "Document Id");
            SendMessage(txtboxTitle.Handle, EM_SETCUEBANNER, 0, "Title");
            ProcessControlBox();
            ProcessFolderTree();
        }
        private void ProcessFolderTree()
        {            
            scTreeViewInformation.Panel2Collapsed = true;
        }
        private void easyCompletionComboBoxOrg_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!easyCompletionComboBoxOrg.Focused)
            {
                easyCompletionComboBoxClient.SelectedIndex = -1;
                easyCompletionComboBoxClient.Text = "---Select Client---";
                return;
            }
            var selectedOrg = easyCompletionComboBoxOrg.SelectedItem as Organization;
            if (selectedOrg != null)
            {
                var clients = APIHelper.GetClients(selectedOrg.Id);
                if (clients.Count == 0)
                {
                    easyCompletionComboBoxClient.DataSource = null;
                    currentClient = null;
                }
                easyCompletionComboBoxClient.DataSource = clients;
                easyCompletionComboBoxClient.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxClient.DisplayMember = "Name";
                easyCompletionComboBoxClient.ValueMember = "id";                
                easyCompletionComboBoxClient.Text = "---Select Client---";
                currentOrganization = selectedOrg;
            }
           
        }

        private void easyCompletionComboBoxClient_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!easyCompletionComboBoxClient.Focused)
            {
                easyCompletionComboBoxMatter.SelectedItem = null;
                easyCompletionComboBoxMatter.Text = "---Select Matter---";
                return;
            }
            var selectedClient = easyCompletionComboBoxClient.SelectedItem as Client;
            var selectedOrg = easyCompletionComboBoxOrg.SelectedItem as Organization;
            if (selectedClient != null && selectedOrg != null)
            {
                var matters = APIHelper.GetMatters(selectedOrg.Id, selectedClient.Id);
                if (matters.Count == 0)
                {
                    easyCompletionComboBoxMatter.DataSource = null;
                    return;
                }
                easyCompletionComboBoxMatter.DataSource = matters;
                easyCompletionComboBoxMatter.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxMatter.DisplayMember = "Name";
                easyCompletionComboBoxMatter.ValueMember = "id";
                currentOrganization = selectedOrg;
                currentClient = selectedClient;

            }
            else
            {
                easyCompletionComboBoxMatter.SelectedItem = null;
                easyCompletionComboBoxMatter.Text = "---Select Matter---";
            }
        }

        private void easyCompletionComboBoxMatter_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectedMatter = easyCompletionComboBoxMatter.SelectedItem as Matter;
            if (easyCompletionComboBoxMatter.Focused)
            {
                dataGridViewFileDetail.DataSource = null;
                if (selectedMatter != null)
                {
                    currentMatter = selectedMatter;                 
                    var folderPath = string.Concat(new string[]
                                {
                                    "/api/dir/Organizations/",
                                    currentOrganization.Title,
                                    "/",
                                    currentClient.Name,
                                    "/",
                                   currentMatter.Name                                    
                                });
                    //ProcessAutoComplete();  
                    searchPath = $"Organizations/{currentOrganization.Title}";
                    GetFiles(folderPath);
                    LoadTreeView(searchPath);
                }
            }
            else
            {
                easyCompletionComboBoxMatter.SelectedItem = null;
                easyCompletionComboBoxMatter.Text = "---Select Matter---";
            }
          
        }

        private void rbOrganization_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
        }
        private void ProcessControlBox()
        {
            currentMatter = null;
            treeView1.Nodes.Clear();
            lalSourceUrl.Text = "Source Url";
            currentProject = null;currentProject = null;currentClient = null;
            easyCompletionComboBoxClient.Enabled = rbOrganization.Checked;
            easyCompletionComboBoxClient.DataSource = null;
            easyCompletionComboBoxMatter.Enabled = rbOrganization.Checked;
            easyCompletionComboBoxMatter.DataSource = null;
            easyCompletionComboBoxProject.Enabled = rbProject.Checked;
            easyCompletionComboBoxProject.DataSource = null;
            easyCompletionComboBoxOrg.DataSource = null;
            easyCompletionComboBoxOrg.Enabled = rbOrganization.Checked;            
            pictureBox1.Location = new System.Drawing.Point(dataGridViewFileDetail.ClientSize.Width / 2, dataGridViewFileDetail.Height/2);
        }
        private void LoadData()
        {
            ProcessControlBox();
            if (rbOrganization.Checked)
            {
                if (orgs.Count == 0)
                {
                    orgs = APIHelper.GetOrganizations();
                }
                easyCompletionComboBoxOrg.DataSource = orgs;
                easyCompletionComboBoxOrg.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxOrg.DisplayMember = "title";
                easyCompletionComboBoxOrg.ValueMember = "id";
                easyCompletionComboBoxOrg.SelectedItem = null;
                easyCompletionComboBoxOrg.Text = "---Select Organization---";
            }
            else if (rbProject.Checked)
            {
                if (projects.Count == 0)
                {
                    projects = APIHelper.GetProjects();
                }
                easyCompletionComboBoxProject.DataSource = projects;
                easyCompletionComboBoxProject.MatchingMethod = StringMatchingMethod.NoWildcards;
                easyCompletionComboBoxProject.DisplayMember = "title";
                easyCompletionComboBoxProject.ValueMember = "id";
                easyCompletionComboBoxProject.SelectedItem = null;
                easyCompletionComboBoxProject.Text = "---Select Project---";
            }
        }

        private void easyCompletionComboBoxProject_SelectedIndexChanged(object sender, EventArgs e)
        {
            var selectProject = easyCompletionComboBoxProject.SelectedItem as Project;
            var cbo = (EasyCompletionComboBox)sender;
            if (!cbo.Focused)
                return;
            if (selectProject != null)
            {
                currentProject = selectProject;
                string folderPath = "/api/dir/Projects/" + currentProject.Title;
                searchPath = $"Projects/{currentProject.Title}";
                GetFiles(folderPath);
                LoadTreeView(folderPath);
            }
        }
        private void LoadTreeView(string currentPath)
        {
            treeView1.Nodes.Clear();
            if (currentMatter != null)
            {
                var treeNode = new TreeNode();
                treeNode.Name = "Node0";
                treeNode.Tag = currentMatter;
                treeNode.Text = currentMatter.Name;
                treeView1.Nodes.AddRange(new TreeNode[] { treeNode });
                var tempFolder = APIHelper.GetFolderPath("", currentOrganization.Title, currentMatter.Name, currentClient.Name);
                var folders = APIHelper.GetFolders(tempFolder);
                if (folders != null)
                    AddTreeNodeToItem(treeNode, folders);
            }
            else if (currentProject != null)
            {
                var treeNode = new TreeNode();
                treeNode.Name = "Node0";
                treeNode.Text = currentProject.Title;
                treeNode.Tag = currentProject;
                treeView1.Nodes.AddRange(new TreeNode[] { treeNode });
                var tempFolder = APIHelper.GetFolderPath(currentProject.Title, "", "", "");
                var folders = APIHelper.GetFolders(tempFolder);
                if (folders != null)
                    AddTreeNodeToItem(treeNode, folders);

            }
            else
            {
                treeView1.Nodes.Clear();
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
                if (nameofItem.Length == 1)
                    continue;
                var treenode = new TreeNode(nameofItem);
                treenode.Tag = item;
                node.Nodes.Add(treenode);
            }
        }
        private void ckbFileTree_CheckedChanged(object sender, EventArgs e)
        {            
            if (ckbFileTree.Checked)
            {
                scTreeViewInformation.Panel2Collapsed = false;
                scTreeViewInformation.Show();
                tableLayoutPanel1.RowStyles[4].Height = 50F;
                tableLayoutPanel1.RowStyles[4].SizeType = SizeType.Percent;
                tableLayoutPanel1.RowStyles[8].Height = 50F;
                tableLayoutPanel1.RowStyles[8].SizeType = SizeType.Percent;
            }
            else
            {
               
                tableLayoutPanel1.RowStyles[4].Height = 0;
                tableLayoutPanel1.RowStyles[4].SizeType = SizeType.AutoSize;
                tableLayoutPanel1.RowStyles[8].Height = 100F;
                tableLayoutPanel1.RowStyles[8].SizeType = SizeType.Percent;
                scTreeViewInformation.Panel2.Hide();
                scTreeViewInformation.Panel2Collapsed = true;
            }
        }

        private void ShowHidePath()
        {
            if (dataGridViewFileDetail.Columns.Count > 0)
                dataGridViewFileDetail.Columns[1].Visible = ckbShowPath.Checked;
        }
        private void GetFiles(string folderPath)
        {
            lalSourceUrl.Text = $"Source Url : {folderPath.Replace("/api/dir/",string.Empty)}";
            folderPath = ckbRecursive.Checked ? folderPath + "/*?token=" : folderPath + "?token=";
            if (!AddinCurrentInstance.isOldViewEnable)
            {
                folderPath = folderPath.Replace("files", "dir");
            }
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var files = APIHelper.GetFiles(folderPath);
            var objectList = new List<MetaDataInfo>();
            dataGridViewFileDetail.DataSource = null;
            foreach (var file in files)
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

        private void ckbShowPath_CheckedChanged(object sender, EventArgs e)
        {
            ShowHidePath();
        }

        private void dataGridViewFileDetail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewFileDetail.SelectedRows != null)
            {
                string selectedFile = dataGridViewFileDetail.SelectedRows[0].Cells[1].Value as string;
                AddinCurrentInstance.fileMetadata = new FileMetadataInfo();
                if (currentProject != null)
                {
                    AddinCurrentInstance.fileMetadata.ProjectName = currentProject.Title;
                }
                else
                {
                    if(currentOrganization!=null && currentMatter!=null && currentClient != null)
                    {
                        AddinCurrentInstance.fileMetadata.ClientName = currentClient.Name;
                        AddinCurrentInstance.fileMetadata.MatterName = currentMatter.Name;
                        AddinCurrentInstance.fileMetadata.ProjectName = currentOrganization.Title;
                    }
                }
               
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

        private void txtboxEditors_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtboxContent_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtboxTags_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtboxDocumentId_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void txtboxTitle_TextChanged(object sender, EventArgs e)
        {
            ProcessAutoComplete();
        }

        private void ADXWordTaskPaneOpenDemand_ADXSplitterMove(object sender, ADXSplitterMoveEventArgs e)
        {
            ProcessControlBox();
        }

        private void rbProject_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
        }

        private void ProcessAutoComplete()
        {          
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var client = new WebClient();
            string url = string.Empty;           
            url = $"{WordAddin.AddinModule.CurrentInstance.ProLoopUrl}/api/sayt/f/{searchPath}?keywords={txtboxTags.Text}&editor={txtboxEditors.Text}&s={txtboxTitle.Text}&fid={txtboxDocumentId.Text}&body={txtboxContent.Text}";
           
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

        private void treeView1_AfterExpand(object sender, TreeViewEventArgs e)
        {
            if (e.Node.Tag is null)
                return;           
            
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
                        var folders = APIHelper.GetFolders($"/api/dir/{folderItem.Path}/{folderItem.Name}");
                        if (folders != null)
                            AddTreeNodeToItem(node, folders);
                    }
                   
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
                        var folders = APIHelper.GetFolders($"/api/dir/{folderItem.Path}/{folderItem.Name}");
                        if (folders != null && folders.Count > 0)
                            AddTreeNodeToItem(node, folders);
                    }
                    //currentNode.Nodes.Clear();

                }

            }
        }

        private void treeView1_DoubleClick(object sender, EventArgs e)
        {
            TreeNode selectedNode = treeView1.SelectedNode;
            if (selectedNode == null || selectedNode.Tag == null)
                return;

            else if (selectedNode.Tag is Matter)
            {
                var folderPath = string.Concat(new string[]
                            {
                                    "/api/dir/Organizations/",
                                    currentOrganization.Title,
                                    "/",
                                    currentClient.Name,
                                    "/",
                                   currentMatter.Name,                                    
                            });
                GetFiles(folderPath);
            }
            else if (selectedNode.Tag is Project)
            {
                string folderPath = "/api/dir/Projects/" + currentProject.Title;
                GetFiles(folderPath);
            }
            else if (selectedNode.Tag is ProLoopFolder)
            {
                string folderPath = (selectedNode.Tag as ProLoopFolder).Path;                  
                searchPath = folderPath;
                folderPath = $"/api/dir/{folderPath}/{selectedNode.Text}";
                GetFiles(folderPath);
            }
            
        }
    }
}
