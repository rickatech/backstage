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

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordOpenNewTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        public ADXWordOpenNewTaskPane()
        {
            InitializeComponent();
            this.Load += ADXWordOpenNewTaskPane_Load;
        }

        private void ADXWordOpenNewTaskPane_Load(object sender, EventArgs e)
        {
            List<Organization> orgs = APIHelper.GetOrganizations();
            var treeNodeOrg = new TreeNode($"Organizations ({orgs.Count})");
            treeNodeOrg.Tag = "Organization";
            treeView1.Nodes[0].Nodes.Add(treeNodeOrg);
            foreach (var org in orgs)
            {
                var node = new TreeNode(org.Title);
                node.Tag = org;
                treeView1.Nodes[0].Nodes[0].Nodes.Add(node);
            }
            List<Project> projects = APIHelper.GetProjects();
            var treenodeProject = new TreeNode($"Projects ({projects.Count})");
            treenodeProject.Tag = "project";
            treeView1.Nodes[0].Nodes.Add(treenodeProject);
            foreach (var project in projects)
            {
                var node = new TreeNode(project.Title);
                node.Tag = project;
                treeView1.Nodes[0].Nodes[1].Nodes.Add(node);
            }
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
            if(selectedNode.Tag is Matter)
            {
                txtClient.Enabled = true;
                txtMatter.Enabled = true;
                txtOrgProject.Text = selectedNode.Parent.Parent.Text;
                txtClient.Text = selectedNode.Parent.Text;
                txtMatter.Text = selectedNode.Text;
                var folderPath = string.Concat(new string[]
                            {
                                    "/api/files/Organizations/",
                                    txtOrgProject.Text,
                                    "/",
                                    txtClient.Text,
                                    "/",
                                    txtMatter.Text = selectedNode.Text,
                                    "/*?token="
                            });
                //ProcessAutoComplete();
                GetFiles(folderPath);
            }
            else if(selectedNode.Tag is Project)
            {
                txtClient.Enabled = false;
                txtMatter.Enabled = false;
                txtOrgProject.Text = selectedNode.Text;
                string folderPath = "/api/files/Projects/" + txtOrgProject.Text + "/*?token=";
                GetFiles(folderPath);
                //ProcessAutoComplete();
            }
            else if(selectedNode.Tag is ProLoopFolder)
            {
                string folderPath = (selectedNode.Tag as ProLoopFolder).Path;
                var pathCollaction = folderPath.Split('/');
                txtOrgProject.Text = pathCollaction[1];
                if (folderPath.Contains("Projects"))
                {
                    txtClient.Enabled = false;
                    txtMatter.Enabled = false;                  
                }
                else
                {
                    txtClient.Enabled = true;
                    txtMatter.Enabled = true;
                    txtClient.Text = pathCollaction[2];
                    txtMatter.Text = pathCollaction[3];
                }
                folderPath = $"/api/files/{folderPath}/*?token=";
                GetFiles(folderPath);
            }
        }
        private void GetFiles(string folderPath)
        {
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var files = APIHelper.GetFiles(folderPath);
            var objectList = new List<MetaDataInfo>();
            foreach(var file in files)
            {
                if (!file.isDirctory)
                    objectList.Add(new MetaDataInfo() { Name = file.Name, Path = file.Path, VersionId = file.VersionId });
            }
            dataGridView1.AllowUserToAddRows = true;
            dataGridView1.DataSource = objectList;
            dataGridView1.AllowUserToAddRows = false;
            pictureBox1.Visible = false;
            pictureBox1.SendToBack();
        }
        private void ProcessAutoComplete()
        {
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var client = new WebClient();
            string url = string.Empty;

            if (string.IsNullOrEmpty(txtDocId.Text) && string.IsNullOrEmpty(txtFileName.Text) && string.IsNullOrEmpty(txtEditor.Text)
                && string.IsNullOrEmpty(txtKeyword.Text) && string.IsNullOrEmpty(textBoxContent.Text))
            {
                url = $"{WordAddin.AddinModule.CurrentInstance.ProLoopUrl}/api/sayt/f/?keywords={txtKeyword.Text}&editor={txtEditor.Text}&s={txtFileName.Text}&fid=223&body={textBoxContent.Text}";
            }
            else
            {
                url = $"{WordAddin.AddinModule.CurrentInstance.ProLoopUrl}/api/sayt/f/?keywords={txtKeyword.Text}&editor={txtEditor.Text}&s={txtFileName.Text}&fid={txtDocId.Text}&body={textBoxContent.Text}";
            }

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
                    dataGridView1.DataSource = ObjectList;
                    dataGridView1.AllowUserToAddRows = false;
                    pictureBox1.Visible = false;
                    return;
                }
                var jsonstring = e.Result;
                if (!string.IsNullOrEmpty(jsonstring))
                    ObjectList = JsonConvert.DeserializeObject<List<MetaDataInfo>>(jsonstring);
                dataGridView1.Invoke(new Action(() =>
                {
                    dataGridView1.AllowUserToAddRows = true;
                    dataGridView1.DataSource = ObjectList;
                    dataGridView1.AllowUserToAddRows = false;
                    pictureBox1.Visible = false;
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
    }
}
