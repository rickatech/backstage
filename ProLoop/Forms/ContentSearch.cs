using Newtonsoft.Json;
using ProLoop.WordAddin.Model;
using ProLoop.WordAddin.Utils;
using Serilog;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ProLoop.WordAddin.Forms
{
    public partial class ContentSearch : Form
    {
        SearchParameter _searchParameter;
        public string FilePath { get; set; }
        public ContentSearch(SearchParameter searchParameter)
        {
            InitializeComponent();
            _searchParameter = searchParameter;
            label1.Text = string.Empty;
        }

        private void ContentSearch_Load(object sender, EventArgs e)
        {
            if (_searchParameter != null)
            {
                if (!string.IsNullOrEmpty(_searchParameter.OrgOrProject))
                    label1.Text = _searchParameter.OrgOrProjectName;
                clienttextBox.Text = _searchParameter.ClientName;
                mattertextBox.Text = _searchParameter.MatterName;
            }
        }

        private void documentContenttextBox_TextChanged(object sender, EventArgs e)
        {
            //if(e.cha)
            //ProcessAutoComplete();
        }
        private void ProcessAutoComplete()
        {
            pictureBox1.Visible = true;
            pictureBox1.BringToFront();
            var client = new WebClient();
            string url = string.Empty;
            if (!string.IsNullOrEmpty(documentContenttextBox.Text) && !string.IsNullOrEmpty(AddinModule.CurrentInstance.ProLoopUrl))
            {
                if (AddinModule.CurrentInstance.ProLoopUrl.EndsWith("/"))
                    url = $"{AddinModule.CurrentInstance.ProLoopUrl}api/search/index/find?body={documentContenttextBox.Text}";
                else
                    url = $"{AddinModule.CurrentInstance.ProLoopUrl}/api/search/index/find?body={documentContenttextBox.Text}";
                Log.Debug($"Processing {url} in ProcessAutoComplete()");
                client.DownloadStringAsync(new Uri(url));
                client.DownloadStringCompleted += Client_DownloadStringCompleted;
            }
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
                  
                    return;
                }
                if (e.Error == null)
                {
                    var jsonstring = e.Result;
                    if (!string.IsNullOrEmpty(jsonstring))
                        ObjectList = JsonConvert.DeserializeObject<List<MetaDataInfo>>(jsonstring);
                    dataGridView1.AllowUserToAddRows = true;
                    dataGridView1.DataSource = ObjectList;
                    dataGridView1.AllowUserToAddRows = false;
                }            
            }
            catch (Exception ex)
            {
                Log.Debug($"Failed to get response due to {ex.Message} in Client_DownloadStringCompleted()");
            }
            pictureBox1.Visible = false;
            pictureBox1.SendToBack();
        }

        private void buttonCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonOpen_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows != null)
            {
                FilePath = dataGridView1.SelectedRows[0].Cells[1].Value as string;
                _searchParameter.KeyWord = dataGridView1.SelectedRows[0].Cells[2].Value as string;
                _searchParameter.EditorName = dataGridView1.SelectedRows[0].Cells[3].Value as string;
                this.DialogResult = DialogResult.OK;
                this.Close();
            }

        }

        private void documentContenttextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                ProcessAutoComplete();
            }
        }
    }
}
