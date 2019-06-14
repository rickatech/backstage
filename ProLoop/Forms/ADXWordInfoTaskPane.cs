using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using AddinExpress.WD;
using AddinExpress.MSO;
using Serilog;
using System.Web.Script.Serialization;
using ProLoop.WordAddin.Utils;
using ProLoop.WordAddin.Model;
using System.Text;
using System.Net;
using System.Collections.Specialized;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordInfoTaskPane : AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;
        //Microsoft.Office.Core.DocumentProperties properties;       
        private SettingsForm settingForm;
        private ProLoopFile CurrentFile = null;
        public ADXWordInfoTaskPane()
        {
            InitializeComponent();
            AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            label4.Text = string.Empty;
            groupBox1.Visible = false;
            labelEditors.Text = string.Empty;           
        }

        private void ADXWordInfoTaskPane_Activated(object sender, EventArgs e)
        {
            ProcessMetaData();
        }
        private void ProcessMetaData()
        {
            GetKeyWords();
            GetEditor();
            #region inbuiltProp

            //if (rootMeta.Meta != null)
            //{
            //    textBoxAuthor.Text = rootMeta.Meta.Author;
            //    textBoxSummary.Text = rootMeta.Meta.Summary;
            //}
            //else
            //{
            //    properties = (Microsoft.Office.Core.DocumentProperties)
            //       AddinCurrentInstance.WordApp.ActiveDocument.BuiltInDocumentProperties;
            //    var author = properties["Author"];
            //    if (author != null)
            //    {

            //        textBoxAuthor.Text = author.Value as string;

            //    }
            //    var summmary = properties["Comments"];
            //    if (summmary != null)
            //    {
            //        textBoxSummary.Text = summmary.Value as string;

            //    }             

            //}
            #endregion

            textBoxSummary.ReadOnly = true;           
            label1.Text = AddinCurrentInstance.WordApp.ActiveDocument.Name;
        }

        private void btnSettings_Click(object sender, EventArgs e)
        {
            Serilog.Log.Debug("btnSettings_Click() -- Begin");

            ShowSettingsForm();

#if NETFX45
            Task.Factory.StartNew(AddinCurrentInstance.UnmapAndMapWebDavFolderToDrive);
#elif NETFX35
            ThreadPool.QueueUserWorkItem(obj => { AddinCurrentInstance.UnmapAndMapWebDavFolderToDrive(); });
#endif

            Serilog.Log.Debug("btnSettings_Click() -- End");
        }
        private void ShowSettingsForm()
        {
            Serilog.Log.Debug("ShowSettingsForm() -- Begin");

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

            Serilog.Log.Debug("ShowSettingsForm() -- End");
        }

        private void buttonRefresh_Click(object sender, EventArgs e)
        {
            ProcessMetaData();
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            try
            {
                //save update data
                if (!textBoxSummary.ReadOnly)
                {
                    PostKeyword();
                }                
                textBoxSummary.ReadOnly = true;               

            }
            catch (Exception ex)
            {

            }
        }
        private void PostEditor()
        {
            string filePath = PathBuilder();
            if (filePath == null)
                return;
            string data = "Editor=";
            MetadataHandler.PostMetaDataInfo(data, filePath);
        }

        private void PostKeyword()
        {
            string filePath = PathBuilder();
            if (filePath == null)
                return;
            string data = "keywords=" + textBoxSummary.Text;
            MetadataHandler.PostMetaDataInfo(data, filePath);
        }
        private void GetKeyWords()
        {
            string filePath = PathBuilder();
            if (string.IsNullOrEmpty(filePath))
                return;
            filePath = $"{AddinCurrentInstance.ProLoopUrl}/api/filetags{filePath}";
            var data = MetadataHandler.GetMetaDataInfo<MetaDataInfo>(filePath);
            if (data != null)
            {
                textBoxSummary.Text = string.Empty;
                int index = 0;
                foreach (var item in data)
                {
                    if (index == data.Count - 1)
                    {
                        textBoxSummary.Text += item.Keywords;
                        label4.Text = $"Doc ID:{item.VersionId}";
                        ProcessMetaDataInfo(item.VersionId);
                    }
                    else
                    {
                        textBoxSummary.Text += item.Keywords + ",";
                    }
                }

            }

        }
        private void ProcessMetaDataInfo(string docId)
        {
            string[] metaInfo = docId.Split('-');
            if (metaInfo != null && metaInfo.Length == 5)
            {
                lblorg.Text = $"Org: {RemovedZeroData(metaInfo[0])}";
                lblClient.Text = $"Client: { RemovedZeroData(metaInfo[1])}";
                lblMatter.Text = $"Matter: { RemovedZeroData(metaInfo[2])}";
                lbDocId.Text = $"Doc Id: { RemovedZeroData(metaInfo[3])}";
                lblVersion.Text = $"Version: {RemovedZeroData(metaInfo[4])}";
            }
        }
        private string RemovedZeroData(string data)
        {
            char[] chartrim = { '0' };
            return data.TrimStart(chartrim);
        }

        private void GetEditor()
        {
            try
            {
                string filePath = PathBuilder();
                if (string.IsNullOrEmpty(filePath))
                    return;
                filePath = $"{AddinCurrentInstance.ProLoopUrl}/api/sayt/f{filePath}";
                var data = MetadataHandler.GetMetaDataInfo<ProLoopFile>(filePath);
                if (data != null && data.Count > 0)
                {
                    ProcessCheckInCheckout(data[0]);
                    GetHistory(data[0].Id);
                }
            }
            catch (Exception ex)
            {

            }
        }
        private void ProcessCheckInCheckout(ProLoopFile currentFile)
        {
            groupBox1.Visible = true;
            this.CurrentFile = currentFile;
            //File is already Checkout
            if (!string.IsNullOrWhiteSpace(currentFile.LockingUserId))
            {
                labelCheckin.Text = "Check Out :";
                labelMessage.Text = "Locked";
                buttonCheckin.Visible = true;
                buttonCheckout.Visible = false;
                labelDetail.Text = $"Check out by {currentFile.LockingUserId} on {DateTime.Now.ToShortDateString()}";
                labelDetail.Visible = true;
            }
            else
            {
                labelCheckin.Text = "Check Out :";
                labelMessage.Text = "Not Locked";
                buttonCheckin.Visible = false;
                buttonCheckout.Visible = true;
                labelDetail.Visible = false;
            }
        }
        private void GetHistory(string fileId)
        {
            string url = $"{AddinCurrentInstance.ProLoopUrl}/api/file/{fileId}/history";
            var data = MetadataHandler.GetMetaData<DataInfoRequest>(url);
            if (data != null && data.History != null)
            {
               
                StringBuilder sbAuthor = new StringBuilder();
                foreach (var history in data.History)
                {
                    if (!sbAuthor.ToString().Contains(history.Username))
                        sbAuthor.Append(history.Username + ",");
                }
                if (!string.IsNullOrEmpty(sbAuthor.ToString()))
                    labelEditors.Text = sbAuthor.ToString().Trim(new char[] { ',' });
            }

        }
        private string PathBuilder()
        {
            string orionalLocation = AddinCurrentInstance.WordApp.ActiveDocument.FullName;
            if (orionalLocation.Contains("://"))
            {
                Uri uri = new Uri(orionalLocation);
                string localurl = uri.LocalPath.Replace("/dav", string.Empty);
                return localurl;
            }

            return null;
        }
        private void buttonSum_Click(object sender, EventArgs e)
        {
            textBoxSummary.ReadOnly = false;

        }
        private void buttonCheckout_Click(object sender, EventArgs e)
        {
            using (WebClient client = new WebClient())
            {
                try
                {
                    client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                    var values = new NameValueCollection();
                    values["token"] = AddinCurrentInstance.ProLoopToken;
                    var result = client.UploadValues($"{AddinCurrentInstance.ProLoopUrl}/api/file/{this.CurrentFile.Id }/lock/{AddinCurrentInstance.CurrentUserId}", values);
                    var str = Encoding.Default.GetString(result);
                    if (str != null)
                    {
                        this.CurrentFile.LockingUserId = AddinCurrentInstance.CurrentUserId.ToString();
                        ProcessCheckInCheckout(this.CurrentFile);
                    }
                }
                catch (Exception ex)
                {

                }

            }
        }
        private void buttonCheckin_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(this.CurrentFile.LockingUserId) && CurrentFile.LockingUserId != AddinCurrentInstance.CurrentUserId.ToString())
            {
                var message = $"This Document Checked Out by {CurrentFile.LockingUserId} Would you like to Force this Check In? If you do so, {CurrentFile.LockingUserId} will receive an email notification";
                var response = AddinCurrentInstance.DisplayWaranigMessage(message);
                if (response)
                {
                    using (WebClient client = new WebClient())
                    {
                        try
                        {
                            client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                            var values = new NameValueCollection();
                            values["token"] = AddinCurrentInstance.ProLoopToken;
                            var result = client.UploadValues($"{AddinCurrentInstance.ProLoopUrl}/api/file/{this.CurrentFile.Id }/unlock", values);
                            var str = Encoding.Default.GetString(result);
                            if (str != null)
                            {
                                this.CurrentFile.LockingUserId = "";
                                ProcessCheckInCheckout(this.CurrentFile);
                            }
                        }
                        catch (Exception ex)
                        {

                        }

                    }
                }
            }
            else
            {
                using (WebClient client = new WebClient())
                {
                    try
                    {
                        client.Headers[HttpRequestHeader.ContentType] = "application/x-www-form-urlencoded";
                        var values = new NameValueCollection();
                        values["token"] = AddinCurrentInstance.ProLoopToken;
                        var result = client.UploadValues($"{AddinCurrentInstance.ProLoopUrl}/api/file/{this.CurrentFile.Id }/unlock", values);
                        var str = Encoding.Default.GetString(result);
                        if (str != null)
                        {
                            this.CurrentFile.LockingUserId = "";
                            ProcessCheckInCheckout(this.CurrentFile);
                        }
                    }
                    catch (Exception ex)
                    {

                    }

                }
            }

        }
    }
}
