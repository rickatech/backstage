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

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordInfoTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;
        //Microsoft.Office.Core.DocumentProperties properties;       
        private SettingsForm settingForm;

        public ADXWordInfoTaskPane()
        {
            InitializeComponent();
            AddinCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
        }

        private void ADXWordInfoTaskPane_Activated(object sender, EventArgs e)
        {
            ProcessMetaData();
        }
        private void ProcessMetaData()
        {
            GetKeyWords();
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
            textBoxAuthor.ReadOnly = true;
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

                PostKeyword();
                textBoxSummary.ReadOnly = true;
                textBoxAuthor.ReadOnly = true;
               
            }
            catch (Exception ex)
            {

            }
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
            var data = MetadataHandler.GetMetaDataInfo<MetaDataInfo>(filePath);
            if(data!=null)
            {
                textBoxSummary.Text = string.Empty;
                int index = 0;
                foreach(var item in data)
                {
                    if (index == data.Count - 1)
                    {
                        textBoxSummary.Text += item.Keywords;
                    }
                    else
                    {
                        textBoxSummary.Text += item.Keywords+",";
                    }
                }
               
            }
            
        }

        private string PathBuilder()
        {
            string orionalLocation = AddinCurrentInstance.WordApp.ActiveDocument.FullName;
            if(orionalLocation.Contains("://"))
            {
                Uri uri = new Uri(orionalLocation);
                string localurl = uri.LocalPath.Replace("/dav",string.Empty);
                return localurl;
            }
            
            return null;
        }
        private void buttonSum_Click(object sender, EventArgs e)
        {           
            textBoxSummary.ReadOnly = false;
            
        }

        private void buttonAuther_Click(object sender, EventArgs e)
        {
            textBoxAuthor.ReadOnly = false;
            
        }
    }
}