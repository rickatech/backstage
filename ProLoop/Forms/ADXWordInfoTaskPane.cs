using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using AddinExpress.MSO;
using Serilog;
using System.Web.Script.Serialization;
using ProLoop.WordAddin.Utils;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordInfoTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private readonly AddinModule AddinCurrentInstance;
        Microsoft.Office.Core.DocumentProperties properties;
        RootMeta rootMeta = new RootMeta();
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
            #region inbuiltProp
            DataDeserialization();
            if (rootMeta.Meta != null)
            {
                textBoxAuthor.Text = rootMeta.Meta.Author;
                textBoxSummary.Text = rootMeta.Meta.Summary;
            }
            else
            {
                properties = (Microsoft.Office.Core.DocumentProperties)
                   AddinCurrentInstance.WordApp.ActiveDocument.BuiltInDocumentProperties;
                var author = properties["Author"];
                if (author != null)
                {

                    textBoxAuthor.Text = author.Value as string;

                }
                var summmary = properties["Comments"];
                if (summmary != null)
                {
                    textBoxSummary.Text = summmary.Value as string;

                }
               
                #endregion
            }
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
                DataSerialization();
                textBoxSummary.ReadOnly = true;
                textBoxAuthor.ReadOnly = true;
               
            }
            catch (Exception ex)
            {

            }
        }

        private void DataSerialization()
        {
            string filePath = PathBuilder();
            rootMeta.Meta = new Metadata()
            {
                Author = textBoxAuthor.Text.Trim(new Char[] { ' ', ',', '.' }),
                Summary = textBoxSummary.Text.Trim()
            };
            JavaScriptSerializer js = new JavaScriptSerializer();
            var data = js.Serialize(rootMeta);
            MetadataHandler.GenerateNewMetadataFile(filePath, data);
        }
        private RootMeta DataDeserialization()
        {
            string filePath = PathBuilder();
           var content= MetadataHandler.GetMetadatOfFile(filePath);
            if (!string.IsNullOrEmpty(content))
            {
                JavaScriptSerializer js = new JavaScriptSerializer();
                rootMeta = js.Deserialize<RootMeta>(content);
                return rootMeta;
            }
            return null;
        }

        private string PathBuilder()
        {
            string orionalLocation = AddinCurrentInstance.WordApp.ActiveDocument.FullName;
            return orionalLocation;
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
