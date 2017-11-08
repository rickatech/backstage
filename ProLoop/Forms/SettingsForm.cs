using AddinExpress.MSO;
using ProLoop.WordAddin.Service;
using Serilog;
using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ProLoop.WordAddin.Forms
{
    public partial class SettingsForm : Form
    {
        private const int EM_SETCUEBANNER = 5377;

        private readonly AddinModule AddinModuleCurrentInstance;
        public SettingsForm()
        {
            InitializeComponent();
            this.AddinModuleCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);
        private void btnClear_Click(object sender, EventArgs e)
        {
            this.txtUrl.Clear();
            this.txtUsername.Clear();
            this.txtPassword.Clear();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            string text = this.txtUrl.Text.Trim().ToLower();
            
            if (!text.StartsWith("http://")|| !text.StartsWith("https://"))
            {
                Log.Debug<string>("URL doesn't start with http:// or https:// : {0}", text);
                Log.Debug("Prefixing the URL with https://");
                text = "https://" + text;
                Log.Debug<string>("Updated URL: {0}", text);
            }
           
            if (text.StartsWith("http://"))
            {
                Log.Debug<string>("URL starts with http:// : {0}", text);
                Log.Debug("Replacing the prefix from http:// to https://");
                text = text.Replace("http://", "https://");
                Log.Debug<string>("Updated URL: {0}", text);
            }
            APIHelper.InitClient(new Uri(text));
            if (!text.EndsWith("/"))
            {
                Log.Debug<string>("URL doesn't end with / : {0}", text);
                Log.Debug("Adding the trailing / to the URL");
                text += "/";
                Log.Debug<string>("Updated URL: {0}", text);
            }
            this.AddinModuleCurrentInstance.ProLoopUrl = text;
            this.AddinModuleCurrentInstance.ProLoopUsername = this.txtUsername.Text.Trim();
            this.AddinModuleCurrentInstance.ProLoopPassword = this.txtPassword.Text.Trim();
            this.AddinModuleCurrentInstance.WebDAVValuesUpdated = true;
            APIHelper.InitClient(new Uri(AddinModuleCurrentInstance.ProLoopUrl));
            base.DialogResult = DialogResult.OK;
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            this.txtUrl.Text = this.AddinModuleCurrentInstance.ProLoopUrl;
            this.txtUsername.Text = this.AddinModuleCurrentInstance.ProLoopUsername;
            this.txtPassword.Text = this.AddinModuleCurrentInstance.ProLoopPassword;
            this.SetSaveButtonStatus();
            SendMessage(this.txtUrl.Handle, 5377, 0, "https://newui.proloop.com/");
            SendMessage(this.txtUsername.Handle, 5377, 0, "Username");
            SendMessage(this.txtPassword.Handle, 5377, 0, "Password");
        }
        private void SetSaveButtonStatus()
        {
            if (string.IsNullOrEmpty(txtUrl.Text) || string.IsNullOrEmpty(txtUsername.Text) ||
               string.IsNullOrEmpty(txtPassword.Text))
                btnSave.Enabled = false;
            else
                btnSave.Enabled = true;
        }

        private void txtUrl_TextChanged(object sender, EventArgs e)
        {
            SetSaveButtonStatus();
        }

        private void txtUsername_TextChanged(object sender, EventArgs e)
        {
            SetSaveButtonStatus();
        }

        private void txtPassword_TextChanged(object sender, EventArgs e)
        {
            SetSaveButtonStatus();
        }
    }
}
