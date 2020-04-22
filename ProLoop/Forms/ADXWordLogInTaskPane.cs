using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using AddinExpress.WD;
using AddinExpress.MSO;
using ProLoop.WordAddin.Service;
using Serilog;

namespace ProLoop.WordAddin.Forms
{
    public partial class ADXWordLogInTaskPane: AddinExpress.WD.ADXWordTaskPane
    {
        private const int EM_SETCUEBANNER = 5377;
        private readonly AddinModule AddinModuleCurrentInstance;
        public ADXWordLogInTaskPane()
        {
            InitializeComponent();
            this.AddinModuleCurrentInstance = (ADXAddinModule.CurrentInstance as AddinModule);
            this.Load += ADXWordLogInTaskPanl_Load;
        }

        private void ADXWordLogInTaskPanl_Load(object sender, EventArgs e)
        {
            this.txtUrl.Text = this.AddinModuleCurrentInstance.ProLoopUrl;
            this.txtUsername.Text = this.AddinModuleCurrentInstance.ProLoopUsername;
            this.txtPassword.Text = this.AddinModuleCurrentInstance.ProLoopPassword;
            if (!string.IsNullOrEmpty(AddinModuleCurrentInstance.ProLoopUsername))
            {
                labelCurrentLoginUser.Text = $"Current Login User: {AddinModuleCurrentInstance.ProLoopUsername}";
                labelCurrentLoginUser.Visible = true;
            }
            
            this.SetSaveButtonStatus();
            try
            {
                SendMessage(this.txtUrl.Handle, 5377, 0, "https://newui.proloop.com/");
                SendMessage(this.txtUsername.Handle, 5377, 0, "Username");
                SendMessage(this.txtPassword.Handle, 5377, 0, "Password");
            }
            catch { }
            
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int SendMessage(IntPtr hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        private void btnClear_Click(object sender, EventArgs e)
        {
            if (!String.IsNullOrEmpty(txtUrl.Text))
                txtUrl.Clear();
            if (!String.IsNullOrEmpty(txtUsername.Text))
                txtUsername.Clear();
            if (!String.IsNullOrEmpty(txtPassword.Text))
                txtPassword.Clear();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtUsername.Text) ||
                string.IsNullOrEmpty(txtPassword.Text) ||
                string.IsNullOrEmpty(txtUrl.Text))
            {
                MessageBox.Show("Please enter username and password with server address", "Sign in error", MessageBoxButtons.OK, MessageBoxIcon.Information);

                return;
            }
            else if (!txtUrl.Text.Contains("https:"))
            {
                MessageBox.Show("URL should be https enable", "Sign in error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            else
            {
                string text = txtUrl.Text.Trim();
                this.AddinModuleCurrentInstance.ProLoopUsername = this.txtUsername.Text.Trim();
                this.AddinModuleCurrentInstance.ProLoopPassword = this.txtPassword.Text.Trim();
                var result = APIHelper.InitClient(new Uri(text));
                if (result)
                {
                    if (!text.EndsWith("/"))
                    {
                        Serilog.Log.Debug<string>("URL doesn't end with / : {0}", text);
                        Serilog.Log.Debug("Adding the trailing / to the URL");
                        text += "/";
                        Serilog.Log.Debug<string>("Updated URL: {0}", text);
                    }
                    this.AddinModuleCurrentInstance.ProLoopUrl = text;

                    this.AddinModuleCurrentInstance.WebDAVValuesUpdated = true;
                    // APIHelper.InitClient(new Uri(AddinModuleCurrentInstance.ProLoopUrl));
                    base.DialogResult = DialogResult.OK;
                    if (this.AddinModuleCurrentInstance.SaveValuesToRegistry(this.AddinModuleCurrentInstance.ProLoopUrl, this.AddinModuleCurrentInstance.ProLoopUsername, this.AddinModuleCurrentInstance.ProLoopPassword))
                    {
                        MessageBox.Show("Successfully saved the values.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        this.AddinModuleCurrentInstance.UnmapAndMapWebDavFolderToDrive();
                    }
                    else
                    {
                        MessageBox.Show("Error while saving the values.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    }
                }
                else
                {
                    MessageBox.Show("Unable to sign in to the ProLoop. Please check the username and password.", "Sign in error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }            
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

        private void checkBoxOldView_CheckedChanged(object sender, EventArgs e)
        {
            AddinModuleCurrentInstance.DisplayOldView(checkBoxOldView.Checked);
        }
    }
}
