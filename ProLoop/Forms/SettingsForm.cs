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
               var result= APIHelper.InitClient(new Uri(text));
                if (result)
                {
                    if (!text.EndsWith("/"))
                    {
                        Log.Debug<string>("URL doesn't end with / : {0}", text);
                        Log.Debug("Adding the trailing / to the URL");
                        text += "/";
                        Log.Debug<string>("Updated URL: {0}", text);
                    }
                    this.AddinModuleCurrentInstance.ProLoopUrl = text;

                    this.AddinModuleCurrentInstance.WebDAVValuesUpdated = true;
                    // APIHelper.InitClient(new Uri(AddinModuleCurrentInstance.ProLoopUrl));
                    base.DialogResult = DialogResult.OK;
                }
                else
                {
                    MessageBox.Show("Unable to sign in to the ProLoop. Please check the username and password.", "Sign in error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
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
