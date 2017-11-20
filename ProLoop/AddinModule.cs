#define NETFX45

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;
//using stdole;
using System.Security.AccessControl;
using AddinExpress.MSO;
using Microsoft.Win32;
using Win32;
using Word = Microsoft.Office.Interop.Word;
using ProLoop.WordAddin.Forms;

#if NETFX45
using System.Threading.Tasks;
using Serilog;
//using WebDAVSaveAsEntry.Properties;
#endif

#if NETFX35
using System.Threading;
//using WebDAVSaveAsEntry.Properties;
using System.Reflection;
#endif

namespace ProLoop.WordAddin
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("C2B7CD0B-D065-494E-B7DB-BA8327FFC5ED"), ProgId("ProLoop.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        private Word.Application WordApplication;
        private string WebDAVFolderMappedDriveLetter;
        public Operation Mode { get; set; } = Operation.Open;       
        public bool WebDAVValuesUpdated = false;

        public string ProLoopUrl { get; set; }     

        public string ProLoopPassword { get; set; }

        public string ProLoopUsername { get; set; }
        public string WebDAVUrl => this.ProLoopUrl + "/dav";
        public string WebDAVMappedDriveLetter { get; set; }
        public bool IsProLoopDocument { get; set; }
        public string ProLoopToken { get; set; }

        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();

            AddinInitialize += OnAddinInitialize;
            AddinStartupComplete += OnAddinStartupComplete;
        }

        private void OnAddinStartupComplete(object sender, EventArgs eventArgs)
        {
            //            Log.Verbose("Addin_OnStartupComplete() -- Begin");
            //            Log.Verbose("Addin started in Word Version {0}", this.HostVersion);

            //            WebDAVFolderMappedDriveLetter = string.Empty;          

            //#if NETFX45
            //            Task.Factory.StartNew(MapWebDAVFolderToDriveLetter);
            //#elif NETFX35
            //            ThreadPool.QueueUserWorkItem(obj => { MapWebDAVFolderToDriveLetter(); });
            //#endif

            //            Log.Verbose("Addin_OnStartupComplete() -- End");
            MakeInitAPIHelper();
        }

        private void MakeInitAPIHelper()
        {
            if (!String.IsNullOrEmpty(ProLoopUrl)|| !String.IsNullOrEmpty(ProLoopPassword)||
                !String.IsNullOrEmpty(ProLoopUsername)){
                Service.APIHelper.InitClient(new Uri(ProLoopUrl));
            }
        }

        private void OnAddinInitialize(object sender, EventArgs eventArgs)
        {
            try
            {
                Logger.LogFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Logger.SetupLogging();
            }
            catch (Exception exception)
            {
                Log.Error("Exception during Addin_OnConnection(): {0}", exception.Message);
#if DEBUG
                MessageBox.Show("Exception during Addin_OnConnection: " + exception.Message);
                if (exception.InnerException != null)
                    MessageBox.Show("Inner Exception: " + exception.InnerException.Message);
#endif
            }

            Log.Debug("OnAddinInitialize() -- Begin (after Logging setup)");

            AddinFinalize += OnAddinFinalize;

            // Load ProLoop values from registry and map a drive letter to WebDAV path
#if NETFX45
            Task.Factory.StartNew(LoadRegValuesAndMapDriveLetter);
#elif NETFX35
            ThreadPool.QueueUserWorkItem(obj => { LoadRegValuesMapDriveLetter(); });
#endif

            Log.Debug("OnAddinInitialize() -- End");
        }

        public void LoadRegValuesAndMapDriveLetter()
        {
            this.LoadValuesFromRegistry();
            this.MapWebDAVFolderToDriveLetter();
        }

        private void OnAddinFinalize(object sender, EventArgs eventArgs)
        {
            Log.Debug("OnAddinFinalize() -- Begin");

            UnMapWebDAVFolderToDriveLetter(true);

            Log.Debug("OnAddinFinalize() -- End");

            //Stop logging
            if (Log.Logger != null)
                Logger.CloseLogging();
        }

        public void LoadValuesFromRegistry()
        {
            Log.Debug("LoadValuesFromRegistry() -- Begin");
            this.ProLoopUrl = null;
            this.ProLoopUsername = null;
            this.ProLoopPassword = null;
            try
            {
                using (RegistryKey currentUser = Registry.CurrentUser)
                {
                    RegistryKey registryKey = currentUser.OpenSubKey("Software\\ProLoop", RegistryKeyPermissionCheck.ReadSubTree, RegistryRights.ExecuteKey);
                    object obj = (registryKey != null) ? registryKey.GetValue("URL") : null;
                    bool flag = obj != null;
                    if (flag)
                    {
                        this.ProLoopUrl = obj.ToString();
                        bool flag2 = this.ProLoopUrl.EndsWith("/");
                        if (flag2)
                        {
                            this.ProLoopUrl = this.ProLoopUrl.Substring(0, this.ProLoopUrl.Length - 1);
                        }
                        Log.Debug<string>("Registry URL value: {0}", this.ProLoopUrl);
                    }
                    object obj2 = (registryKey != null) ? registryKey.GetValue("Username") : null;
                    bool flag3 = obj2 != null;
                    if (flag3)
                    {
                        this.ProLoopUsername = obj2.ToString();
                        Log.Debug<string>("Registry Username value: {0}", this.ProLoopUsername);
                    }
                    object obj3 = (registryKey != null) ? registryKey.GetValue("Password") : null;
                    bool flag4 = obj3 != null;
                    if (flag4)
                    {
                        this.ProLoopPassword = obj3.ToString();
                        Log.Debug<string>("Registry Password value: {0}", this.ProLoopPassword);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Exception while loading the WebDAV settings from registry.");
                Log.Error(ex.Message);
            }
            Log.Debug("LoadValuesFromRegistry() -- End");
        }

        public bool SaveValuesToRegistry(string webDavFolderUrl, string webDavUsername, string webDavPassword)
        {
            Log.Debug("SaveValuesToRegistry() -- Begin");
            bool result = false;
            try
            {
                using (RegistryKey currentUser = Registry.CurrentUser)
                {
                    RegistryKey registryKey = currentUser.OpenSubKey("Software\\ProLoop", RegistryKeyPermissionCheck.ReadWriteSubTree, RegistryRights.SetValue);
                    bool flag = registryKey == null;
                    if (flag)
                    {
                        registryKey = currentUser.CreateSubKey("Software\\ProLoop", RegistryKeyPermissionCheck.ReadWriteSubTree);
                    }
                    if (registryKey != null)
                    {
                        registryKey.SetValue("URL", webDavFolderUrl, RegistryValueKind.String);
                    }
                    Log.Debug<string>("Registry URL value set to {0}", webDavFolderUrl);
                    if (registryKey != null)
                    {
                        registryKey.SetValue("Username", webDavUsername, RegistryValueKind.String);
                    }
                    Log.Debug<string>("Registry Username value set to {0}", webDavUsername);
                    if (registryKey != null)
                    {
                        registryKey.SetValue("Password", webDavPassword, RegistryValueKind.String);
                    }
                    Log.Debug<string>("Registry Password value set to {0}", webDavPassword);
                    result = true;
                    Log.Debug("SaveValuesToRegistry() -- End");
                }
            }
            catch (Exception ex)
            {
                result = false;
                Log.Verbose("Exception while saving the WebDAV settings to registry.");
                Log.Verbose(ex.Message);
            }
            return result;
        }


        private void UnMapWebDAVFolderToDriveLetter(bool forceUnmap)
        {
            Log.Verbose("UnMapWebDAVFolderToDriveLetter() -- Begin");

            //Disconnect the mapped WebDAV drive
            if (!string.IsNullOrEmpty(WebDAVFolderMappedDriveLetter))
            {
                var result = Utility.WNetCancelConnection2(WebDAVFolderMappedDriveLetter,
                    Utility.CONNECT_UPDATE_PROFILE, forceUnmap);

                if (result == 0) //Success
                {
                    Log.Verbose("Successfully unmapped the WebDAV folder from drive letter {0}",
                        WebDAVFolderMappedDriveLetter);
                    WebDAVFolderMappedDriveLetter = string.Empty;
                }
                else
                {
                    Log.Verbose("Unable to unmap the WebDAV folder. Please verify the network connectivity to WebDAV server");
                }
            }

            Log.Verbose("UnMapWebDAVFolderToDriveLetter() -- End");
        }

       

        private string GetUncFromUrl(string url)
        {
           
            Log.Debug("GetUncFromUrl() -- Begin");
            string str = string.Empty;
            string text = this.ProLoopUrl.ToUpper();           
            if (text.Contains("HTTPS://"))
            {
                text = text.Replace("HTTPS://", "");
                str = "@SSL";
            }
            else
            {
               
                if (text.Contains("HTTP://"))
                {
                    text = text.Replace("HTTP://", "");
                }
            }
           
            if (text.EndsWith("/"))
            {
                text = text.Substring(0, text.Length - 1);
            }
            string result = "\\\\" + text + str + "\\DavWWWRoot\\dav";
            Log.Debug("GetUncFromUrl() -- End");
            return result;
        }

        public void MapWebDAVFolderToDriveLetter()
        {
            Log.Debug("MapWebDAVFolderToDriveLetter() -- Begin");

            //Verify if a drive letter is unmapped and available for mapping
            var driveLetters = new List<string>
            {
                @"W:\",
                @"Z:\",
                @"Y:\",
                @"X:\",
                @"V:\",
                @"U:\",
                @"T:\",
                @"S:\",
                @"R:\",
                @"Q:\",
                @"P:\",
                @"O:\",
                @"N:\",
                @"M:\",
                @"L:\",
                @"K:\",
                @"J:\",
                @"I:\",
                @"H:\",
                @"G:\",
                @"F:\",
                @"E:\",
                @"D:\",
                @"B:\"
            };

            // Remove the drive letters which are already being used, from the available drive letter list
            var usedDrives = DriveInfo.GetDrives();
            foreach (var usedDrive in usedDrives)
            {
                var usedDriveLetter = usedDrive.Name; // C:\, E:\, etc:\
                var usedDriveType = usedDrive.DriveType;
                Log.Debug("Drive letter in use: {0}, Drive type: {1}", usedDriveLetter, usedDriveType);

                //Remove this drive letter from the unused drive letter list
                driveLetters.Remove(usedDriveLetter);
            }

            //Verify if any used drive letter is already mapped to the WebDAV folder
            try
            {
                var searcher = new ManagementObjectSearcher("select * from Win32_MappedLogicalDisk");
                foreach (var resultItem in searcher.Get())
                {
                    var drive = (ManagementObject)resultItem;
                    Log.Debug("Drive letter {0} is mapped to {1}", drive["Name"].ToString(),
                        drive["ProviderName"].ToString());

                    var mappedDrivePath = drive["ProviderName"].ToString();
                    var mappedFolderPathAsUnc = GetUncFromUrl(WebDAVUrl);

                    #region Hide
                    ////if (mappedDrivePath.ToUpper().StartsWith(@"\\NEWUI.PROLOOP.COM\DAVWWWROOT"))

                    //// Convert the URL to UNC: 
                    //// From http://newui.proloop.com/dav/ to \\NEWUI.PROLOOP.COM\DAV
                    //var mappedFolderPathAsUnc = WebDAVFolderURL.ToUpper();

                    //if (mappedFolderPathAsUnc.Contains("HTTPS"))
                    //{
                    //    mappedFolderPathAsUnc = mappedFolderPathAsUnc.Replace(@"HTTPS://", @"\\").Replace(@"/", @"\");
                    //}
                    //else
                    //{
                    //    mappedFolderPathAsUnc = mappedFolderPathAsUnc.Replace(@"HTTP://", @"\\").Replace(@"/", @"\");
                    //}

                    //if (mappedFolderPathAsUnc.EndsWith(@"\")) // Remove the trailing \
                    //    mappedFolderPathAsUnc = mappedFolderPathAsUnc.Substring(0, mappedFolderPathAsUnc.Length - 1);

                    #endregion

                    if (mappedDrivePath.ToUpper().StartsWith(mappedFolderPathAsUnc.ToUpper()))
                    {
                        //A drive letter is already mapped to WebDAV folder
                        WebDAVMappedDriveLetter = drive["Name"].ToString();
                        Log.Information("WebDAV folder {0} is already mapped to drive {1}",
                            mappedFolderPathAsUnc, WebDAVMappedDriveLetter);
                        return;
                    }
                    //Console.WriteLine(Regex.Match(drive["ProviderName"].ToString(), @"\\\\([^\\]+)").Groups[1]);
                }
            }
            catch (Exception exception)
            {
                Log.Error("Exception while iterating over drive letters and their mapped path.");
                Log.Error(exception.Message);

                MessageBox.Show("Error while iterating over drive letters and their mapped path.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);

                return;
            }
            if (string.IsNullOrEmpty(ProLoopUrl))
                return;

            //Map the WebDAV folder to the drive letter -- WebDAVFolderMappedDriveLetter
            //Get the first unused drive letter
            WebDAVMappedDriveLetter = driveLetters.First().Replace("\\", "");

            var networkResource = new Utility.NETRESOURCE
            {
                dwType = Utility.RESOURCETYPE_DISK,
                lpLocalName = WebDAVMappedDriveLetter,
                lpRemoteName = GetUncFromUrl(WebDAVUrl), //WebDAVUrl,
                lpProvider = null
            };

            // Hard coded: Authenticate using username and password
            //string usernameProLoopWebDAV = "sn";
            //string passwordProLoopWebDAV = "p@ssw0rd";
            //var result = Utility.WNetAddConnection2(ref networkResource, passwordProLoopWebDAV, usernameProLoopWebDAV, 0);
            //var result = Utility.WNetAddConnection2(ref networkResource, null, null, 0);

            var result = Utility.WNetAddConnection2(ref networkResource, ProLoopPassword, ProLoopUsername, 0);

            // TODO: Task 3: Handle exceptions when WebDAV folder is not accessible 
            // (access denied, permission errors, user account disabled, etc)
            // Error Code and description: https://msdn.microsoft.com/en-us/library/windows/desktop/ms681382(v=vs.85).aspx
            if (result == 0) // No Error
            {
                Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}", WebDAVMappedDriveLetter);
            }
            else if (result == 1244) //ERROR_NOT_AUTHENTICATED 
            {
                Log.Verbose(
                    "Unable to connect to the ProLoop server. Please verify the username and/or password.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop server. Please verify the username and/or password.",
                    "Authentication error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 53)
            {
                Log.Verbose(
                    "Unable to connect to the ProLoop server. Please verify the ProLoop URL in Settings.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop server. Please verify the ProLoop URL in Settings.",
                    "Settings error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 67) //ERROR_BAD_NET_NAME -- No network situation
            {
                Log.Verbose(
                    "Unable to connect to the ProLoop server. Please verify the network connectivity.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop server. Please verify the network connectivity.",
                    "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 487) //ERROR_INVALID_ADDRESS -- No URL or invalid address
            {
                Log.Verbose("Unable to connect to the ProLoop server. Please verify the URL.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop server. Please verify the URL.",
                    "Invalid Address", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 85) //ERROR_ALREADY_ASSIGNED -- Drive is already mapped
            {
                //Unmap the existing drive & remap
                Log.Verbose(
                    "The drive letter is already assigned to another network resource. \nWill attempt to unmap and remap");
                UnMapWebDAVFolderToDriveLetter(true);

                var resultNew = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
                if (resultNew == 0)
                {
                    WebDAVMappedDriveLetter = driveLetters.First().Replace("\\", "");
                    Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}",
                        WebDAVMappedDriveLetter);
                }
                else
                {
                    Log.Verbose(
                        "Unable to open the WebDAV folder. \nPlease verify the access permissions to WebDAV folder");
                    MessageBox.Show(
                        "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
                        "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Log.Verbose(
                    "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
                    "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //MessageBox.Show(result.ToString());

            Log.Verbose("MapWebDAVFolderToDriveLetter() -- End");
        }

        private void adxBackstageGroupButtonOpen_OnClick(object sender, IRibbonControl control)
        {
            Log.Verbose("OpenButton_Click() -- Begin");

            //if (Application.Documents.Count > 0)
            //{
            //    var currentDocument = Application.ActiveDocument;
            //    var currentDocumentName = currentDocument.Name;
            //    var currentDocumentHasExtension = currentDocumentName.IndexOf(".") > 0;
            //}

            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter =
                "Word 2007-2016 files (*.docx)|*.docx|Word 97-2003 files (*.doc)|*.doc|All files (*.*)|*.*";
            openFileDialog.InitialDirectory = WebDAVFolderMappedDriveLetter;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var documentToBeOpened = openFileDialog.FileName;
                try
                {
                    WordApplication.Documents.Open(documentToBeOpened);
                }
                catch (Exception exception)
                {
                    Log.Error("Error while opening the selected document file using File->Open.");
                    Log.Error(exception.Message);
                }
            }

            Log.Verbose("OpenButton_Click() -- End");
        }

        private void adxBackstageGroupButtonSave_OnClick(object sender, IRibbonControl control)
        {
            Log.Verbose("SaveAsButton_Click() -- Begin");

            //MessageBox.Show(String.Format("WebDAV Save As Entry Version {0}", this.GetType().Assembly.GetName().Version),
            //	"About WebDAV Save As Entry", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            var currentDocument = WordApplication.ActiveDocument;

            var currentDocumentName = currentDocument.Name;
            var currentDocumentHasExtension = currentDocumentName.IndexOf(".") > 0;

            //Present a Save As dialog
            var saveFileDialog = new SaveFileDialog();

            if (currentDocumentHasExtension)
                saveFileDialog.FileName = currentDocumentName;
            else
                saveFileDialog.FileName = "New Doc.docx";

            saveFileDialog.Filter =
                "Word 2007-2016 files (*.docx)|*.docx|Word 97-2003 files (*.doc)|*.doc|All files (*.*)|*.*";

            //TODO: Select the correct type of Filter based on the type of file opened (.docx/.doc)
            //saveFileDialog.FilterIndex = currentDocument.

            saveFileDialog.InitialDirectory = WebDAVFolderMappedDriveLetter;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                var filterIndex = saveFileDialog.FilterIndex;

                //// Show the Properties dialog
                //SaveAsUserControl saveAsUserControl = new SaveAsUserControl();

                //TextBox NameTextBoxControl = (TextBox)saveAsUserControl.Controls.Find("txtName", true)[0];
                //NameTextBoxControl.Text = saveFileDialog.FileName;


                //saveAsUserControl.ShowDialog();

                //TODO: Add support for other file types (with Macros, RTF, PDF, Txt, XML etc
                switch (filterIndex)
                {
                    case 1: //Save as DOCX
                        SaveDocument(currentDocument, Word.WdSaveFormat.wdFormatDocumentDefault, saveFileDialog.FileName);
                        break;

                    case 2: //Save as DOC
                        SaveDocument(currentDocument, Word.WdSaveFormat.wdFormatDocument97, saveFileDialog.FileName);
                        break;

                    default: //Save as DOCX
                        SaveDocument(currentDocument, Word.WdSaveFormat.wdFormatDocumentDefault, saveFileDialog.FileName);
                        break;
                }
            }

            //currentDocument.Close();

            Log.Verbose("SaveAsButton_Click() -- End");
        }

        public void SaveDocument(Word.Document document, Word.WdSaveFormat format, string fileName)
        {
            Log.Verbose("Save_Document() -- Begin");

            try
            {
                document.SaveAs(fileName, format);
            }
            catch (Exception exception)
            {
                Log.Error("Error while saving the file.");
                Log.Error(exception.Message);
            }

            Log.Verbose("Save_Document() -- End");
        }

        

        #region Add-in Express automatic code

        // Required by Add-in Express - do not modify
        // the methods within this region

        public override System.ComponentModel.IContainer GetContainer()
        {
            if (components == null)
                components = new System.ComponentModel.Container();
            return components;
        }

        [ComRegisterFunctionAttribute]
        public static void AddinRegister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXRegister(t);
        }

        [ComUnregisterFunctionAttribute]
        public static void AddinUnregister(Type t)
        {
            AddinExpress.MSO.ADXAddinModule.ADXUnregister(t);
        }

        public override void UninstallControls()
        {
            base.UninstallControls();
        }

        #endregion

        public static new AddinModule CurrentInstance
        {
            get
            {
                return AddinExpress.MSO.ADXAddinModule.CurrentInstance as AddinModule;
            }
        }

        public Word._Application WordApp
        {
            get
            {
                return (HostApplication as Word._Application);
            }
        }

        private void adxRibbonButtonSetting_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            Log.Debug("rbtnSettings_OnClick() -- Begin");
            var settingForm = new SettingsForm(); 
            this.LoadValuesFromRegistry();            
            if (settingForm.ShowDialog() == DialogResult.OK)
            {        
                
                if (SaveValuesToRegistry(this.ProLoopUrl, this.ProLoopUsername, this.ProLoopPassword))
                {
                    MessageBox.Show("Successfully saved the values.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    this.UnmapAndMapWebDavFolderToDrive();
                }
                else
                {
                    MessageBox.Show("Error while saving the values.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
            Log.Debug("rbtnSettings_OnClick() -- End");
        }
        public void UnmapAndMapWebDavFolderToDrive()
        {
            Log.Debug("UnmapAndMapWebDavFolderToDrive() -- Begin");
            this.UnMapWebDAVFolderToDriveLetter(false);
            this.MapWebDAVFolderToDriveLetter();
            Log.Debug("UnmapAndMapWebDavFolderToDrive() -- End");
        }

        private void adxRibbonButtonInfo_OnClick(object sender, IRibbonControl control, bool pressed)
        {
            new AboutUs().ShowDialog();
        }
    }
    
}

