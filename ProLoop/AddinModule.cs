#define NETFX45

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Windows.Forms;
//using stdole;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Security.AccessControl;
using AddinExpress.MSO;
using Microsoft.Win32;
using ProLoop.WordAddin.Properties;
using Win32;
using Word = Microsoft.Office.Interop.Word;

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
        private string WebDAVFolderURL;
        private string WebDAVUsername;
        private string WebDAVPassword;

        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();

            AddinInitialize += OnAddinInitialize;
            AddinStartupComplete += OnAddinStartupComplete;
        }

        private void OnAddinStartupComplete(object sender, EventArgs eventArgs)
        {
            Log.Verbose("Addin_OnStartupComplete() -- Begin");
            Log.Verbose("Addin started in Word Version {0}", this.HostVersion);

            WebDAVFolderMappedDriveLetter = string.Empty;

            // Hard coded path
            //WebDAVFolderURL = @"http://newui.proloop.com/dav/";

            // Path values stored in settings file (app.config)
            //LoadValuesFromSettingsFile(out WebDAVFolderURL, out WebDAVUsername, out WebDAVPassword);

            // Path values stored in the Registry (HKCU\Software\ProLoop\)
            LoadValuesFromRegistry(out WebDAVFolderURL, out WebDAVUsername, out WebDAVPassword);

            //Map the WebDAV folder to a drive letter
#if NETFX45
            Task.Factory.StartNew(MapWebDAVFolderToDriveLetter);
#elif NETFX35
            ThreadPool.QueueUserWorkItem(obj => { MapWebDAVFolderToDriveLetter(); });
#endif

            Log.Verbose("Addin_OnStartupComplete() -- End");
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
                Log.Verbose("Exception during Addin_OnConnection(): {0}", exception.Message);
#if DEBUG
                MessageBox.Show("Exception during Addin_OnConnection: " + exception.Message);
                MessageBox.Show("Inner Exception: " + exception.InnerException);
#endif
            }

            Log.Verbose("Addin_OnConnection() -- Begin (after Logging setup)");

            AddinFinalize += OnAddinFinalize;

            WordApplication = this.WordApp.Application;

            Log.Verbose("Addin_OnConnection() -- End");
        }

        private void OnAddinFinalize(object sender, EventArgs eventArgs)
        {
            Log.Verbose("Application_OnQuit()");

            UnMapWebDAVFolderToDriveLetter(true);

            //Stop logging
            if (Log.Logger != null)
                Logger.CloseLogging();
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

        private void LoadValuesFromRegistry(out string webDavFolderUrl, out string webDavUsername, out string webDavPassword)
        {
            webDavFolderUrl = null;
            webDavUsername = null;
            webDavPassword = null;

            try
            {
                using (RegistryKey regCurrentUserRootKey = Registry.CurrentUser)
                {
                    var regProLoopSubKey =
                        regCurrentUserRootKey.OpenSubKey(
                            @"Software\ProLoop",
                            RegistryKeyPermissionCheck.ReadSubTree,
                            RegistryRights.ReadKey);

                    var urlValue = regProLoopSubKey?.GetValue("URL");
                    if (urlValue != null)
                    {
                        webDavFolderUrl = urlValue.ToString();
                        Log.Verbose("Registry URL value: {0}", webDavFolderUrl);
                    }

                    var usernameValue = regProLoopSubKey?.GetValue("Username");
                    if (usernameValue != null)
                    {
                        webDavUsername = usernameValue.ToString();
                        Log.Verbose("Registry Username value: {0}", webDavUsername);
                    }

                    var passwordValue = regProLoopSubKey?.GetValue("Password");
                    if (passwordValue != null)
                    {
                        webDavPassword = passwordValue.ToString();
                        Log.Verbose("Registry Password value: {0}", webDavPassword);
                    }
                }
            }
            catch (Exception exception)
            {
                Log.Verbose("Exception while loading the WebDAV settings from registry.");
                Log.Verbose(exception.Message);
            }
        }

        private void MapWebDAVFolderToDriveLetter()
        {
            Log.Verbose("MapWebDAVFolderToDriveLetter() -- Begin");

            //Verify if a driver letter is unmapped and available for mapping

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
                var driveName = usedDrive.Name; // C:\, E:\, etc:\
                var driveType = usedDrive.DriveType;
                Log.Verbose("Used drive letter: {0}, Type: {1}", driveName, driveType);

                //Remove this drive letter from the unused drive letter list
                driveLetters.Remove(driveName);
            }

            //Verify if any used drive letter is already mapped to the WebDAV folder
            try
            {
                var searcher = new ManagementObjectSearcher("select * from Win32_MappedLogicalDisk");
                foreach (var resultItem in searcher.Get())
                {
                    var drive = (ManagementObject)resultItem;
                    Log.Verbose("Drive letter {0} is mapped to {1}", drive["Name"].ToString(), drive["ProviderName"].ToString());

                    var mappedDrivePath = drive["ProviderName"].ToString();

                    //if (mappedDrivePath.ToUpper().StartsWith(@"\\NEWUI.PROLOOP.COM\DAVWWWROOT"))

                    // Conver the URL to UNC: 
                    // From http://newui.proloop.com/dav/ to \\NEWUI.PROLOOP.COM\DAV
                    string mappedFolderPathAsUnc = WebDAVFolderURL.ToUpper()
                        .Replace(@"HTTP://", @"\\")
                        .Replace(@"/", @"\");
                    if (mappedFolderPathAsUnc.EndsWith(@"\")) // Remove the trailing \
                    {
                        mappedFolderPathAsUnc = mappedDrivePath.Substring(0, mappedFolderPathAsUnc.Length - 2);
                    }

                    if (mappedDrivePath.ToUpper().StartsWith(mappedFolderPathAsUnc))
                    {
                        //A drive letter is already mapped to WebDAV folder
                        WebDAVFolderMappedDriveLetter = drive["Name"].ToString();
                        //Log.Information("WebDAV folder {0} is already mapped to drive {1}", 
                        //  @"\\NEWUI.PROLOOP.COM\DAVWWWROOT", WebDAVFolderMappedDriveLetter);
                        Log.Information("WebDAV folder {0} is already mapped to drive {1}",
                            mappedFolderPathAsUnc, WebDAVFolderMappedDriveLetter);

                        return;
                    }
                    //Console.WriteLine(Regex.Match(drive["ProviderName"].ToString(), @"\\\\([^\\]+)").Groups[1]);
                }
            }
            catch (Exception exception)
            {
                Log.Error("Exception while iterating over drive letters and their mapped path.");
                Log.Error(exception.Message);
            }

            //Map the WebDAV folder to the drive letter -- WebDAVFolderMappedDriveLetter
            //Get the first unused drive letter
            WebDAVFolderMappedDriveLetter = driveLetters.First().Replace("\\", "");

            var networkResource = new Utility.NETRESOURCE
            {
                dwType = Utility.RESOURCETYPE_DISK,
                lpLocalName = WebDAVFolderMappedDriveLetter,
                lpRemoteName = WebDAVFolderURL,
                lpProvider = null
            };

            // Hard coded: Authenticate using username and password
            //string usernameProLoopWebDAV = "sn";
            //string passwordProLoopWebDAV = "p@ssw0rd";
            //var result = Utility.WNetAddConnection2(ref networkResource, passwordProLoopWebDAV, usernameProLoopWebDAV, 0);
            //var result = Utility.WNetAddConnection2(ref networkResource, null, null, 0);

            var result = Utility.WNetAddConnection2(ref networkResource, WebDAVPassword, WebDAVUsername, 0);

            // TODO: Task 3: Handle exceptions when WebDAV folder is not accessible 
            // (access denied, permission errors, user account disabled, etc)
            if (result == 0)  // No Error
            {
                Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}", WebDAVFolderMappedDriveLetter);
            }
            else if (result == 1244) //ERROR_NOT_AUTHENTICATED 
            {
                Log.Verbose("Unable to connect to the ProLoop WebDAV folder. \nPlease verify the username and/or password.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the username and/or password.",
                    "Authentication error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 67) //ERROR_BAD_NET_NAME -- No network situation
            {
                Log.Verbose("Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
                    "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 487) //ERROR_INVALID_ADDRESS -- No URL or invalid address
            {
                Log.Verbose("Unable to connect to the ProLoop WebDAV folder. \nPlease verify the URL.");
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the URL.",
                    "Invalid Address", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 85) //ERROR_ALREADY_ASSIGNED -- Drive is already mapped
            {
                //Unmap the existing drive & remap
                Log.Verbose("The drive letter is already assigned to another network resource. \nWill attempt to unmap and remap");
                UnMapWebDAVFolderToDriveLetter(true);

                var resultNew = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
                if (resultNew == 0)
                {
                    WebDAVFolderMappedDriveLetter = driveLetters.First().Replace("\\", "");
                    Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}", WebDAVFolderMappedDriveLetter);
                }
                else
                {
                    Log.Verbose("Unable to open the WebDAV folder. \nPlease verify the access permissions to WebDAV folder");
                    MessageBox.Show(
                        "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
                        "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                Log.Verbose("Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.");
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

        //        public IPictureDisp GetButtonImage(Office.IRibbonControl control)
        //        {
        //            IPictureDisp tempIPictureDisp = null;

        //            Log.Verbose("GetButtonImage() -- Begin");
        //            try
        //            {
        //                //var logoBitmap = (Bitmap) Image.FromFile(@"Images\proloop.bmp");
        //#if NETFX45

        //                var logoBitmap = Resources.ProLoopLogo;
        //                Log.Verbose("Bitmap Height: {0}", logoBitmap.PhysicalDimension.Height);
        //                Log.Verbose("Bitmap Width: {0}", logoBitmap.PhysicalDimension.Width);

        //                tempIPictureDisp = ImageConverter.Convert(logoBitmap);
        //#elif NETFX35
        //                //var logoBitmap = Resources.proloop;

        //                var thisExe = Assembly.GetExecutingAssembly();
        //                //string[] resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
        //                //Stream file = WebDAVSaveAsEntry.Properties.Resources.proloop;
        //                //Stream file = thisExe.GetManifestResourceStream("WebDAVSaveAsEntry.Properties.Resources.proloop");
        //                var file = Resources.ProLoopLogo;
        //                //var logoBitmap = Image.FromStream(file);

        //                tempIPictureDisp = ImageConverter.Convert(file);
        //#endif				


        //            }
        //            catch (Exception exception)
        //            {
        //                Log.Error(exception.Message);
        //#if DEBUG
        //                MessageBox.Show(exception.Message);
        //#endif
        //            }

        //            return tempIPictureDisp;
        //            //var stdoleIPictureDisp = OleCreateConverter.ImageToPictureDisp(logoBitmap);
        //            //Log.Verbose("stdoleIPictureDisp Height: {0}", stdoleIPictureDisp.Height);
        //            //Log.Verbose("stdoleIPictureDisp Width: {0}", stdoleIPictureDisp.Width);

        //            //Log.Verbose("GetButtonImage() -- End (before return)");
        //            //return stdoleIPictureDisp;
        //        }

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
    }

    internal class ImageConverter : AxHost
    {
        public ImageConverter() : base(null)
        {
        }

        public static IPictureDisp Convert(Image image)
        {
            return (IPictureDisp)GetIPictureDispFromPicture(image);
        }
    }
}

