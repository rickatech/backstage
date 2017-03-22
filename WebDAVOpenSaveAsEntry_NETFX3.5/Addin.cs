#define NETFX35
#define AUTH_HARDCODED

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Management;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using NetOffice.Tools;
using NetOffice.WordApi.Enums;
using NetOffice.WordApi.Tools;
using stdole;
using WebDAVSaveAsEntry.Properties;
using Win32;
using Word = NetOffice.WordApi;
using Office = NetOffice.OfficeApi;
using VBIDE = NetOffice.VBIDEApi;
#if NETFX45
using System.Threading.Tasks;
using Serilog;
using WebDAVSaveAsEntry.Properties;
#endif
#if NETFX35
using System.Reflection;

#endif

namespace WebDAVSaveAsEntry
{
    [COMAddin("WebDAVSaveAsEntry", "Add WebDAV Folder to Save As Pane", 3)]
    [ProgId("WebDAVSaveAsEntry.Addin")]
    [Guid("DC5DDA46-10EA-4CFF-B31A-AD6A86DA57CE")]
    [RegistryLocation(RegistrySaveLocation.LocalMachine)]
    [CustomUI("WebDAVSaveAsEntry.RibbonUI.xml")]
    public class Addin : COMAddin
    {
        private string WebDAVFolderMappedDriveLetter;
        private string WebDAVFolderURL;

        public Addin()
        {
            OnConnection += Addin_OnConnection;
            OnStartupComplete += Addin_OnStartupComplete;
            //OnDisconnection += Addin_OnDisconnection;
        }

        internal Office.IRibbonUI RibbonUI { get; private set; }

        private void Addin_OnConnection(object application, ext_ConnectMode connectmode, object addininst,
            ref Array custom)
        {
            Application.QuitEvent += Application_OnQuit;
        }

        private void Application_OnQuit()
        {
            UnMapWebDAVFolderToDriveLetter(true);
        }

        private void UnMapWebDAVFolderToDriveLetter(bool forceUnmap)
        {
            //Disconnect the mapped WebDAV drive
            if (!string.IsNullOrEmpty(WebDAVFolderMappedDriveLetter))
            {
                var result = Utility.WNetCancelConnection2(WebDAVFolderMappedDriveLetter, Utility.CONNECT_UPDATE_PROFILE,
                    forceUnmap);

                if (result == 0) //Success
                    WebDAVFolderMappedDriveLetter = string.Empty;
            }
        }

        private void Addin_OnStartupComplete(ref Array custom)
        {
            WebDAVFolderMappedDriveLetter = string.Empty;
            WebDAVFolderURL = @"http://newui.proloop.com/dav/";

            //Map the WebDAV folder to a drive letter
#if NETFX45
            Task.Factory.StartNew(MapWebDAVFolderToDriveLetter);
#elif NETFX35
            ThreadPool.QueueUserWorkItem(obj => { MapWebDAVFolderToDriveLetter(); });
#endif
        }

        private void MapWebDAVFolderToDriveLetter()
        {
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

            //Remote the drive letters which are already being used
            var usedDrives = DriveInfo.GetDrives();
            foreach (var usedDrive in usedDrives)
            {
                var driveName = usedDrive.Name; // C:\, E:\, etc:\
                var driveType = usedDrive.DriveType;

                //Remove this drive letter from the unused drive letter list
                driveLetters.Remove(driveName);
            }

            //Verify if any used drive letter is already mapped to the WebDAV folder
            try
            {
                var searcher = new ManagementObjectSearcher("select * from Win32_MappedLogicalDisk");
                foreach (ManagementObject drive in searcher.Get())
                {
                    var mappedDrivePath = drive["ProviderName"].ToString();
                    if (mappedDrivePath.ToUpper().StartsWith(@"\\NEWUI.PROLOOP.COM\DAVWWWROOT"))
                    {
                        //A drive letter is already mapped to WebDAV folder
                        WebDAVFolderMappedDriveLetter = drive["Name"].ToString();
                        return;
                    }
                    //Console.WriteLine(Regex.Match(drive["ProviderName"].ToString(), @"\\\\([^\\]+)").Groups[1]);
                }
            }
            catch (Exception exception)
            {
            }

            //Map the WebDAV folder to the drive letter -- WebDAVFolderMappedDriveLetter
            //Get the first unused drive letter
            WebDAVFolderMappedDriveLetter = driveLetters.First().Replace("\\", "");

            var networkResource = new Utility.NETRESOURCE();
            networkResource.dwType = Utility.RESOURCETYPE_DISK;
            networkResource.lpLocalName = WebDAVFolderMappedDriveLetter;
            networkResource.lpRemoteName = WebDAVFolderURL;
            networkResource.lpProvider = null;

#if AUTH_HARDCODED
            //Authenticate using username and password
            string usernameProLoopWebDAV = "sn";
            string passwordProLoopWebDAV = "p@ssw0rd";
            var result = Utility.WNetAddConnection2(ref networkResource, passwordProLoopWebDAV, usernameProLoopWebDAV, 0);            //var result = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
#else
            var result = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
#endif

            //TODO: Task 3: Handle exceptions when WebDAV folder is not accessible (access denied, permission errors, user account disabled, etc)
            if (result == 0)
            {
            }
            if (result == 1244) //ERROR_NOT_AUTHENTICATED 
            {
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. Please verify the username and/or password.",
                    "Authentication error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 67) //ERROR_BAD_NET_NAME -- No network situation
            {
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. Please verify the network connectivity",
                    "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (result == 85) //ERROR_ALREADY_ASSIGNED -- Drive is already mapped
            {
                //Unmap the existing drive & remap
                UnMapWebDAVFolderToDriveLetter(true);

                var resultNew = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
                if (resultNew == 0)
                {
                    WebDAVFolderMappedDriveLetter = driveLetters.First().Replace("\\", "");
                    ;
                }
                else
                {
                    MessageBox.Show(
                        "Unable to connect to the ProLoop WebDAV folder. Please verify the network connectivity.",
                        "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show(
                    "Unable to connect to the ProLoop WebDAV folder. Please verify the network connectivity.",
                    "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            //MessageBox.Show(result.ToString());
        }

        //private void Addin_OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        //{

        //}

        public void OpenButton_Click(Office.IRibbonControl control)
        {
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
                    Application.Documents.Open(documentToBeOpened);
                }
                catch (Exception exception)
                {
                }
            }
        }

        public void SaveAsButton_Click(Office.IRibbonControl control)
        {
            //MessageBox.Show(String.Format("WebDAV Save As Entry Version {0}", this.GetType().Assembly.GetName().Version),
            //	"About WebDAV Save As Entry", MessageBoxButtons.OK, MessageBoxIcon.Information);

            var currentDocument = Application.ActiveDocument;

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


                //TODO: Add support for other file types (with Macros, RTF, PDF, Txt, XML etc
                switch (filterIndex)
                {
                    case 1: //Save as DOCX
                        SaveDocument(currentDocument, WdSaveFormat.wdFormatDocumentDefault, saveFileDialog.FileName);
                        break;

                    case 2: //Save as DOC
                        SaveDocument(currentDocument, WdSaveFormat.wdFormatDocument97, saveFileDialog.FileName);
                        break;

                    default: //Save as DOCX
                        SaveDocument(currentDocument, WdSaveFormat.wdFormatDocumentDefault, saveFileDialog.FileName);
                        break;
                }
            }
        }

        public void SaveDocument(Word.Document document, WdSaveFormat format, string fileName)
        {
            try
            {
                document.SaveAs(fileName, format);
            }
            catch (Exception exception)
            {
            }
        }

        public void OnLoadRibonUI(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;
        }

        public IPictureDisp GetButtonImage(Office.IRibbonControl control)
        {
            IPictureDisp tempIPictureDisp = null;

            try
            {
                //var logoBitmap = (Bitmap) Image.FromFile(@"Images\proloop.bmp");
                //var logoBitmap = Resources.proloop;

                var thisExe = Assembly.GetExecutingAssembly();
                //string[] resourceNames = Assembly.GetExecutingAssembly().GetManifestResourceNames();
                //Stream file = WebDAVSaveAsEntry.Properties.Resources.proloop;
                //Stream file = thisExe.GetManifestResourceStream("WebDAVSaveAsEntry.Properties.Resources.proloop");
                var file = Resources.ProLoopLogo;
                //var logoBitmap = Image.FromStream(file);

                tempIPictureDisp = ImageConverter.Convert(file);
            }
            catch (Exception exception)
            {
#if DEBUG
                MessageBox.Show(exception.Message);
#endif
            }

            return tempIPictureDisp;
            //var stdoleIPictureDisp = OleCreateConverter.ImageToPictureDisp(logoBitmap);
            //Log.Verbose("stdoleIPictureDisp Height: {0}", stdoleIPictureDisp.Height);
            //Log.Verbose("stdoleIPictureDisp Width: {0}", stdoleIPictureDisp.Width);

            //Log.Verbose("GetButtonImage() -- End (before return)");
            //return stdoleIPictureDisp;
        }

        protected override void OnError(ErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind, "WebDAV Save As Entry");
        }

        [RegisterErrorHandler]
        public static void RegisterErrorHandler(RegisterErrorMethodKind methodKind, Exception exception)
        {
            MessageBox.Show("An error occurend in " + methodKind, "WebDAV Save As Entry");
        }
    }

    internal class ImageConverter : AxHost
    {
        public ImageConverter() : base(null)
        {
        }

        public static IPictureDisp Convert(Image image)
        {
            return (IPictureDisp) GetIPictureDispFromPicture(image);
        }
    }
}