using System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Windows.Forms;
using AddinExpress.MSO;
using Word = Microsoft.Office.Interop.Word;
using System.Collections.Generic;
using System.IO;
using System.Management;
using System.Linq;
using Serilog;
using Win32;
using ProLoop.Startup;
using Autofac;
using ProLoop.Dashboard;
using AddinExpress.WD;

namespace ProLoop
{
    /// <summary>
    ///   Add-in Express Add-in Module
    /// </summary>
    [GuidAttribute("B0F447C4-E459-4D33-86CB-50A4394F0418"), ProgId("ProLoop.AddinModule")]
    public partial class AddinModule : AddinExpress.MSO.ADXAddinModule
    {
        public string ProLoopUrl { get; set; }
        public string WebDAVUrl => this.ProLoopUrl + "/dav";
        public string WebDAVMappedDriveLetter { get; set; }
        private string WebDAVFolderMappedDriveLetter;
        public AddinModule()
        {
            Application.EnableVisualStyles();
            InitializeComponent();
            // Please add any initialization code to the AddinInitialize event handler
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

        private void AddinModule_AddinStartupComplete(object sender, EventArgs e)
        {
            var bootstrapper = new Bootstrapper();
            var container = bootstrapper.Bootstrap();
            container.Resolve<Controller.MainEntry>();
        }

        private void AddinModule_AddinInitialize(object sender, EventArgs e)
        {
           
        }
        public void ActivatePanel()
        {
            var loginItem = adxWordTaskPanesManager1.Items[0] as ADXWordTaskPanesCollectionItem;
            if (loginItem != null)
            {
                loginItem.CurrentTaskPaneInstance.Hide();
            }

            var currentItem = adxWordTaskPanesManager1.Items[1] as ADXWordTaskPanesCollectionItem;
            if (currentItem != null)
            {
                currentItem.CurrentTaskPaneInstance.Visible = true ;
            }
        }
        //private void UnMapWebDAVFolderToDriveLetter(bool forceUnmap)
        //{
        //    Log.Verbose("UnMapWebDAVFolderToDriveLetter() -- Begin");

        //    //Disconnect the mapped WebDAV drive
        //    if (!string.IsNullOrEmpty(WebDAVFolderMappedDriveLetter))
        //    {
        //        var result = Utility.WNetCancelConnection2(WebDAVFolderMappedDriveLetter,
        //            Utility.CONNECT_UPDATE_PROFILE, forceUnmap);

        //        if (result == 0) //Success
        //        {
        //            Log.Verbose("Successfully unmapped the WebDAV folder from drive letter {0}",
        //                WebDAVFolderMappedDriveLetter);
        //            WebDAVFolderMappedDriveLetter = string.Empty;
        //        }
        //        else
        //        {
        //            Log.Verbose("Unable to unmap the WebDAV folder. Please verify the network connectivity to WebDAV server");
        //        }
        //    }

        //    Log.Verbose("UnMapWebDAVFolderToDriveLetter() -- End");
        //}
        //public void MapWebDAVFolderToDriveLetter()
        //{
        //    //Log.Debug("MapWebDAVFolderToDriveLetter() -- Begin");

        //    //Verify if a drive letter is unmapped and available for mapping
        //    var driveLetters = new List<string>
        //    {
        //        @"W:\",
        //        @"Z:\",
        //        @"Y:\",
        //        @"X:\",
        //        @"V:\",
        //        @"U:\",
        //        @"T:\",
        //        @"S:\",
        //        @"R:\",
        //        @"Q:\",
        //        @"P:\",
        //        @"O:\",
        //        @"N:\",
        //        @"M:\",
        //        @"L:\",
        //        @"K:\",
        //        @"J:\",
        //        @"I:\",
        //        @"H:\",
        //        @"G:\",
        //        @"F:\",
        //        @"E:\",
        //        @"D:\",
        //        @"B:\"
        //    };

        //    // Remove the drive letters which are already being used, from the available drive letter list
        //    var usedDrives = DriveInfo.GetDrives();
        //    foreach (var usedDrive in usedDrives)
        //    {
        //        var usedDriveLetter = usedDrive.Name; // C:\, E:\, etc:\
        //        var usedDriveType = usedDrive.DriveType;
        //       // Log.Debug("Drive letter in use: {0}, Drive type: {1}", usedDriveLetter, usedDriveType);

        //        //Remove this drive letter from the unused drive letter list
        //        driveLetters.Remove(usedDriveLetter);
        //    }

        //    //Verify if any used drive letter is already mapped to the WebDAV folder
        //    try
        //    {
        //        var searcher = new ManagementObjectSearcher("select * from Win32_MappedLogicalDisk");
        //        foreach (var resultItem in searcher.Get())
        //        {
        //            var drive = (ManagementObject)resultItem;
        //           // Log.Debug("Drive letter {0} is mapped to {1}", drive["Name"].ToString(),
        //               // drive["ProviderName"].ToString());

        //            var mappedDrivePath = drive["ProviderName"].ToString();
        //            var mappedFolderPathAsUnc = GetUncFromUrl(WebDAVUrl);

        //            #region Hide
        //            ////if (mappedDrivePath.ToUpper().StartsWith(@"\\NEWUI.PROLOOP.COM\DAVWWWROOT"))



        //            #endregion

        //            if (mappedDrivePath.ToUpper().StartsWith(mappedFolderPathAsUnc.ToUpper()))
        //            {
        //                //A drive letter is already mapped to WebDAV folder
        //                WebDAVMappedDriveLetter = drive["Name"].ToString();
        //                Log.Information("WebDAV folder {0} is already mapped to drive {1}",
        //                    mappedFolderPathAsUnc, WebDAVMappedDriveLetter);
        //                return;
        //            }
        //            //Console.WriteLine(Regex.Match(drive["ProviderName"].ToString(), @"\\\\([^\\]+)").Groups[1]);
        //        }
        //    }
        //    catch (Exception exception)
        //    {              

        //        return;
        //    }
        //    if (string.IsNullOrEmpty(ProLoopUrl))
        //        return;           
        //    WebDAVMappedDriveLetter = driveLetters.First().Replace("\\", "");

        //    var networkResource = new Utility.NETRESOURCE
        //    {
        //        dwType = Utility.RESOURCETYPE_DISK,
        //        lpLocalName = WebDAVMappedDriveLetter,
        //        lpRemoteName = GetUncFromUrl(WebDAVUrl), //WebDAVUrl,
        //        lpProvider = null
        //    };



        //    var result = Utility.WNetAddConnection2(ref networkResource, ProLoopPassword, ProLoopUsername, 0);


        //    if (result == 0) // No Error
        //    {
        //        Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}", WebDAVMappedDriveLetter);
        //    }
        //    else if (result == 1244) //ERROR_NOT_AUTHENTICATED 
        //    {
        //        Log.Verbose(
        //            "Unable to connect to the ProLoop server. Please verify the username and/or password.");
        //        MessageBox.Show(
        //            "Unable to connect to the ProLoop server. Please verify the username and/or password.",
        //            "Authentication error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    else if (result == 53)
        //    {
        //        Log.Verbose(
        //            "Unable to connect to the ProLoop server. Please verify the ProLoop URL in Settings.");
        //        MessageBox.Show(
        //            "Unable to connect to the ProLoop server. Please verify the ProLoop URL in Settings.",
        //            "Settings error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    else if (result == 67) //ERROR_BAD_NET_NAME -- No network situation
        //    {
        //        Log.Verbose(
        //            "Unable to connect to the ProLoop server. Please verify the network connectivity.");
        //        MessageBox.Show(
        //            "Unable to connect to the ProLoop server. Please verify the network connectivity.",
        //            "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    else if (result == 487) //ERROR_INVALID_ADDRESS -- No URL or invalid address
        //    {
        //        Log.Verbose("Unable to connect to the ProLoop server. Please verify the URL.");
        //        MessageBox.Show(
        //            "Unable to connect to the ProLoop server. Please verify the URL.",
        //            "Invalid Address", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    else if (result == 85) //ERROR_ALREADY_ASSIGNED -- Drive is already mapped
        //    {
        //        //Unmap the existing drive & remap
        //        Log.Verbose(
        //            "The drive letter is already assigned to another network resource. \nWill attempt to unmap and remap");
        //        UnMapWebDAVFolderToDriveLetter(true);

        //        var resultNew = Utility.WNetAddConnection2(ref networkResource, null, null, 0);
        //        if (resultNew == 0)
        //        {
        //            WebDAVMappedDriveLetter = driveLetters.First().Replace("\\", "");
        //            Log.Verbose("Successfully mapped the WebDAV folder to drive letter {0}",
        //                WebDAVMappedDriveLetter);
        //        }
        //        else
        //        {
        //            Log.Verbose(
        //                "Unable to open the WebDAV folder. \nPlease verify the access permissions to WebDAV folder");
        //            MessageBox.Show(
        //                "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
        //                "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //        }
        //    }
        //    else
        //    {
        //        Log.Verbose(
        //            "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.");
        //        MessageBox.Show(
        //            "Unable to connect to the ProLoop WebDAV folder. \nPlease verify the network connectivity.",
        //            "Network error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        //    }
        //    //MessageBox.Show(result.ToString());

        //   // Log.Verbose("MapWebDAVFolderToDriveLetter() -- End");
        //}
        //private string GetUncFromUrl(string url)
        //{
        //    //  Log.Debug("GetUncFromUrl() -- Begin");          
        //    string uncURl = string.Empty;
        //    if (ProLoopUrl.ToUpper().Contains("HTTPS://"))
        //    {
        //        ProLoopUrl = ProLoopUrl.ToUpper().Replace("HTTPS://", "");
        //        uncURl = "@SSL";
        //    }
        //    else
        //    {
        //        if (ProLoopUrl.ToUpper().Contains("HTTP://"))
        //        {
        //            ProLoopUrl = ProLoopUrl.ToUpper().Replace("HTTP://", "");
        //        }
        //    }
        //    if (ProLoopUrl.EndsWith("/"))
        //    {
        //        ProLoopUrl = ProLoopUrl.Substring(0, ProLoopUrl.Length - 1);
        //    }
        //    string result = "\\\\" + ProLoopUrl + uncURl + "\\DavWWWRoot\\dav";
        //    //Log.Debug("GetUncFromUrl() -- End");
        //    return result;
        //}
    }
}

