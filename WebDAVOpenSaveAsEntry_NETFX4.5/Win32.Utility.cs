using System.Runtime.InteropServices;

namespace Win32
{
    public static class Utility
    {
        //NETRESOURCE.dwType
        public const uint RESOURCETYPE_DISK = 1;

        public const uint CONNECT_UPDATE_PROFILE = 0x1;
        public const uint CONNECT_INTERACTIVE = 0x8;
        public const uint CONNECT_PROMPT = 0x10;
        public const uint CONNECT_REDIRECT = 0x80;
        public const uint CONNECT_COMMANDLINE = 0x800;
        public const uint CONNECT_CMD_SAVECRED = 0x1000;

        [DllImport("mpr.dll")]
        public static extern uint WNetAddConnection2(
            ref NETRESOURCE lpNetResource, string lpPassword, string lpUsername, uint dwFlags);

        [DllImport("mpr.dll")]
        public static extern uint WNetCancelConnection2(string lpName, uint dwFlags, bool bForce);

        [StructLayout(LayoutKind.Sequential)]
        public struct NETRESOURCE
        {
            public uint dwScope;
            public uint dwType;
            public uint dwDisplayType;
            public uint dwUsage;
            public string lpLocalName;
            public string lpRemoteName;
            public string lpComment;
            public string lpProvider;
        }
    }
}