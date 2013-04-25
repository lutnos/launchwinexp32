using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace LaunchWinExp32
{
    [ComImport]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("50A87BAA-5F79-4C31-B6B3-28F6F2D097E6")]
    public interface IExplorerHost
    {
        void ShowWindow(IntPtr pidlItem,
                        [MarshalAs(UnmanagedType.I4)] Int32 flags,
                        [MarshalAs(UnmanagedType.I4)] Int32 x,
                        [MarshalAs(UnmanagedType.I4)] Int32 y,
                        [MarshalAs(UnmanagedType.I4)] Int32 nCmdShow);
    }

    static class Program
    {
        public static readonly int SW_SHOWNORMAL = 1;
        public static readonly uint CLSCTX_LOCAL_SERVE = 4;
        public static readonly uint CLSCTX_ACTIVATE_32_BIT_SERVER = 0x40000;

        public static readonly Guid CLSID_SeparateMultiProcessExplorerHost = new Guid("{75dff2b7-6936-4c06-a8bb-676a7b00b24b}");

        public static readonly Guid IID_IUnknown = new Guid("00000000-0000-0000-C000-000000000046");

        public static readonly string COMPUTER = "::{20d04fe0-3aea-1069-a2d8-08002b30309d}";

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr ILCreateFromPath([In, MarshalAs(UnmanagedType.LPWStr)] string pszPath);

        [DllImport("shell32.dll")]
        public static extern void ILFree([In] IntPtr pidl);

        [DllImport("ole32.dll", EntryPoint = "CoCreateInstance", CallingConvention = CallingConvention.StdCall)]
        public static extern UInt32 CoCreateInstance([In, MarshalAs(UnmanagedType.LPStruct)] Guid clsid,
                                                     [In, MarshalAs(UnmanagedType.IUnknown)] object inner,
                                                     [In, MarshalAs(UnmanagedType.I4)] UInt32 context,
                                                     [In, MarshalAs(UnmanagedType.LPStruct)] Guid uuid,
                                                     [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object rReturnedComObject);


        /// <summary>
        /// 
        /// The main entry point for the application.
        /// 
        /// Launches a 32-bit Windows Explorer on x64 Windows 7 or Windows 8,
        /// where otherwise a %windir%\SYSWOW64.explorer.exe will always
        /// launch a 64-bit %windir\explorer.exe . Run with no arguments to 
        /// explore Computer, run with one to explore that.
        /// 
        /// Dave E English  lutnos.com  19 March 2013
        /// 
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 1)
            {
                MessageBox.Show("Usage: LaunchExp32 [object], e.g.: LaunchExp32 D:");
            }
            else if (Environment.OSVersion.Version.Major < 6 ||
                     (Environment.OSVersion.Version.Major == 6 &&
                       Environment.OSVersion.Version.Minor < 1))
            {
                MessageBox.Show("This program requires Windows 7 or later");
            }
            else
            {
                object explorerHost = null;

                uint result = CoCreateInstance(CLSID_SeparateMultiProcessExplorerHost, null,
                               CLSCTX_LOCAL_SERVE | CLSCTX_ACTIVATE_32_BIT_SERVER, IID_IUnknown, ref explorerHost);

                IntPtr pidl = ILCreateFromPath((args.Length < 1) ? COMPUTER : args[0]);

                ((IExplorerHost)explorerHost).ShowWindow(pidl, 0, 0, 0, SW_SHOWNORMAL);

                ILFree(pidl);
            }
        }
    }
}
