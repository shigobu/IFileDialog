using System;
using System.Runtime.InteropServices;

namespace COMInterfaceWrapper.Native
{
    internal static class NativeMethods
    {
        [DllImport("shell32.dll", ExactSpelling = true, CharSet = CharSet.Unicode, PreserveSig = false)]
        public static extern IShellItem SHCreateItemFromParsingName(
            string pszPath,
            IntPtr pbc,
            [MarshalAs(UnmanagedType.LPStruct)]Guid riid);

        /// <summary>
        /// win32API。Windowsが管理している全てのウィンドウのうち、ユーザの操作対象になっているフォアグラウンドウィンドウのハンドルを取得する。
        /// </summary>
        /// <returns></returns>
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();

        public static int ERROR_CANCELLED { get; } = BitConverter.ToInt32(BitConverter.GetBytes(0x800704C7), 0);

        public static int S_OK { get; } = 0;
    }
}
