using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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

        public const uint ERROR_CANCELLED = 0x800704C7;
    }
}
