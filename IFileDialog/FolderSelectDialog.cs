using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using COMInterfaceWrapper.Native;
using System.Diagnostics;

namespace COMInterfaceWrapper
{
    public class FolderSelectDialog
    {
        public string Path { get; set; }
        public string Title { get; set; }

        /// <summary>
        /// フォアグラウンドウィンドウを親にして、ダイアログを表示します。
        /// </summary>
        /// <returns>
        /// 押されたボタンを示します。
        /// true  = OKボタン
        /// false = キャンセル または バツボタン
        /// </returns>
        public bool ShowDialog()
        {
            return ShowDialog(IntPtr.Zero);
        }

        /// <summary>
        /// 親ウィンドウを指定して、ダイアログを表示します。
        /// ownerにIntPtr.Zeroを指定した場合、フォアグラウンドウィンドウを親にします。
        /// </summary>
        /// <param name="hwndOwner">親ウィンドウのウィンドウハンドル</param>
        /// <returns>
        /// 押されたボタンを示します。
        /// true  = OKボタン
        /// false = キャンセル または バツボタン
        /// </returns>
        public bool ShowDialog(IntPtr hwndOwner)
        {
            var dlg = new FileOpenDialog() as IFileOpenDialog;
            try
            {
                if (hwndOwner == IntPtr.Zero)
                {
                    hwndOwner = NativeMethods.GetForegroundWindow();
                }

                FILEOPENDIALOGOPTIONS option = dlg.GetOptions();

                dlg.SetOptions(option | FILEOPENDIALOGOPTIONS.FOS_PICKFOLDERS);

                IShellItem item;
                if (!string.IsNullOrEmpty(this.Path))
                {
                    item = NativeMethods.SHCreateItemFromParsingName(Path, IntPtr.Zero, typeof(IShellItem).GUID);

                    dlg.SetFolder(item);
                }

                if (!string.IsNullOrEmpty(this.Title))
                    dlg.SetTitle(this.Title);

                var hr = dlg.Show(hwndOwner);
                if (hr.Equals(NativeMethods.ERROR_CANCELLED))
                {
                    return false;
                }
                if (!hr.Equals(0))
                {
                    return false;
                }

                dlg.GetResult(out item);
                try
                {
                    item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out string name);
                    this.Path = name;
                }
                catch (ArgumentException ex) when (ex.HResult == -2147024809)
                {
                    Path = "";
                    throw new InvalidOperationException("「PC」「ネットワーク」は指定できません。", ex);
                }

                return true;
            }
            finally
            {
                Marshal.FinalReleaseComObject(dlg);
            }
        }
    }
}
