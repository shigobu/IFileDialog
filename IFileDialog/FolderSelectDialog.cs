using System;
using System.Runtime.InteropServices;
using COMInterfaceWrapper.Native;

namespace COMInterfaceWrapper
{
    public class FolderSelectDialog
    {
        /// <summary>
        /// エクスプローラにPCを表示するための文字列
        /// </summary>
        public const string PcPath = "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}";

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
        /// hwndOwnerにIntPtr.Zeroを指定した場合、フォアグラウンドウィンドウを親にします。
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

                //ファイル選択のオプションと、ファイルシステムのアイテムであることを確認するオプション。PCとネットワークが選択できない。
                dlg.SetOptions(option | FILEOPENDIALOGOPTIONS.FOS_PICKFOLDERS | FILEOPENDIALOGOPTIONS.FOS_FORCEFILESYSTEM);

                IShellItem item;
                if (!string.IsNullOrEmpty(this.Path))
                {
                    item = NativeMethods.SHCreateItemFromParsingName(Path, IntPtr.Zero, typeof(IShellItem).GUID);

                    dlg.SetFolder(item);
                }

                if (!string.IsNullOrEmpty(this.Title))
                {
                    dlg.SetTitle(this.Title);
                }

                int hr = dlg.Show(hwndOwner);
                if (hr == NativeMethods.ERROR_CANCELLED)
                {
                    return false;
                }
                if (hr != NativeMethods.S_OK)
                {
                    Marshal.ThrowExceptionForHR(hr);
                }

                dlg.GetResult(out item);
                item.GetDisplayName(SIGDN.SIGDN_FILESYSPATH, out string name);
                this.Path = name;

                return true;
            }
            finally
            {
                Marshal.FinalReleaseComObject(dlg);
            }
        }
    }
}
