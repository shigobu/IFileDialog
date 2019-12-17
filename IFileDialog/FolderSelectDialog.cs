﻿using System;
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

        public bool ShowDialog()
        {
            return ShowDialog(IntPtr.Zero);
        }

        public bool ShowDialog(IntPtr owner)
        {
            var dlg = new FileOpenDialog() as IFileOpenDialog;
            try
            {
                if (owner == IntPtr.Zero)
                {
                    owner = NativeMethods.GetForegroundWindow();
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

                var hr = dlg.Show(owner);
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
                catch (ArgumentException ex) when (ex.Message == "値が有効な範囲にありません。")
                {
                    throw new Exception("「PC」「ネットワーク」は指定できません。", ex);
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
