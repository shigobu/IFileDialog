﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace COMInterfaceWrapper.Native
{
	[SuppressUnmanagedCodeSecurity]
	internal static class NativeMethods
	{
		[DllImport("shell32.dll", ExactSpelling = true, CharSet = CharSet.Unicode, PreserveSig = false)]
		public static extern IShellItem SHCreateItemFromParsingName(
			string pszPath,
			IntPtr pbc,
			[MarshalAs(UnmanagedType.LPStruct)]Guid riid);
	}

	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	[Guid("43826d1e-e718-42ee-bc55-a1e261c37bfe")]
	internal interface IShellItem
	{
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object BindToHandler(
			IntPtr pbc,
			[MarshalAs(UnmanagedType.LPStruct)]Guid bhid,
			[MarshalAs(UnmanagedType.LPStruct)]Guid riid);
		IShellItem GetParent();
		SafeCoTaskMemHandle GetDisplayName(SIGDN sigdnName);
		SFGAOF GetAttributes(SFGAOF sfgaoMask);
		int Compare(IShellItem psi, SICHINTF hint);
	}

	public enum SIGDN : uint
	{
		SIGDN_NORMALDISPLAY = 0x00000000,
		SIGDN_PARENTRELATIVEPARSING = 0x80018001,
		SIGDN_DESKTOPABSOLUTEPARSING = 0x80028000,
		SIGDN_PARENTRELATIVEEDITING = 0x80031001,
		SIGDN_DESKTOPABSOLUTEEDITING = 0x8004c000,
		SIGDN_FILESYSPATH = 0x80058000,
		SIGDN_URL = 0x80068000,
		SIGDN_PARENTRELATIVEFORADDRESSBAR = 0x8007c001,
		SIGDN_PARENTRELATIVE = 0x80080001,
		SIGDN_PARENTRELATIVEFORUI = 0x80094001,
	}

	public enum SICHINTF : uint
	{
		SICHINT_DISPLAY = 0x00000000,
		SICHINT_ALLFIELDS = 0x80000000,
		SICHINT_CANONICAL = 0x10000000,
		SICHINT_TEST_FILESYSPATH_IF_NOT_EQUAL = 0x20000000,
	}

	[Flags]
	public enum SFGAOF : uint
	{
		SFGAO_CANCOPY = 0x1,
		SFGAO_CANMOVE = 0x2,
		SFGAO_CANLINK = 0x4,
		SFGAO_STORAGE = 0x00000008,
		SFGAO_CANRENAME = 0x00000010,
		SFGAO_CANDELETE = 0x00000020,
		SFGAO_HASPROPSHEET = 0x00000040,
		SFGAO_DROPTARGET = 0x00000100,
		SFGAO_CAPABILITYMASK = 0x00000177,
		SFGAO_PLACEHOLDER = 0x00000800,
		SFGAO_SYSTEM = 0x00001000,
		SFGAO_ENCRYPTED = 0x00002000,
		SFGAO_ISSLOW = 0x00004000,
		SFGAO_GHOSTED = 0x00008000,
		SFGAO_LINK = 0x00010000,
		SFGAO_SHARE = 0x00020000,
		SFGAO_READONLY = 0x00040000,
		SFGAO_HIDDEN = 0x00080000,
		SFGAO_DISPLAYATTRMASK = 0x000FC000,
		SFGAO_FILESYSANCESTOR = 0x10000000,
		SFGAO_FOLDER = 0x20000000,
		SFGAO_FILESYSTEM = 0x40000000,
		SFGAO_HASSUBFOLDER = 0x80000000,
		SFGAO_CONTENTSMASK = 0x80000000,
		SFGAO_VALIDATE = 0x01000000,
		SFGAO_REMOVABLE = 0x02000000,
		SFGAO_COMPRESSED = 0x04000000,
		SFGAO_BROWSABLE = 0x08000000,
		SFGAO_NONENUMERATED = 0x00100000,
		SFGAO_NEWCONTENT = 0x00200000,
		SFGAO_CANMONIKER = 0x00400000,
		SFGAO_HASSTORAGE = 0x00400000,
		SFGAO_STREAM = 0x00400000,
		SFGAO_STORAGEANCESTOR = 0x00800000,
		SFGAO_STORAGECAPMASK = 0x70C50008,
		SFGAO_PKEYSFGAOMASK = 0x81044000,
	}
}

internal sealed class SafeCoTaskMemHandle : SafeHandle
{
	public SafeCoTaskMemHandle()
		: base(IntPtr.Zero, true)
	{
	}

	public SafeCoTaskMemHandle(IntPtr handle, bool ownsHandle)
		: base(IntPtr.Zero, ownsHandle)
	{
		SetHandle(handle);
	}

	public override bool IsInvalid => handle == IntPtr.Zero;

	protected override bool ReleaseHandle()
	{
		Marshal.FreeCoTaskMem(handle);
		return true;
	}
}
