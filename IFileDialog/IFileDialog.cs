using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMInterfaceWrapper.Native
{
	[StructLayout(LayoutKind.Sequential)]
	public struct COMDLG_FILTERSPEC
	{
		[MarshalAs(UnmanagedType.LPWStr)]
		string pszName;
		[MarshalAs(UnmanagedType.LPWStr)]
		string pszSpec;
	}

	[Flags]
	public enum FILEOPENDIALOGOPTIONS : uint
	{
		FOS_OVERWRITEPROMPT =			0x2,
		FOS_STRICTFILETYPES =			0x4,
		FOS_NOCHANGEDIR =				0x8,
		FOS_PICKFOLDERS =				0x20,
		FOS_FORCEFILESYSTEM =			0x40,
		FOS_ALLNONSTORAGEITEMS =		0x80,
		FOS_NOVALIDATE =				0x100,
		FOS_ALLOWMULTISELECT =			0x200,
		FOS_PATHMUSTEXIST =				0x800,
		FOS_FILEMUSTEXIST =				0x1000,
		FOS_CREATEPROMPT =				0x2000,
		FOS_SHAREAWARE =				0x4000,
		FOS_NOREADONLYRETURN =			0x8000,
		FOS_NOTESTFILECREATE =			0x10000,
		FOS_HIDEMRUPLACES =				0x20000,
		FOS_HIDEPINNEDPLACES =			0x40000,
		FOS_NODEREFERENCELINKS =		0x100000,
		FOS_OKBUTTONNEEDSINTERACTION =	0x200000,
		FOS_DONTADDTORECENT =			0x2000000,
		FOS_FORCESHOWHIDDEN =			0x10000000,
		FOS_DEFAULTNOMINIMODE =			0x20000000,
		FOS_FORCEPREVIEWPANEON =		0x40000000,
		FOS_SUPPORTSTREAMABLEITEMS =	0x80000000
	};

	enum FDAP
	{
		FDAP_BOTTOM,
		FDAP_TOP
	};

	[ComImport, Guid("42f85136-db7e-439c-85f1-e4075d135fc8")]
	internal interface IFileDialog : IModalWindow
	{
		void SetFileTypes(uint cFileTypes, COMDLG_FILTERSPEC[] rgFilterSpec);

		void SetFileTypeIndex(uint iFileType);
        
		//[return:MarshalAs(UnmanagedType.U4)]
        uint GetFileTypeIndex();
        
		[return:MarshalAs(UnmanagedType.U4)]
        uint Advise(IntPtr pfde);

		void Unadvise(uint dwCookie);
        
        void SetOptions([MarshalAs(UnmanagedType.U4)]FILEOPENDIALOGOPTIONS fos);

		[return: MarshalAs(UnmanagedType.U4)]
		uint GetOptions();
        
        void SetDefaultFolder(IShellItem psi);
        
        void SetFolder(IShellItem psi);

		IShellItem GetFolder();

		IShellItem GetCurrentSelection();
        
        void SetFileName(string pszName);
        
        string GetFileName();
        
        void SetTitle(string pszTitle);
        
        void SetOkButtonLabel(string pszText);
        
        void SetFileNameLabel(string pszLabel);

		IShellItem GetResult();
        
        void AddPlace(IShellItem psi, FDAP fdap);
        
        void SetDefaultExtension(string pszDefaultExtension);
        
        void Close(uint hr);
        
        void SetClientGuid(in Guid guid);
        
        void ClearClientData();

		/// <summary>
		/// SetFilter is no longer available for use as of Windows 7.
		/// </summary>
		/// <param name="pFilter"></param>
		void SetFilter(IntPtr pFilter);

	}
}
