using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace COMInterfaceWrapper.Native
{
	[ComImport, Guid("b4db1657-70d7-485e-8e3e-6fcb5a5c1802"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	internal interface IModalWindow
	{
		[PreserveSig]
		uint Show(IntPtr hwndOwner);
	}
}
