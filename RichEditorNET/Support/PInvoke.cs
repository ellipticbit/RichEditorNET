using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.Support
{
	internal static class PInvoke
	{
		internal const string MSFTEDIT_DLL = "msftedit.dll";
		internal const string RICHEDIT_CLASS = "RICHEDIT50W";

		internal const int WM_SETREDRAW = 0x000B;
		internal const int WM_CONTEXTMENU = 0x007B;
		internal const int WM_ENTERMENULOOP = 0x0211;
		internal const int WM_USER = 0x0400;
		internal const int EM_GETOLEINTERFACE = WM_USER + 60;
		internal const int EM_SETOLECALLBACK = WM_USER + 76;
		internal const int EM_SETOPTIONS = WM_USER + 77;
		internal const int EM_GETOPTIONS = WM_USER + 78;
		internal const int EM_SETLANGOPTIONS = WM_USER + 120;
		internal const int EM_GETLANGOPTIONS = WM_USER + 121;
		internal const int IMF_SPELLCHECKING = 0x0800;

		internal const int ECO_AUTOWORDSELECTION = 0x00000001;
		internal const int ECOOP_OR = 0x0002;
		internal const int ECOOP_AND = 0x0003;

		[DllImport("kernel32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
		internal static extern IntPtr LoadLibrary(string lpLibFileName);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		internal static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, ref IntPtr lParam);

		[DllImport("ole32.dll")]
		internal static extern int CreateStreamOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out System.Runtime.InteropServices.ComTypes.IStream ppstm);

		[DllImport("ole32.dll")]
		internal static extern void ReleaseStgMedium([In] ref System.Runtime.InteropServices.ComTypes.STGMEDIUM pmedium);

		[DllImport("ole32.dll")]
		internal static extern int CreateILockBytesOnHGlobal(IntPtr hGlobal, [MarshalAs(UnmanagedType.Bool)] bool fDeleteOnRelease, out ILockBytes pplkbyt);

		[DllImport("ole32.dll")]
		internal static extern int StgCreateDocfileOnILockBytes(ILockBytes plkbyt, uint grfMode, uint reserved, out IStorage ppstgOpen);

		[DllImport("ole32.dll")]
		internal static extern int OleCreateStaticFromData(
			System.Runtime.InteropServices.ComTypes.IDataObject pSrcDataObj,
			ref Guid riid,
			uint renderopt,
			ref System.Runtime.InteropServices.ComTypes.FORMATETC pFormatEtc,
			IOleClientSite pClientSite,
			IStorage pStg,
			out IntPtr ppvObj);

		[DllImport("ole32.dll")]
		internal static extern int OleSetContainedObject(
			[MarshalAs(UnmanagedType.IUnknown)] object pUnknown,
			[MarshalAs(UnmanagedType.Bool)] bool fContained);

		internal const uint OLERENDER_DRAW = 1;
		internal const uint OLERENDER_FORMAT = 2;
		internal const uint GMEM_MOVEABLE = 0x0002;
		internal const uint GMEM_ZEROINIT = 0x0040;

		[DllImport("kernel32.dll")]
		internal static extern IntPtr GlobalAlloc(uint uFlags, UIntPtr dwBytes);

		[DllImport("kernel32.dll")]
		internal static extern IntPtr GlobalFree(IntPtr hMem);

		[DllImport("gdi32.dll")]
		internal static extern IntPtr SetEnhMetaFileBits(uint cbBuffer, byte[] lpData);

		[DllImport("gdi32.dll")]
		internal static extern bool DeleteEnhMetaFile(IntPtr hemf);

		[DllImport("gdi32.dll")]
		internal static extern uint GetEnhMetaFileBits(IntPtr hemf, uint cbBuffer, byte[] lpbBuffer);

		[DllImport("kernel32.dll")]
		internal static extern IntPtr GlobalLock(IntPtr hMem);

		[DllImport("kernel32.dll")]
		internal static extern bool GlobalUnlock(IntPtr hMem);

		[DllImport("kernel32.dll")]
		internal static extern UIntPtr GlobalSize(IntPtr hMem);
	}
}
