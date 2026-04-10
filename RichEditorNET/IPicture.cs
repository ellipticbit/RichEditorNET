using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET
{
	/// <summary>
	/// COM IPicture interface for extracting image data from OLE embedded objects.
	/// Vtable order matches ocidl.h exactly.
	/// </summary>
	[ComImport]
	[Guid("7BF80980-BF32-101A-8BBB-00AA00300CAB")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	internal interface IPicture
	{
		int get_Handle();
		int get_hPal();
		short get_Type();
		int get_Width();
		int get_Height();
		void Render(IntPtr hdc, int x, int y, int cx, int cy, int xSrc, int ySrc, int cxSrc, int cySrc, IntPtr prcWBounds);
		void set_hPal(int hPal);
		IntPtr get_CurDC();
		void SelectPicture(IntPtr hdcIn, out IntPtr phdcOut, out int phbmpOut);
		[return: MarshalAs(UnmanagedType.Bool)]
		bool get_KeepOriginalFormat();
		void set_KeepOriginalFormat([MarshalAs(UnmanagedType.Bool)] bool fKeep);
		void PictureChanged();
		void SaveAsFile([MarshalAs(UnmanagedType.Interface)] System.Runtime.InteropServices.ComTypes.IStream pstm, [MarshalAs(UnmanagedType.Bool)] bool fSaveMemCopy, out int pcbSize);
		int get_Attributes();
	}
}
