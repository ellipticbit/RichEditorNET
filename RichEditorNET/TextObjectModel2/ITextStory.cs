using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to a story in a TOM2 document.
	/// </summary>
	[ComImport]
	[Guid("C241F5F3-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	public interface ITextStory
	{
		int Active { get; }

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Display { get; }

		int Index { get; }

		int Type { get; set; }

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 GetRange(int cpActive, int cpAnchor);

		[return: MarshalAs(UnmanagedType.BStr)]
		string GetText(int Flags);

		void SetFormattedText([MarshalAs(UnmanagedType.IUnknown)] object pUnk);

		void SetText(int Flags, [MarshalAs(UnmanagedType.BStr)] string bstr);
	}
}
