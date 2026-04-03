using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Enumerates the stories in a document.
	/// </summary>
	[ComImport]
	[Guid("8CC497C5-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextStoryRanges
	{
		[DispId(-4)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetEnumerator();

		[DispId(0x0000)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange Item(int Index);

		[DispId(0x0002)]
		int Count { get; }
	}
}
