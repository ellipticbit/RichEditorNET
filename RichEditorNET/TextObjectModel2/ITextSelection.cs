using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Represents the active selection in a document. Extends <see cref="ITextRange"/> with selection-specific navigation and input methods.
	/// </summary>
	[ComImport]
	[Guid("8CC497C1-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextSelection : ITextRange
	{
		[DispId(0x0301)]
		int Flags { get; set; }

		[DispId(0x0302)]
		int Type { get; }

		[DispId(0x0303)]
		int MoveLeft(int Unit, int Count, int Extend);

		[DispId(0x0304)]
		int MoveRight(int Unit, int Count, int Extend);

		[DispId(0x0305)]
		int MoveUp(int Unit, int Count, int Extend);

		[DispId(0x0306)]
		int MoveDown(int Unit, int Count, int Extend);

		[DispId(0x0307)]
		int HomeKey(int Unit, int Extend);

		[DispId(0x0308)]
		int EndKey(int Unit, int Extend);

		[DispId(0x0309)]
		void TypeText([MarshalAs(UnmanagedType.BStr)] string bstr);
	}
}
