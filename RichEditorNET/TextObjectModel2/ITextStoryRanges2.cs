using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextStoryRanges"/> with TOM2 story range enumeration capabilities.
	/// </summary>
	[ComImport]
	[Guid("C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextStoryRanges2 : ITextStoryRanges
	{
		[DispId(0x0003)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 Item2(int Index);
	}
}
