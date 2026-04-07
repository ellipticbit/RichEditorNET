using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextStoryRanges"/> with TOM2 capabilities.
	/// All <see cref="ITextStoryRanges"/> members are re-declared to ensure correct COM vtable layout,
	/// working around a known .NET interop issue with inherited dual interfaces.
	/// Member order must exactly match the native tom.h IDL vtable layout.
	/// </summary>
	[ComImport]
	[Guid("C241F5E5-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextStoryRanges2
	{
		// -------- ITextStoryRanges vtable entries (3 slots) --------

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetEnumerator();

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange Item(int Index);

		int Count { get; }

		// -------- ITextStoryRanges2 methods (1 slot) --------

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 Item2(int Index);
	}
}
