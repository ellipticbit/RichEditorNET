using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Represents the active selection in a TOM2 document. Extends <see cref="ITextRange2"/> with the combined
	/// capabilities of both <see cref="ITextSelection"/> and <see cref="ITextRange2"/>.
	/// </summary>
	[ComImport]
	[Guid("C241F5E1-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextSelection2 : ITextRange2
	{
	}
}
