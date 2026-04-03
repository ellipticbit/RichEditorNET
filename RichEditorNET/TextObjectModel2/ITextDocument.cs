using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to a rich text document. This is the base TOM (Text Object Model) interface.
	/// </summary>
	[ComImport]
	[Guid("8CC497C0-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextDocument
	{
		[DispId(0x0000)]
		string Name
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(0x0001)]
		ITextSelection Selection
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x0002)]
		int StoryCount { get; }

		[DispId(0x0003)]
		ITextStoryRanges StoryRanges
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x0004)]
		int Saved { get; set; }

		[DispId(0x0005)]
		float DefaultTabStop { get; set; }

		[DispId(0x0006)]
		void New();

		[DispId(0x0007)]
		void Open(ref object pVar, int Flags, int CodePage);

		[DispId(0x0008)]
		void Save(ref object pVar, int Flags, int CodePage);

		[DispId(0x0009)]
		int Freeze();

		[DispId(0x000A)]
		int Unfreeze();

		[DispId(0x000B)]
		void BeginEditCollection();

		[DispId(0x000C)]
		void EndEditCollection();

		[DispId(0x000D)]
		int Undo(int Count);

		[DispId(0x000E)]
		int Redo(int Count);

		[DispId(0x000F)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange Range(int cpActive, int cpAnchor);

		[DispId(0x0010)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange RangeFromPoint(int x, int y);
	}
}
