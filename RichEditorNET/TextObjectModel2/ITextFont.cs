using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to character formatting properties for a text range.
	/// </summary>
	[ComImport]
	[Guid("8CC497C3-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextFont
	{
		[DispId(0x0000)]
		ITextFont Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0501)]
		int CanChange();

		[DispId(0x0502)]
		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextFont pFont);

		[DispId(0x0503)]
		void Reset(int Value);

		[DispId(0x0504)]
		int Style { get; set; }

		[DispId(0x0505)]
		int AllCaps { get; set; }

		[DispId(0x0506)]
		int Animation { get; set; }

		[DispId(0x0507)]
		int BackColor { get; set; }

		[DispId(0x0508)]
		int Bold { get; set; }

		[DispId(0x0509)]
		int Emboss { get; set; }

		[DispId(0x0510)]
		int ForeColor { get; set; }

		[DispId(0x0511)]
		int Hidden { get; set; }

		[DispId(0x0512)]
		int Engrave { get; set; }

		[DispId(0x0513)]
		int Italic { get; set; }

		[DispId(0x0514)]
		float Kerning { get; set; }

		[DispId(0x0515)]
		int LanguageID { get; set; }

		[DispId(0x0516)]
		string Name
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(0x0517)]
		int Outline { get; set; }

		[DispId(0x0518)]
		float Position { get; set; }

		[DispId(0x0519)]
		int Protected { get; set; }

		[DispId(0x0520)]
		int Shadow { get; set; }

		[DispId(0x0521)]
		float Size { get; set; }

		[DispId(0x0522)]
		int SmallCaps { get; set; }

		[DispId(0x0523)]
		float Spacing { get; set; }

		[DispId(0x0524)]
		int StrikeThrough { get; set; }

		[DispId(0x0525)]
		int Subscript { get; set; }

		[DispId(0x0526)]
		int Superscript { get; set; }

		[DispId(0x0527)]
		int Underline { get; set; }

		[DispId(0x0528)]
		int Weight { get; set; }
	}
}
