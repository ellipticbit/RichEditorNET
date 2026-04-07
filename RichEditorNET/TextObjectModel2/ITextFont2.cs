using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// TOM2 font interface. All <see cref="ITextFont"/> members re-declared for correct vtable layout.
	/// Member order matches the native tom.h vtable exactly.
	/// </summary>
	[ComImport]
	[Guid("C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextFont2
	{
		// -------- ITextFont (55 slots) --------

		ITextFont Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int CanChange();

		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextFont pFont);

		void Reset(int Value);

		int Style { get; set; }

		int AllCaps { get; set; }

		int Animation { get; set; }

		int BackColor { get; set; }

		int Bold { get; set; }

		int Emboss { get; set; }

		int ForeColor { get; set; }

		int Hidden { get; set; }

		int Engrave { get; set; }

		int Italic { get; set; }

		float Kerning { get; set; }

		int LanguageID { get; set; }

		string Name
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		int Outline { get; set; }

		float Position { get; set; }

		int Protected { get; set; }

		int Shadow { get; set; }

		float Size { get; set; }

		int SmallCaps { get; set; }

		float Spacing { get; set; }

		int StrikeThrough { get; set; }

		int Subscript { get; set; }

		int Superscript { get; set; }

		int Underline { get; set; }

		int Weight { get; set; }

		// -------- ITextFont2 own entries (46 slots, exact tom.h order) --------
		// Properties first (38 slots), then methods (8 slots).

		int Count { get; }

		int AutoLigatures { get; set; }

		int AutospaceAlpha { get; set; }

		int AutospaceNumeric { get; set; }

		int AutospaceParens { get; set; }

		int CharRep { get; set; }

		int CompressionMode { get; set; }

		int Cookie { get; set; }

		int DoubleStrike { get; set; }

		ITextFont2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int LinkType { get; }

		int MathZone { get; set; }

		int ModWidthPairs { get; set; }

		int ModWidthSpace { get; set; }

		int OldNumbers { get; set; }

		int Overlapping { get; set; }

		int PositionSubSuper { get; set; }

		int Scaling { get; set; }

		float SpaceExtension { get; set; }

		int UnderlinePositionMode { get; set; }

		void GetEffects(out int pValue, out int pMask);

		void GetEffects2(out int pValue, out int pMask);

		void GetProperty2(int Type, out int pValue);

		void GetPropertyInfo(int Index, out int pType, out int pValue);

		int IsEqual2([MarshalAs(UnmanagedType.Interface)] ITextFont2 pFont);

		void SetEffects(int Value, int Mask);

		void SetEffects2(int Value, int Mask);

		void SetProperty2(int Type, int Value);
	}
}
