using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextFont"/> with additional TOM2 character formatting properties.
	/// </summary>
	[ComImport]
	[Guid("C241F5E3-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextFont2 : ITextFont
	{
		[DispId(0x0600)]
		int Count { get; }

		[DispId(0x0601)]
		int AutoLigatures { get; set; }

		[DispId(0x0602)]
		int AutospaceAlpha { get; set; }

		[DispId(0x0603)]
		int AutospaceNumeric { get; set; }

		[DispId(0x0604)]
		int AutospaceParens { get; set; }

		[DispId(0x0605)]
		int CharRep { get; set; }

		[DispId(0x0606)]
		int CompressionMode { get; set; }

		[DispId(0x0607)]
		int Cookie { get; set; }

		[DispId(0x0608)]
		int DoubleStrike { get; set; }

		[DispId(0x0609)]
		ITextFont2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0610)]
		int LinkType { get; }

		[DispId(0x0611)]
		int MathZone { get; set; }

		[DispId(0x0612)]
		int ModWidthPairs { get; set; }

		[DispId(0x0613)]
		int ModWidthSpace { get; set; }

		[DispId(0x0614)]
		int OldNumbers { get; set; }

		[DispId(0x0615)]
		int OverlappingText { get; set; }

		[DispId(0x0616)]
		int PositionSubSuper { get; set; }

		[DispId(0x0617)]
		int Scaling { get; set; }

		[DispId(0x0618)]
		float SpaceExtension { get; set; }

		[DispId(0x0619)]
		int UnderlinePositionMode { get; set; }

		[DispId(0x0620)]
		void GetEffects(out int pValue, out int pMask);

		[DispId(0x0621)]
		void GetEffects2(out int pValue, out int pMask);

		[DispId(0x0622)]
		int GetProperty(int Type);

		[DispId(0x0623)]
		void GetPropertyInfo(int Index, out int pType, out int pValue);

		[DispId(0x0624)]
		int IsEqual2([MarshalAs(UnmanagedType.Interface)] ITextFont2 pFont);

		[DispId(0x0625)]
		void SetEffects(int Value, int Mask);

		[DispId(0x0626)]
		void SetEffects2(int Value, int Mask);

		[DispId(0x0627)]
		void SetProperty(int Type, int Value);
	}
}
