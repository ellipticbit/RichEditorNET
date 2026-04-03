using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to a collection of rich text strings used for building up math and other complex content in a TOM2 document.
	/// </summary>
	[ComImport]
	[Guid("C241F5E7-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextStrings
	{
		[DispId(0x0000)]
		ITextRange2 Item
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x0A01)]
		int Count { get; }

		[DispId(0x0A02)]
		void Add([MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0A03)]
		void Append([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange, int iString);

		[DispId(0x0A04)]
		void Cat(int iString);

		[DispId(0x0A05)]
		void CatTop2([MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0A06)]
		void DeleteRange([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange);

		[DispId(0x0A07)]
		void EncodeFunction(
			OBJECTTYPE Type,
			int Align,
			MANCODE Char,
			MANCODE Char1,
			MANCODE Char2,
			int Count,
			int TeXStyle,
			int cCol,
			[MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange);

		[DispId(0x0A08)]
		int GetCch(int iString);

		[DispId(0x0A09)]
		void InsertNullStr(int iString);

		[DispId(0x0A10)]
		void MoveBoundary(int iString, int cch);

		[DispId(0x0A11)]
		void PrefixTop([MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0A12)]
		void Remove(int iString, int cString);

		[DispId(0x0A13)]
		void SetFormattedText(
			[MarshalAs(UnmanagedType.Interface)] ITextRange2 pRangeD,
			[MarshalAs(UnmanagedType.Interface)] ITextRange2 pRangeS);

		[DispId(0x0A14)]
		void SetOpCp(int iString, int cp);

		[DispId(0x0A15)]
		void SuffixTop(
			[MarshalAs(UnmanagedType.BStr)] string bstr,
			[MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange);

		[DispId(0x0A16)]
		void Swap();
	}
}
