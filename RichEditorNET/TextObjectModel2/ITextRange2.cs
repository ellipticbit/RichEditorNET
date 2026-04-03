using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextSelection"/> with TOM2 range capabilities including math, tables, and inline objects.
	/// </summary>
	[ComImport]
	[Guid("C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextRange2 : ITextSelection
	{
		[DispId(0x0400)]
		int Count { get; }

		[DispId(0x0401)]
		int Gravity { get; set; }

		[DispId(0x0402)]
		string URL
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(0x0403)]
		ITextFont2 Font2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0404)]
		ITextPara2 Para2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0405)]
		ITextRange2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0406)]
		ITextRange2 FormattedText2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0407)]
		int GetCch();

		[DispId(0x0408)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetCells();

		[DispId(0x0409)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetColumn();

		[DispId(0x0410)]
		void GetDropCap(out int pcLine, out int pPosition);

		[DispId(0x0411)]
		void GetInlineObject(
			out OBJECTTYPE pType,
			out int pAlign,
			out MANCODE pChar,
			out MANCODE pChar1,
			out MANCODE pChar2,
			out int pCount,
			out int pTeXStyle,
			out int pcCol,
			out int pLevel);

		[DispId(0x0412)]
		int GetProperty(int Type);

		[DispId(0x0413)]
		void GetRect(int Type, out int pLeft, out int pTop, out int pRight, out int pBottom, out int pHit);

		[DispId(0x0414)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRow GetRow();

		[DispId(0x0415)]
		int GetStartPara();

		[DispId(0x0416)]
		void GetSubrange(int iSubrange, out int pcpFirst, out int pcpLim);

		[DispId(0x0417)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetTable();

		[DispId(0x0418)]
		[return: MarshalAs(UnmanagedType.BStr)]
		string GetText2(int Flags);

		[DispId(0x0419)]
		void HexToUnicode();

		[DispId(0x0420)]
		void InsertImage(int Width, int Height, int Ascent, int Type, [MarshalAs(UnmanagedType.BStr)] string bstrAltText, [MarshalAs(UnmanagedType.Interface)] IStream pStream);

		[DispId(0x0421)]
		void InsertTable(int cCol, int cRow, int AutoFit);

		[DispId(0x0422)]
		void Linearize(int Flags);

		[DispId(0x0423)]
		void SetActiveSubrange(int cpAnchor, int cpActive);

		[DispId(0x0424)]
		void SetDropCap(int cLine, int Position);

		[DispId(0x0425)]
		void SetInlineObject(OBJECTTYPE Type, int Align, MANCODE Char, MANCODE Char1, MANCODE Char2, int Count, int TeXStyle, int cCol);

		[DispId(0x0426)]
		void SetProperty(int Type, int Value);

		[DispId(0x0427)]
		void SetText2(int Flags, [MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0428)]
		void UnicodeToHex();

		[DispId(0x0429)]
		void SetURL([MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0430)]
		int InRange2([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange);

		[DispId(0x0431)]
		int IsEqual2([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange);

		[DispId(0x0432)]
		new void Copy([MarshalAs(UnmanagedType.IUnknown)] out object ppObject);

		[DispId(0x0433)]
		new void Cut([MarshalAs(UnmanagedType.IUnknown)] out object ppObject);

		[DispId(0x0434)]
		void DeleteSubrange(int cpFirst, int cpLim);

		[DispId(0x0435)]
		int Find([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange, int Count, int Flags);

		[DispId(0x0436)]
		int GetChar2(int Offset);

		[DispId(0x0437)]
		int GetMathFunctionType([MarshalAs(UnmanagedType.BStr)] string bstr);

		[DispId(0x0438)]
		void InsertMath(int Flags);

		[DispId(0x0439)]
		void AddSubrange(int cp1, int cp2, int Activate);

		[DispId(0x0440)]
		void BuildUpMath(int Flags);
	}
}
