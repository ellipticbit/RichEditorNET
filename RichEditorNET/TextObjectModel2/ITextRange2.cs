using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// TOM2 range interface. Inherits from <see cref="ITextSelection"/> via flat re-declaration.
	/// Member order matches the native tom.h vtable exactly.
	/// </summary>
	[ComImport]
	[Guid("C241F5E2-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextRange2
	{
		// -------- ITextRange (52 slots) --------

		string Text
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		int Char { get; set; }

		ITextRange Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		ITextRange FormattedText
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int Start { get; set; }

		int End { get; set; }

		ITextFont Font
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		ITextPara Para
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int StoryLength { get; }

		int StoryType { get; }

		void Collapse(int bStart);

		int Expand(int Unit);

		int GetIndex(int Unit);

		void SetIndex(int Unit, int Index, int Extend);

		void SetRange(int cpAnchor, int cpActive);

		int InRange([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		int InStory([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		void Select();

		int StartOf(int Unit, int Extend);

		int EndOf(int Unit, int Extend);

		int Move(int Unit, int Count);

		int MoveStart(int Unit, int Count);

		int MoveEnd(int Unit, int Count);

		int MoveWhile(ref object Cset, int Count);

		int MoveStartWhile(ref object Cset, int Count);

		int MoveEndWhile(ref object Cset, int Count);

		int MoveUntil(ref object Cset, int Count);

		int MoveStartUntil(ref object Cset, int Count);

		int MoveEndUntil(ref object Cset, int Count);

		int FindText([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		int FindTextStart([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		int FindTextEnd([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		int Delete(int Unit, int Count);

		void RangeCut(out object pVar);

		void RangeCopy(out object pVar);

		void Paste(ref object pVar, int Format);

		int CanPaste(ref object pVar, int Format);

		int CanEdit();

		void ChangeCase(int Type);

		void GetPoint(int Type, out int px, out int py);

		void SetPoint(int x, int y, int Type, int Extend);

		void ScrollIntoView(int Value);

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetEmbeddedObject();

		// -------- ITextSelection (10 slots) --------

		int Flags { get; set; }

		int Type { get; }

		int MoveLeft(int Unit, int Count, int Extend);

		int MoveRight(int Unit, int Count, int Extend);

		int MoveUp(int Unit, int Count, int Extend);

		int MoveDown(int Unit, int Count, int Extend);

		int HomeKey(int Unit, int Extend);

		int EndKey(int Unit, int Extend);

		void TypeText([MarshalAs(UnmanagedType.BStr)] string bstr);

		// -------- ITextRange2 own entries (40 slots, exact tom.h order) --------
		// Properties first (18 slots), then methods (22 slots).

		int Cch { get; }

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Cells { get; }

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Column { get; }

		int Count { get; }

		ITextRange2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		ITextFont2 Font2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		ITextRange2 FormattedText2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int Gravity { get; set; }

		ITextPara2 Para2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Row { get; }

		int StartPara { get; }

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Table { get; }

		string URL
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		void AddSubrange(int cp1, int cp2, int Activate);

		void BuildUpMath(int Flags);

		void DeleteSubrange(int cpFirst, int cpLim);

		void Find([MarshalAs(UnmanagedType.Interface)] ITextRange2 pRange, int Count, int Flags, out int pDelta);

		void GetChar2(out int pChar, int Offset);

		void GetDropCap(out int pcLine, out int pPosition);

		void GetInlineObject(out int pType, out int pAlign, out int pChar, out int pChar1, out int pChar2, out int pCount, out int pTeXStyle, out int pcCol, out int pLevel);

		void GetProperty(int Type, out int pValue);

		void GetRect(int Type, out int pLeft, out int pTop, out int pRight, out int pBottom, out int pHit);

		void GetSubrange(int iSubrange, out int pcpFirst, out int pcpLim);

		[return: MarshalAs(UnmanagedType.BStr)]
		string GetText2(int Flags);

		void HexToUnicode();

		void InsertTable(int cCol, int cRow, int AutoFit);

		void Linearize(int Flags);

		void SetActiveSubrange(int cpAnchor, int cpActive);

		void SetDropCap(int cLine, int Position);

		void SetProperty(int Type, int Value);

		void SetText2(int Flags, [MarshalAs(UnmanagedType.BStr)] string bstr);

		void UnicodeToHex();

		void SetInlineObject(int Type, int Align, int Char, int Char1, int Char2, int Count, int TeXStyle, int cCol);

		void GetMathFunctionType([MarshalAs(UnmanagedType.BStr)] string bstr, out int pValue);

		void InsertImage(int width, int height, int ascent, int Type, [MarshalAs(UnmanagedType.BStr)] string bstrAltText, [MarshalAs(UnmanagedType.Interface)] IStream pStream);
	}
}
