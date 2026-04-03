using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Represents a span of continuous text in a document. This is the base range interface in the Text Object Model.
	/// </summary>
	[ComImport]
	[Guid("8CC497C2-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextRange
	{
		[DispId(0x0000)]
		string Text
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
			[param: MarshalAs(UnmanagedType.BStr)]
			set;
		}

		[DispId(0x0201)]
		int Char { get; set; }

		[DispId(0x0202)]
		ITextRange Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0203)]
		ITextRange FormattedText
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0204)]
		int Start { get; set; }

		[DispId(0x0205)]
		int End { get; set; }

		[DispId(0x0206)]
		ITextFont Font
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0207)]
		ITextPara Para
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0208)]
		int StoryLength { get; }

		[DispId(0x0209)]
		int StoryType { get; }

		[DispId(0x0210)]
		void Collapse(int bStart);

		[DispId(0x0211)]
		int Expand(int Unit);

		[DispId(0x0212)]
		int GetIndex(int Unit);

		[DispId(0x0213)]
		void SetIndex(int Unit, int Index, int Extend);

		[DispId(0x0214)]
		void SetRange(int cpAnchor, int cpActive);

		[DispId(0x0215)]
		int InRange([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		[DispId(0x0216)]
		int InStory([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		[DispId(0x0217)]
		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextRange pRange);

		[DispId(0x0218)]
		void Select();

		[DispId(0x0219)]
		int StartOf(int Unit, int Extend);

		[DispId(0x0220)]
		int EndOf(int Unit, int Extend);

		[DispId(0x0221)]
		int Move(int Unit, int Count);

		[DispId(0x0222)]
		int MoveStart(int Unit, int Count);

		[DispId(0x0223)]
		int MoveEnd(int Unit, int Count);

		[DispId(0x0224)]
		int MoveWhile(ref object Cset, int Count);

		[DispId(0x0225)]
		int MoveStartWhile(ref object Cset, int Count);

		[DispId(0x0226)]
		int MoveEndWhile(ref object Cset, int Count);

		[DispId(0x0227)]
		int MoveUntil(ref object Cset, int Count);

		[DispId(0x0228)]
		int MoveStartUntil(ref object Cset, int Count);

		[DispId(0x0229)]
		int MoveEndUntil(ref object Cset, int Count);

		[DispId(0x0230)]
		int FindText([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		[DispId(0x0231)]
		int FindTextStart([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		[DispId(0x0232)]
		int FindTextEnd([MarshalAs(UnmanagedType.BStr)] string bstr, int Count, int Flags);

		[DispId(0x0233)]
		int Delete(int Unit, int Count);

		[DispId(0x0234)]
		void Cut(out object pVar);

		[DispId(0x0235)]
		void Copy(out object pVar);

		[DispId(0x0236)]
		void Paste(ref object pVar, int Format);

		[DispId(0x0237)]
		int CanPaste(ref object pVar, int Format);

		[DispId(0x0238)]
		int CanEdit();

		[DispId(0x0239)]
		void ChangeCase(int Type);

		[DispId(0x0240)]
		void GetPoint(int Type, out int px, out int py);

		[DispId(0x0241)]
		void SetPoint(int x, int y, int Type, int Extend);

		[DispId(0x0242)]
		void ScrollIntoView(int Value);

		[DispId(0x0243)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetEmbeddedObject();
	}
}
