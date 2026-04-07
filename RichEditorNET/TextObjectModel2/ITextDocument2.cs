using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// TOM2 document interface. Inherits directly from <see cref="ITextDocument"/> (not ITextDocument2Old).
	/// All base members are re-declared and member order matches the native tom.h vtable exactly.
	/// </summary>
	[ComImport]
	[Guid("C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextDocument2
	{
		// -------- ITextDocument (19 slots) --------

		string Name
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		ITextSelection Selection
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		int StoryCount { get; }

		ITextStoryRanges StoryRanges
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		int Saved { get; set; }

		float DefaultTabStop { get; set; }

		void New();

		void Open(ref object pVar, int Flags, int CodePage);

		void Save(ref object pVar, int Flags, int CodePage);

		int Freeze();

		int Unfreeze();

		void BeginEditCollection();

		void EndEditCollection();

		int Undo(int Count);

		int Redo(int Count);

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange Range(int cpActive, int cpAnchor);

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange RangeFromPoint(int x, int y);

		// -------- ITextDocument2 own entries (44 slots, exact tom.h order) --------
		// Properties first, then methods.

		int CaretType { get; set; }

		ITextDisplays Displays
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		ITextFont2 DocumentFont
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		ITextPara2 DocumentPara
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int EastAsianFlags { get; }

		string Generator
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		int IMEInProgress { set; }

		int NotificationMode { get; set; }

		ITextSelection2 Selection2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		ITextStoryRanges2 StoryRanges2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		int TypographyOptions { get; }

		int Version { get; }

		long Window { get; }

		void AttachMsgFilter([MarshalAs(UnmanagedType.IUnknown)] object pFilter);

		void CheckTextLimit(int cch, out int pcch);

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetCallManager();

		void GetClientRect(int Type, out int pLeft, out int pTop, out int pRight, out int pBottom);

		void GetEffectColor(int Index, out int pValue);

		long GetImmContext();

		void GetPreferredFont(
			int cp,
			int CharRep,
			int Options,
			int curCharRep,
			int curFontSize,
			[MarshalAs(UnmanagedType.BStr)] out string pbstr,
			out int pPitchAndFamily,
			out int pNewFontSize);

		void GetProperty(int Type, out int pValue);

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStrings GetStrings();

		void Notify(int Notify);

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 Range2(int cpActive, int cpAnchor);

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 RangeFromPoint2(int x, int y, int Type);

		void ReleaseCallManager([MarshalAs(UnmanagedType.IUnknown)] object pVoid);

		void ReleaseImmContext(long Context);

		void SetEffectColor(int Index, int Value);

		void SetProperty(int Type, int Value);

		void SetTypographyOptions(int Options, int Mask);

		void SysBeep();

		void Update(int Value);

		void UpdateWindow();

		int GetMathProperties();

		void SetMathProperties(int Options, int Mask);

		ITextStory ActiveStory
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		ITextStory MainStory
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		ITextStory NewStory
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStory GetStory(int Index);
	}
}
