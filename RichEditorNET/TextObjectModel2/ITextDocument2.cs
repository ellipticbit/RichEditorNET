using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextDocument"/> with additional TOM2 (Text Object Model 2) capabilities.
	/// </summary>
	[ComImport]
	[Guid("C241F5E0-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextDocument2 : ITextDocument
	{
		[DispId(0x0100)]
		int CaretType { get; set; }

		[DispId(0x0101)]
		ITextDisplays Displays
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x0102)]
		ITextFont2 DocumentFont
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0103)]
		ITextPara2 DocumentPara
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0104)]
		int EastAsianFlags { get; }

		[DispId(0x0105)]
		string Generator
		{
			[return: MarshalAs(UnmanagedType.BStr)]
			get;
		}

		[DispId(0x0106)]
		int IMEInProgress { get; }

		[DispId(0x0107)]
		int NotificationMode { get; set; }

		[DispId(0x0108)]
		ITextSelection2 Selection2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x0109)]
		ITextStoryRanges2 StoryRanges2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
		}

		[DispId(0x010A)]
		int TypographyOptions { get; }

		[DispId(0x010B)]
		int Version { get; }

		[DispId(0x010C)]
		int Window { get; }

		[DispId(0x010D)]
		void AttachMsgFilter([MarshalAs(UnmanagedType.IUnknown)] object pFilter);

		[DispId(0x010E)]
		int CheckTextLimit(int cch);

		[DispId(0x010F)]
		[return: MarshalAs(UnmanagedType.IUnknown)]
		object GetCallManager();

		[DispId(0x0110)]
		void GetClientRect(int Type, out int pLeft, out int pTop, out int pRight, out int pBottom);

		[DispId(0x0111)]
		int GetEffectColor(int Index);

		[DispId(0x0112)]
		int GetImmContext();

		[DispId(0x0113)]
		void GetPreferredFont(
			int cp,
			int CharRep,
			int Options,
			int curCharRep,
			int curFontSize,
			[MarshalAs(UnmanagedType.BStr)] out string pbstr,
			out int pPitchAndFamily,
			out int pNewFontSize);

		[DispId(0x0114)]
		int GetProperty(int Type);

		[DispId(0x0115)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStrings GetStrings();

		[DispId(0x0116)]
		void Notify(int Notify);

		[DispId(0x0117)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 Range2(int cpActive, int cpAnchor);

		[DispId(0x0118)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextRange2 RangeFromPoint2(int x, int y);

		[DispId(0x0119)]
		void ReleaseCallManager([MarshalAs(UnmanagedType.IUnknown)] object pVoid);

		[DispId(0x011A)]
		void ReleaseImmContext(int Context);

		[DispId(0x011B)]
		void SetEffectColor(int Index, int Value);

		[DispId(0x011C)]
		void SetProperty(int Type, int Value);

		[DispId(0x011D)]
		void SetTypographyOptions(int Options, int Mask);

		[DispId(0x011E)]
		void SysBeep();

		[DispId(0x011F)]
		void Update(int Value);

		[DispId(0x0120)]
		void UpdateWindow();

		[DispId(0x0121)]
		int GetMathProperties();

		[DispId(0x0122)]
		void SetMathProperties(int Options, int Mask);

		[DispId(0x0123)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStory GetActiveStory();

		[DispId(0x0124)]
		void SetActiveStory([MarshalAs(UnmanagedType.Interface)] ITextStory pStory);

		[DispId(0x0125)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStory GetMainStory();

		[DispId(0x0126)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStory GetNewStory();

		[DispId(0x0127)]
		[return: MarshalAs(UnmanagedType.Interface)]
		ITextStory GetStory(int Index);
	}
}
