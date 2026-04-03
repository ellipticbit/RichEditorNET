using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to paragraph formatting properties for a text range.
	/// </summary>
	[ComImport]
	[Guid("8CC497C4-A1DF-11CE-8098-00AA0047BE5D")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextPara
	{
		[DispId(0x0000)]
		ITextPara Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0701)]
		int CanChange();

		[DispId(0x0702)]
		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextPara pPara);

		[DispId(0x0703)]
		void Reset(int Value);

		[DispId(0x0704)]
		int Style { get; set; }

		[DispId(0x0705)]
		int Alignment { get; set; }

		[DispId(0x0706)]
		int Hyphenation { get; set; }

		[DispId(0x0707)]
		float FirstLineIndent { get; }

		[DispId(0x0708)]
		int KeepTogether { get; set; }

		[DispId(0x0709)]
		int KeepWithNext { get; set; }

		[DispId(0x0710)]
		float LeftIndent { get; }

		[DispId(0x0711)]
		float LineSpacing { get; }

		[DispId(0x0712)]
		int LineSpacingRule { get; }

		[DispId(0x0713)]
		int ListAlignment { get; set; }

		[DispId(0x0714)]
		int ListLevelIndex { get; set; }

		[DispId(0x0715)]
		int ListStart { get; set; }

		[DispId(0x0716)]
		float ListTab { get; set; }

		[DispId(0x0717)]
		int ListType { get; set; }

		[DispId(0x0718)]
		int NoLineNumber { get; set; }

		[DispId(0x0719)]
		int PageBreakBefore { get; set; }

		[DispId(0x0720)]
		float RightIndent { get; set; }

		[DispId(0x0721)]
		void SetIndents(float First, float Left, float Right);

		[DispId(0x0722)]
		void SetLineSpacing(int Rule, float Spacing);

		[DispId(0x0723)]
		float SpaceAfter { get; set; }

		[DispId(0x0724)]
		float SpaceBefore { get; set; }

		[DispId(0x0725)]
		int WidowControl { get; set; }

		[DispId(0x0726)]
		int TabCount { get; }

		[DispId(0x0727)]
		void AddTab(float tbPos, int tbAlign, int tbLeader);

		[DispId(0x0728)]
		void ClearAllTabs();

		[DispId(0x0729)]
		void DeleteTab(float tbPos);

		[DispId(0x0730)]
		void GetTab(int iTab, out float ptbPos, out int ptbAlign, out int ptbLeader);
	}
}
