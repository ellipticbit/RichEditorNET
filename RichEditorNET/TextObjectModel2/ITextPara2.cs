using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// TOM2 paragraph interface. All <see cref="ITextPara"/> members re-declared for correct vtable layout.
	/// Member order matches the native tom.h vtable exactly.
	/// </summary>
	[ComImport]
	[Guid("C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextPara2
	{
		// -------- ITextPara (48 slots) --------

		ITextPara Duplicate
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int CanChange();

		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextPara pPara);

		void Reset(int Value);

		int Style { get; set; }

		int Alignment { get; set; }

		int Hyphenation { get; set; }

		float FirstLineIndent { get; }

		int KeepTogether { get; set; }

		int KeepWithNext { get; set; }

		float LeftIndent { get; }

		float LineSpacing { get; }

		int LineSpacingRule { get; }

		int ListAlignment { get; set; }

		int ListLevelIndex { get; set; }

		int ListStart { get; set; }

		float ListTab { get; set; }

		int ListType { get; set; }

		int NoLineNumber { get; set; }

		int PageBreakBefore { get; set; }

		float RightIndent { get; set; }

		void SetIndents(float First, float Left, float Right);

		void SetLineSpacing(int Rule, float Spacing);

		float SpaceAfter { get; set; }

		float SpaceBefore { get; set; }

		int WidowControl { get; set; }

		int TabCount { get; }

		void AddTab(float tbPos, int tbAlign, int tbLeader);

		void ClearAllTabs();

		void DeleteTab(float tbPos);

		void GetTab(int iTab, out float ptbPos, out int ptbAlign, out int ptbLeader);

		// -------- ITextPara2 own entries (16 slots, exact tom.h order) --------
		// Properties first (11 slots), then methods (5 slots).

		[return: MarshalAs(UnmanagedType.IUnknown)]
		object Borders { get; }

		ITextPara2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		int FontAlignment { get; set; }

		int HangingPunctuation { get; set; }

		int SnapToGrid { get; set; }

		int TrimPunctuationAtStart { get; set; }

		void GetEffects(out int pValue, out int pMask);

		void GetProperty(int Type, out int pValue);

		int IsEqual2([MarshalAs(UnmanagedType.Interface)] ITextPara2 pPara);

		void SetEffects(int Value, int Mask);

		void SetProperty(int Type, int Value);
	}
}
