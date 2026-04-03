using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Provides access to table row properties in a TOM2 document.
	/// </summary>
	[ComImport]
	[Guid("C241F5EF-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextRow
	{
		[DispId(0x0000)]
		int Alignment { get; set; }

		[DispId(0x0901)]
		int CellCount { get; set; }

		[DispId(0x0902)]
		int CellCountCache { get; set; }

		[DispId(0x0903)]
		int CellIndex { get; set; }

		[DispId(0x0904)]
		int CellMargin { get; set; }

		[DispId(0x0905)]
		int Height { get; set; }

		[DispId(0x0906)]
		int IndentLevel { get; set; }

		[DispId(0x0907)]
		int KeepTogether { get; set; }

		[DispId(0x0908)]
		int KeepWithNext { get; set; }

		[DispId(0x0909)]
		int NestLevel { get; }

		[DispId(0x0910)]
		int RTL { get; set; }

		[DispId(0x0911)]
		int CellAlignment { get; set; }

		[DispId(0x0912)]
		int CellColorBack { get; set; }

		[DispId(0x0913)]
		int CellColorFore { get; set; }

		[DispId(0x0914)]
		int CellMergeFlags { get; set; }

		[DispId(0x0915)]
		int CellShading { get; set; }

		[DispId(0x0916)]
		int CellVerticalText { get; set; }

		[DispId(0x0917)]
		int CellWidth { get; set; }

		[DispId(0x0918)]
		void GetCellBorderColors(out int pcrLeft, out int pcrTop, out int pcrRight, out int pcrBottom);

		[DispId(0x0919)]
		void GetCellBorderWidths(out int pduLeft, out int pduTop, out int pduRight, out int pduBottom);

		[DispId(0x0920)]
		void SetCellBorderColors(int crLeft, int crTop, int crRight, int crBottom);

		[DispId(0x0921)]
		void SetCellBorderWidths(int duLeft, int duTop, int duRight, int duBottom);

		[DispId(0x0922)]
		void Apply(int cRow, int Flags);

		[DispId(0x0923)]
		int CanChange();

		[DispId(0x0924)]
		int GetProperty(int Type);

		[DispId(0x0925)]
		void Insert(int cRow);

		[DispId(0x0926)]
		int IsEqual([MarshalAs(UnmanagedType.Interface)] ITextRow pRow);

		[DispId(0x0927)]
		void Reset(int Value);

		[DispId(0x0928)]
		void SetProperty(int Type, int Value);
	}
}
