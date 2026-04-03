using System;
using System.Runtime.InteropServices;

namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Extends <see cref="ITextPara"/> with additional TOM2 paragraph formatting properties.
	/// </summary>
	[ComImport]
	[Guid("C241F5E4-7206-11D8-A2C7-00A0D1D6C6B3")]
	[InterfaceType(ComInterfaceType.InterfaceIsDual)]
	public interface ITextPara2 : ITextPara
	{
		[DispId(0x0800)]
		object Borders
		{
			[return: MarshalAs(UnmanagedType.IUnknown)]
			get;
		}

		[DispId(0x0801)]
		ITextPara2 Duplicate2
		{
			[return: MarshalAs(UnmanagedType.Interface)]
			get;
			[param: MarshalAs(UnmanagedType.Interface)]
			set;
		}

		[DispId(0x0802)]
		int FontAlignment { get; set; }

		[DispId(0x0803)]
		int HangingPunctuation { get; set; }

		[DispId(0x0804)]
		int SnapToGrid { get; set; }

		[DispId(0x0805)]
		int TrimPunctuationAtStart { get; set; }

		[DispId(0x0806)]
		void GetEffects(out int pValue, out int pMask);

		[DispId(0x0807)]
		int GetProperty(int Type);

		[DispId(0x0808)]
		int IsEqual2([MarshalAs(UnmanagedType.Interface)] ITextPara2 pPara);

		[DispId(0x0809)]
		void SetEffects(int Value, int Mask);

		[DispId(0x0810)]
		void SetProperty(int Type, int Value);
	}
}
