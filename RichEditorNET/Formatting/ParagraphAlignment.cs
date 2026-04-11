using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET.Formatting
{
	/// <summary>
	/// Specifies paragraph alignment. Values correspond to HTML5/CSS text-align options.
	/// </summary>
	public enum ParagraphAlignment
	{
		/// <summary>Left-aligned text (CSS: left).</summary>
		Left = tomConstants.tomAlignLeft,
		/// <summary>Center-aligned text (CSS: center).</summary>
		Center = tomConstants.tomAlignCenter,
		/// <summary>Right-aligned text (CSS: right).</summary>
		Right = tomConstants.tomAlignRight,
		/// <summary>Justified text (CSS: justify).</summary>
		Justify = tomConstants.tomAlignJustify,
	}
}
