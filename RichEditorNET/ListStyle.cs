using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET
{
	/// <summary>
	/// Specifies the list style to apply. Values correspond to HTML5/CSS list-style-type options.
	/// </summary>
	public enum ListStyle
	{
		/// <summary>An unordered bullet list (CSS: disc).</summary>
		Bullet = tomConstants.tomListBullet,
		/// <summary>An ordered list with decimal numbers (CSS: decimal).</summary>
		Decimal = tomConstants.tomListNumberAsArabic,
		/// <summary>An ordered list with lowercase letters (CSS: lower-alpha).</summary>
		LowerAlpha = tomConstants.tomListNumberAsLCLetter,
		/// <summary>An ordered list with uppercase letters (CSS: upper-alpha).</summary>
		UpperAlpha = tomConstants.tomListNumberAsUCLetter,
		/// <summary>An ordered list with lowercase Roman numerals (CSS: lower-roman).</summary>
		LowerRoman = tomConstants.tomListNumberAsLCRoman,
		/// <summary>An ordered list with uppercase Roman numerals (CSS: upper-roman).</summary>
		UpperRoman = tomConstants.tomListNumberAsUCRoman,
	}
}
