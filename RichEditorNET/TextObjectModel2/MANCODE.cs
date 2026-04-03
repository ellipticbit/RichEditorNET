namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Specifies mathematical annotation character codes used with inline math objects.
	/// Used with <see cref="ITextRange2.GetInlineObject"/>, <see cref="ITextRange2.SetInlineObject"/>,
	/// and <see cref="ITextStrings.EncodeFunction"/> for the Char, Char1, and Char2 parameters.
	/// </summary>
	public enum MANCODE
	{
		None = 0,

		// Parentheses and brackets
		LeftParen = 0x0028,
		RightParen = 0x0029,
		LeftSquareBracket = 0x005B,
		RightSquareBracket = 0x005D,
		LeftCurlyBrace = 0x007B,
		VerticalBar = 0x007C,
		RightCurlyBrace = 0x007D,

		// Ceiling and floor
		LeftCeiling = 0x2308,
		RightCeiling = 0x2309,
		LeftFloor = 0x230A,
		RightFloor = 0x230B,

		// Double and angle brackets
		DoubleVerticalBar = 0x2016,
		LeftAngleBracket = 0x27E8,
		RightAngleBracket = 0x27E9,

		// N-ary operators
		Product = 0x220F,
		Coproduct = 0x2210,
		Summation = 0x2211,
		Integral = 0x222B,
		DoubleIntegral = 0x222C,
		TripleIntegral = 0x222D,
		ContourIntegral = 0x222E,
		SurfaceIntegral = 0x222F,
		VolumeIntegral = 0x2230,
		LogicalAnd = 0x22C0,
		LogicalOr = 0x22C1,
		Intersection = 0x22C2,
		Union = 0x22C3,
	}
}
