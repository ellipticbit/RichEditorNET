namespace EllipticBit.RichEditorNET.TextObjectModel2
{
	/// <summary>
	/// Specifies the type of inline object in a TOM2 document.
	/// Used with <see cref="ITextRange2.GetInlineObject"/>, <see cref="ITextRange2.SetInlineObject"/>,
	/// and <see cref="ITextStrings.EncodeFunction"/>.
	/// </summary>
	public enum OBJECTTYPE
	{
		SimpleText = 0,
		Ruby = 1,
		HorzVert = 2,
		Warichu = 3,
		Accent = 10,
		Box = 11,
		BoxedFormula = 12,
		Brackets = 13,
		BracketsWithSeps = 14,
		EquationArray = 15,
		Fraction = 16,
		FunctionApply = 17,
		LeftSubSup = 18,
		LowerLimit = 19,
		Matrix = 20,
		Nary = 21,
		OpChar = 22,
		Overbar = 23,
		Phantom = 24,
		Radical = 25,
		SlashedFraction = 26,
		Stack = 27,
		StretchStack = 28,
		Subscript = 29,
		SubSup = 30,
		Superscript = 31,
		Underbar = 32,
		UpperLimit = 33,
	}
}
