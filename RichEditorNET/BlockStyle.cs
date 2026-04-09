using System;

namespace EllipticBit.RichEditorNET
{
	/// <summary>
	/// Defines block-level styles that map to HTML heading elements,
	/// standard paragraphs, and preformatted blocks.
	/// </summary>
	public enum BlockStyle
	{
		Paragraph = 0,
		Heading1 = 1,
		Heading2 = 2,
		Heading3 = 3,
		Heading4 = 4,
		Heading5 = 5,
		Heading6 = 6,
		Preformatted = 7,
	}

	internal static class BlockStyleHelper
	{
		internal static readonly float[] HeadingPointSizes = { 0f, 24f, 18f, 14f, 12f, 10f, 8f };
		internal const float ParagraphPointSize = 12f;
		internal const float PreformattedPointSize = 10f;
		internal const string PreformattedFontName = "Courier New";

		internal static readonly string[] BlockStyleNames =
		{
			"Paragraph",
			"Heading 1",
			"Heading 2",
			"Heading 3",
			"Heading 4",
			"Heading 5",
			"Heading 6",
			"Preformatted",
		};

		internal static int GetHeadingLevelFromSize(float size)
		{
			for (int i = 1; i < HeadingPointSizes.Length; i++)
			{
				if (Math.Abs(size - HeadingPointSizes[i]) < 0.5f)
					return i;
			}
			return 0;
		}

		internal static BlockStyle GetBlockStyleFromFont(float size, bool bold, string? fontName)
		{
			if (!string.IsNullOrEmpty(fontName) &&
				fontName.Equals(PreformattedFontName, StringComparison.OrdinalIgnoreCase) &&
				Math.Abs(size - PreformattedPointSize) < 0.5f)
				return BlockStyle.Preformatted;

			if (bold)
			{
				int level = GetHeadingLevelFromSize(size);
				if (level > 0) return (BlockStyle)level;
			}

			return BlockStyle.Paragraph;
		}
	}
}
