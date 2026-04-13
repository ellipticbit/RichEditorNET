using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;

using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET.Formatting
{
	internal static class MarkdownFormatter
	{
		private static readonly char[] LinkTextEscapeChars = { '[', ']' };

		internal static string ToMarkdown(ITextDocument2 doc, bool githubMarkdown) {
			var sb = new StringBuilder(4096);
			var probe = doc.Range2(0, 0);
			int storyLength;
			try {
				storyLength = probe.StoryLength;
			}
			finally {
				Marshal.ReleaseComObject(probe);
			}

			int pos = 0;

			while (pos < storyLength) {
				var paraRange = doc.Range2(pos, pos);
				try {
					paraRange.Expand(tomConstants.tomParagraph);
					int paraStart = paraRange.Start;
					int paraEnd = paraRange.End;
					int textEnd = paraEnd > paraStart ? paraEnd - 1 : paraEnd;

					var para = paraRange.Para;
					int listType;
					int listLevel;
					try {
						listType = para.ListType & 0xFFFF;
						listLevel = para.ListLevelIndex;
					}
					finally {
						Marshal.ReleaseComObject(para);
					}

					bool isList = listType != tomConstants.tomListNone;

					int headingLevel = 0;
					if (!isList && textEnd > paraStart) {
						var probeChar = doc.Range2(paraStart, paraStart + 1);
						try {
							var pFont = probeChar.Font;
							try {
								if (pFont.Bold == tomConstants.tomTrue)
									headingLevel = BlockStyleHelper.GetHeadingLevelFromSize(pFont.Size);
							}
							finally {
								Marshal.ReleaseComObject(pFont);
							}
						}
						finally {
							Marshal.ReleaseComObject(probeChar);
						}
					}

					if (headingLevel > 0) {
						for (int h = 0; h < headingLevel; h++)
							sb.Append('#');
						sb.Append(' ');
					}
					else if (isList) {
						for (int i = 0; i < listLevel; i++)
							sb.Append("    ");
						sb.Append(listType == tomConstants.tomListBullet ? "- " : "1. ");
					}

					if (textEnd > paraStart)
						EmitInlineContent(doc, sb, paraStart, textEnd, githubMarkdown, headingLevel > 0);

					sb.Append('\n');

					pos = paraEnd;
					if (pos <= paraStart) pos = paraStart + 1;
				}
				finally {
					Marshal.ReleaseComObject(paraRange);
				}
			}

			return sb.ToString().TrimEnd();
		}

		private static void EmitInlineContent(ITextDocument2 doc, StringBuilder sb, int start, int end, bool githubMarkdown, bool suppressBold = false) {
			bool bold = false;
			bool italic = false;
			bool strikethrough = false;
			string currentUrl = string.Empty;

			var runRange = doc.Range2(start, start);
			try {
				while (runRange.Start < end) {
					int moved = runRange.MoveEnd(tomConstants.tomCharFormat, 1);
					if (moved == 0) break;
					if (runRange.End > end) runRange.End = end;

					int ch = runRange.Char;

					if (ch == 0xFFFC) {
						CloseFormatting(sb, ref strikethrough, ref bold, ref italic);
						CloseLink(sb, ref currentUrl);

						int imgPos = runRange.Start;
						runRange.SetRange(imgPos, imgPos + 1);
						string altText = runRange.GetText2(tomConstants.tomTextize) ?? string.Empty;
						string imageUrl = GetCleanUrl(runRange.URL);

						sb.Append("![").Append(EscapeLinkText(altText)).Append("](").Append(imageUrl).Append(')');

						runRange.SetRange(imgPos + 1, imgPos + 1);
						continue;
					}

					var font = runRange.Font;
					bool newBold, newItalic, newStrikethrough;
					try {
						newBold = !suppressBold && font.Bold == tomConstants.tomTrue;
						newItalic = font.Italic == tomConstants.tomTrue;
						newStrikethrough = githubMarkdown && font.StrikeThrough == tomConstants.tomTrue;
					}
					finally {
						Marshal.ReleaseComObject(font);
					}

					string newUrl = GetCleanUrl(runRange.URL);

					if (currentUrl != newUrl) {
						CloseFormatting(sb, ref strikethrough, ref bold, ref italic);
						CloseLink(sb, ref currentUrl);

						if (newUrl.Length > 0) {
							currentUrl = newUrl;
							sb.Append('[');
						}
					}

					if (strikethrough && !newStrikethrough) { sb.Append("~~"); strikethrough = false; }
					if (bold && !newBold) { sb.Append("**"); bold = false; }
					if (italic && !newItalic) { sb.Append('*'); italic = false; }

					if (!italic && newItalic) { sb.Append('*'); italic = true; }
					if (!bold && newBold) { sb.Append("**"); bold = true; }
					if (!strikethrough && newStrikethrough) { sb.Append("~~"); strikethrough = true; }

					string? text = runRange.Text;
					if (text != null) {
						if (currentUrl.Length > 0) {
							text = StripHyperlinkFieldCode(text);
							if (text.Length > 0)
								sb.Append(EscapeLinkText(text));
						}
						else {
							sb.Append(text);
						}
					}

					runRange.Collapse(tomConstants.tomCollapseEnd);
				}

				CloseFormatting(sb, ref strikethrough, ref bold, ref italic);
				CloseLink(sb, ref currentUrl);
			}
			finally {
				Marshal.ReleaseComObject(runRange);
			}
		}

		private static void CloseFormatting(StringBuilder sb, ref bool strikethrough, ref bool bold, ref bool italic) {
			if (strikethrough) { sb.Append("~~"); strikethrough = false; }
			if (bold) { sb.Append("**"); bold = false; }
			if (italic) { sb.Append('*'); italic = false; }
		}

		private static void CloseLink(StringBuilder sb, ref string currentUrl) {
			if (currentUrl.Length > 0) {
				sb.Append("](").Append(currentUrl).Append(')');
				currentUrl = string.Empty;
			}
		}

		private static string GetCleanUrl(string? url) {
			if (string.IsNullOrEmpty(url)) return string.Empty;
			if (url.Length >= 2 && url[0] == '"' && url[url.Length - 1] == '"')
				return url.Substring(1, url.Length - 2);
			return url;
		}

		private static string StripHyperlinkFieldCode(string text) {
			if (!text.StartsWith("HYPERLINK \"", StringComparison.Ordinal)) return text;
			int endQuote = text.IndexOf('"', 11);
			if (endQuote < 0) return text;
			return text.Substring(endQuote + 1);
		}

		private static string EscapeLinkText(string text) {
			if (text.IndexOfAny(LinkTextEscapeChars) < 0) return text;
			return text.Replace("[", "\\[").Replace("]", "\\]");
		}

		#region - FromMarkdown -

		private struct InlineSpan
		{
			public string Text;
			public bool Bold;
			public bool Italic;
			public bool Strikethrough;
			public string Url;
		}

		private struct ParagraphBlock
		{
			public List<InlineSpan> Spans;
			public int ListType;
			public int ListLevel;
			public int HeadingLevel;
		}

		internal static void FromMarkdown(ITextDocument2 doc, string markdown, bool githubMarkdown) {
			if (markdown == null) throw new ArgumentNullException(nameof(markdown));

			doc.Freeze();
			try {
				doc.BeginEditCollection();
				try {
					var clearRange = doc.Range2(0, 0);
					try {
						int storyLength = clearRange.StoryLength;
						if (storyLength > 1) {
							clearRange.SetRange(0, storyLength);
							clearRange.Text = string.Empty;
						}
						clearRange.SetRange(0, 0);
						var resetFont = clearRange.Font;
						try {
							resetFont.Reset(tomConstants.tomDefault);
						}
						finally {
							Marshal.ReleaseComObject(resetFont);
						}
					}
					finally {
						Marshal.ReleaseComObject(clearRange);
					}

					var blocks = ParseBlocks(markdown, githubMarkdown);
					InsertBlocks(doc, blocks);
				}
				finally {
					doc.EndEditCollection();
				}
			}
			finally {
				doc.Unfreeze();
			}
		}

		private static int TryParseHeadingPrefix(string line) {
			int i = 0;
			while (i < line.Length && line[i] == '#') i++;
			if (i == 0 || i > 6) return 0;
			if (i < line.Length && line[i] == ' ') return i;
			if (i == line.Length) return i;
			return 0;
		}

		private static List<ParagraphBlock> ParseBlocks(string markdown, bool githubMarkdown) {
			var blocks = new List<ParagraphBlock>();
			var lines = markdown.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');

			for (int i = 0; i < lines.Length; i++) {
				string line = lines[i];

				if (line.Length == 0) {
					blocks.Add(new ParagraphBlock {
						Spans = new List<InlineSpan>(),
						ListType = tomConstants.tomListNone,
						ListLevel = 0,
						HeadingLevel = 0
					});
					continue;
				}

				int indent = 0;
				int pos = 0;
				while (pos < line.Length && (line[pos] == ' ' || line[pos] == '\t')) {
					indent += line[pos] == '\t' ? 4 : 1;
					pos++;
				}
				string trimmed = line.Substring(pos);
				int listLevel = indent / 4;

				int headingLevel = 0;
				if (listLevel == 0) {
					headingLevel = TryParseHeadingPrefix(trimmed);
				}

				if (headingLevel > 0) {
					string content = headingLevel < trimmed.Length ? trimmed.Substring(headingLevel + 1) : string.Empty;
					blocks.Add(new ParagraphBlock {
						Spans = ParseInline(content, githubMarkdown),
						ListType = tomConstants.tomListNone,
						ListLevel = 0,
						HeadingLevel = headingLevel
					});
					continue;
				}

				int listType = tomConstants.tomListNone;
				string listContent = trimmed;

				if (trimmed.Length >= 2 && (trimmed[0] == '-' || trimmed[0] == '*' || trimmed[0] == '+') && trimmed[1] == ' ') {
					listType = tomConstants.tomListBullet;
					listContent = trimmed.Substring(2);
				}
				else {
					int orderedLen = TryParseOrderedPrefix(trimmed);
					if (orderedLen > 0) {
						listType = tomConstants.tomListNumberAsArabic;
						listContent = trimmed.Substring(orderedLen);
					}
				}

				blocks.Add(new ParagraphBlock {
					Spans = ParseInline(listContent, githubMarkdown),
					ListType = listType,
					ListLevel = listLevel,
					HeadingLevel = 0
				});
			}

			return blocks;
		}

		private static int TryParseOrderedPrefix(string line) {
			int i = 0;
			while (i < line.Length && char.IsDigit(line[i])) i++;
			if (i == 0 || i >= line.Length) return 0;
			if (line[i] == '.' && i + 1 < line.Length && line[i + 1] == ' ')
				return i + 2;
			if (line[i] == ')' && i + 1 < line.Length && line[i + 1] == ' ')
				return i + 2;
			return 0;
		}

		private static List<InlineSpan> ParseInline(string text, bool githubMarkdown) {
			var spans = new List<InlineSpan>();
			ParseInlineCore(text, 0, text.Length, false, false, false, string.Empty, githubMarkdown, spans);
			return spans;
		}

		private static void ParseInlineCore(string text, int start, int end, bool bold, bool italic, bool strikethrough, string url, bool githubMarkdown, List<InlineSpan> spans) {
			var buf = new StringBuilder();
			int i = start;

			while (i < end) {
				char c = text[i];

				if (c == '\\' && i + 1 < end && IsMarkdownEscapable(text[i + 1])) {
					buf.Append(text[i + 1]);
					i += 2;
					continue;
				}

				if (c == '~' && githubMarkdown && i + 1 < end && text[i + 1] == '~') {
					int closeIdx = text.IndexOf("~~", i + 2, StringComparison.Ordinal);
					if (closeIdx >= 0 && closeIdx < end) {
						FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
						ParseInlineCore(text, i + 2, closeIdx, bold, italic, true, url, githubMarkdown, spans);
						i = closeIdx + 2;
						continue;
					}
				}

				if (c == '*') {
					int starCount = CountRun(text, i, end, '*');
					if (starCount >= 3) {
						string closeToken = "***";
						int closeIdx = text.IndexOf(closeToken, i + 3, StringComparison.Ordinal);
						if (closeIdx >= 0 && closeIdx < end) {
							FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
							ParseInlineCore(text, i + 3, closeIdx, true, true, strikethrough, url, githubMarkdown, spans);
							i = closeIdx + 3;
							continue;
						}
					}
					if (starCount >= 2) {
						string closeToken = "**";
						int closeIdx = text.IndexOf(closeToken, i + 2, StringComparison.Ordinal);
						if (closeIdx >= 0 && closeIdx < end) {
							FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
							ParseInlineCore(text, i + 2, closeIdx, true, italic, strikethrough, url, githubMarkdown, spans);
							i = closeIdx + 2;
							continue;
						}
					}
					if (starCount >= 1) {
						int closeIdx = text.IndexOf('*', i + 1);
						if (closeIdx >= 0 && closeIdx < end) {
							FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
							ParseInlineCore(text, i + 1, closeIdx, bold, true, strikethrough, url, githubMarkdown, spans);
							i = closeIdx + 1;
							continue;
						}
					}
					buf.Append(c);
					i++;
					continue;
				}

				if (c == '!' && i + 1 < end && text[i + 1] == '[') {
					string? altText = ReadBracketContent(text, i + 1, end);
					if (altText != null) {
						int afterBracket = i + 2 + altText.Length + 1;
						string? imgUrl = ReadParenContent(text, afterBracket, end);
						if (imgUrl != null) {
							FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
							string displayText = altText.Length > 0 ? altText : imgUrl;
							spans.Add(new InlineSpan {
								Text = displayText,
								Bold = false,
								Italic = false,
								Strikethrough = false,
								Url = imgUrl
							});
							i = afterBracket + 1 + imgUrl.Length + 1;
							continue;
						}
					}
				}

				if (c == '[') {
					string? linkText = ReadBracketContent(text, i, end);
					if (linkText != null) {
						int afterBracket = i + 1 + linkText.Length + 1;
						string? linkUrl = ReadParenContent(text, afterBracket, end);
						if (linkUrl != null) {
							FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
							ParseInlineCore(linkText, 0, linkText.Length, bold, italic, strikethrough, linkUrl, githubMarkdown, spans);
							i = afterBracket + 1 + linkUrl.Length + 1;
							continue;
						}
					}
				}

				buf.Append(c);
				i++;
			}

			FlushSpanBuffer(buf, bold, italic, strikethrough, url, spans);
		}

		private static void FlushSpanBuffer(StringBuilder buf, bool bold, bool italic, bool strikethrough, string url, List<InlineSpan> spans) {
			if (buf.Length == 0) return;
			spans.Add(new InlineSpan {
				Text = buf.ToString(),
				Bold = bold,
				Italic = italic,
				Strikethrough = strikethrough,
				Url = url ?? string.Empty
			});
			buf.Clear();
		}

		private static string? ReadBracketContent(string text, int pos, int end) {
			if (pos >= end || text[pos] != '[') return null;
			int depth = 1;
			int start = pos + 1;
			int i = start;
			while (i < end && depth > 0) {
				if (text[i] == '\\' && i + 1 < end) { i += 2; continue; }
				if (text[i] == '[') depth++;
				else if (text[i] == ']') depth--;
				if (depth > 0) i++;
			}
			if (depth != 0) return null;
			return text.Substring(start, i - start);
		}

		private static string? ReadParenContent(string text, int pos, int end) {
			if (pos >= end || text[pos] != '(') return null;
			int depth = 1;
			int start = pos + 1;
			int i = start;
			while (i < end && depth > 0) {
				if (text[i] == '\\' && i + 1 < end) { i += 2; continue; }
				if (text[i] == '(') depth++;
				else if (text[i] == ')') depth--;
				if (depth > 0) i++;
			}
			if (depth != 0) return null;
			return text.Substring(start, i - start);
		}

		private static int CountRun(string text, int pos, int end, char ch) {
			int count = 0;
			while (pos + count < end && text[pos + count] == ch) count++;
			return count;
		}

		private static bool IsMarkdownEscapable(char c) {
			return c == '\\' || c == '*' || c == '_' || c == '[' || c == ']'
				|| c == '(' || c == ')' || c == '!' || c == '~' || c == '`'
				|| c == '#' || c == '+' || c == '-' || c == '.' || c == '|';
		}

		private static void InsertBlocks(ITextDocument2 doc, List<ParagraphBlock> blocks) {
			if (blocks.Count == 0) return;

			var range = doc.Range2(0, 0);
			try {
				for (int b = 0; b < blocks.Count; b++) {
					var block = blocks[b];

					if (b > 0) {
						int beforePara = range.Start;
						range.Text = "\r";
						range.SetRange(range.End, range.End);

						var prevParaRange = doc.Range2(beforePara, beforePara);
						try {
							var prevPara = prevParaRange.Para;
							try {
								var prevBlock = blocks[b - 1];
								prevPara.ListType = prevBlock.ListType != tomConstants.tomListNone
									? (prevBlock.ListType == tomConstants.tomListBullet
										? prevBlock.ListType
										: prevBlock.ListType | tomConstants.tomListPeriod)
									: tomConstants.tomListNone;
								prevPara.ListLevelIndex = prevBlock.ListLevel;
								if (prevBlock.ListType != tomConstants.tomListNone)
									prevPara.SetIndents(0, prevBlock.ListLevel * 18f, 0);
							}
							finally {
								Marshal.ReleaseComObject(prevPara);
							}
						}
						finally {
							Marshal.ReleaseComObject(prevParaRange);
						}
					}

					int blockStart = range.Start;
					InsertSpans(doc, range, block.Spans);
					int blockEnd = range.Start;

					if (block.HeadingLevel > 0 && blockEnd > blockStart) {
						var headingRange = doc.Range2(blockStart, blockEnd);
						try {
							var hFont = headingRange.Font;
							try {
								hFont.Bold = tomConstants.tomTrue;
								hFont.Size = BlockStyleHelper.HeadingPointSizes[block.HeadingLevel];
							}
							finally {
								Marshal.ReleaseComObject(hFont);
							}
						}
						finally {
							Marshal.ReleaseComObject(headingRange);
						}

						var resetRange = doc.Range2(blockEnd, blockEnd);
						try {
							var rFont = resetRange.Font;
							try {
								rFont.Bold = tomConstants.tomFalse;
								rFont.Size = BlockStyleHelper.ParagraphPointSize;
							}
							finally {
								Marshal.ReleaseComObject(rFont);
							}
						}
						finally {
							Marshal.ReleaseComObject(resetRange);
						}
					}
				}

				var lastBlock = blocks[blocks.Count - 1];
				var lastParaRange = doc.Range2(range.Start, range.Start);
				try {
					var lastPara = lastParaRange.Para;
					try {
						lastPara.ListType = lastBlock.ListType != tomConstants.tomListNone
							? (lastBlock.ListType == tomConstants.tomListBullet
								? lastBlock.ListType
								: lastBlock.ListType | tomConstants.tomListPeriod)
							: tomConstants.tomListNone;
						lastPara.ListLevelIndex = lastBlock.ListLevel;
						if (lastBlock.ListType != tomConstants.tomListNone)
							lastPara.SetIndents(0, lastBlock.ListLevel * 18f, 0);
					}
					finally {
						Marshal.ReleaseComObject(lastPara);
					}
				}
				finally {
					Marshal.ReleaseComObject(lastParaRange);
				}
			}
			finally {
				Marshal.ReleaseComObject(range);
			}
		}

		private static void InsertSpans(ITextDocument2 doc, ITextRange2 range, List<InlineSpan> spans) {
			for (int s = 0; s < spans.Count; s++) {
				var span = spans[s];
				if (string.IsNullOrEmpty(span.Text)) continue;

				int insertStart = range.Start;
					range.Text = span.Text;
					int insertEnd = insertStart + span.Text.Length;
					range.SetRange(insertEnd, insertEnd);

				if (span.Bold || span.Italic || span.Strikethrough || !string.IsNullOrEmpty(span.Url)) {
					var fmtRange = doc.Range2(insertStart, insertEnd);
					try {
						if (span.Bold || span.Italic || span.Strikethrough) {
							var font = fmtRange.Font;
							try {
								if (span.Bold) font.Bold = tomConstants.tomTrue;
								if (span.Italic) font.Italic = tomConstants.tomTrue;
								if (span.Strikethrough) font.StrikeThrough = tomConstants.tomTrue;
							}
							finally {
								Marshal.ReleaseComObject(font);
							}
						}

						if (!string.IsNullOrEmpty(span.Url)) {
							fmtRange.URL = "\"" + span.Url + "\"";
						}
					}
					finally {
						Marshal.ReleaseComObject(fmtRange);
					}
				}
			}
		}

		#endregion
	}
}
