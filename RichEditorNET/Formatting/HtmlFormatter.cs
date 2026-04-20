using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.Net;
using System.Runtime.InteropServices;
using System.Text;

using EllipticBit.RichEditorNET.Support;
using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET.Formatting
{
	internal static class HtmlFormatter
	{
		internal static string ToHtml(ITextDocument2 doc, IntPtr hwnd, bool htmlFontSizing, bool completeDocument = false, string defaultFontName = "", float defaultFontSize = 0f) {
			var sb = new StringBuilder(4096);
			if (completeDocument) sb.Append("<html><body>");
			var probe = doc.Range2(0, 0);
			int storyLength;
			try {
				storyLength = probe.StoryLength;
			}
			finally {
				Marshal.ReleaseComObject(probe);
			}

			IRichEditOle richOle = null;
			if (hwnd != IntPtr.Zero) {
				IntPtr pOle = IntPtr.Zero;
				PInvoke.SendMessage(hwnd, PInvoke.EM_GETOLEINTERFACE, IntPtr.Zero, ref pOle);
				if (pOle != IntPtr.Zero) {
					try { richOle = Marshal.GetObjectForIUnknown(pOle) as IRichEditOle; }
					catch { richOle = null; }
					Marshal.Release(pOle);
				}
			}

			try {
			int pos = 0;
			var listTypeStack = new List<int>();
			var liOpenStack = new List<bool>();

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
					int alignment;
					try {
						listType = para.ListType & 0xFFFF;
						listLevel = para.ListLevelIndex;
						alignment = para.Alignment;
						if (alignment < 0 || alignment > 3) alignment = tomConstants.tomAlignLeft;
					}
					finally {
						Marshal.ReleaseComObject(para);
					}

					bool isList = listType != tomConstants.tomListNone;

					int headingLevel = 0;
					bool preformatted = false;
					if (!isList && textEnd > paraStart) {
						var probeChar = doc.Range2(paraStart, paraStart + 1);
						try {
							var pFont = probeChar.Font;
							try {
								var blockStyle = BlockStyleHelper.GetBlockStyleFromFont(pFont.Size, pFont.Bold == tomConstants.tomTrue, pFont.Name);
								if (blockStyle != BlockStyle.Paragraph && blockStyle != BlockStyle.Preformatted)
									headingLevel = (int)blockStyle;
								else if (blockStyle == BlockStyle.Preformatted)
									preformatted = true;
							}
							finally {
								Marshal.ReleaseComObject(pFont);
							}
						}
						finally {
							Marshal.ReleaseComObject(probeChar);
						}
					}

					if (isList) {
						int targetDepth = listLevel + 1;

						while (listTypeStack.Count > targetDepth) {
							int top = listTypeStack.Count - 1;
							if (liOpenStack[top]) sb.Append("</li>");
							sb.Append(GetEmitListCloseTag(listTypeStack[top]));
							listTypeStack.RemoveAt(top);
							liOpenStack.RemoveAt(top);
						}

						if (listTypeStack.Count == targetDepth) {
							int top = listTypeStack.Count - 1;
							if (listTypeStack[top] != listType) {
								if (liOpenStack[top]) sb.Append("</li>");
								sb.Append(GetEmitListCloseTag(listTypeStack[top]));
								listTypeStack.RemoveAt(top);
								liOpenStack.RemoveAt(top);
							}
							else if (liOpenStack[top]) {
								sb.Append("</li>");
								liOpenStack[top] = false;
							}
						}

						while (listTypeStack.Count < targetDepth) {
							if (listTypeStack.Count > 0 && !liOpenStack[listTypeStack.Count - 1]) {
								sb.Append("<li>");
								liOpenStack[listTypeStack.Count - 1] = true;
							}
							sb.Append(GetEmitListOpenTag(listType));
							listTypeStack.Add(listType);
							liOpenStack.Add(false);
						}

						sb.Append("<li");
						if (alignment != tomConstants.tomAlignLeft)
							sb.Append(" style=\"text-align: ").Append(AlignmentToCss(alignment)).Append("\"");
						sb.Append('>');
						liOpenStack[liOpenStack.Count - 1] = true;

						if (textEnd > paraStart)
							EmitHtmlInlineContent(doc, richOle, sb, paraStart, textEnd, !htmlFontSizing, false, defaultFontName, defaultFontSize);
					}
					else {
						CloseAllLists(sb, listTypeStack, liOpenStack);

						string tag = "p";
						if (headingLevel > 0) tag = "h" + headingLevel;
						else if (preformatted) tag = "pre";

						sb.Append('<').Append(tag);
						if (alignment != tomConstants.tomAlignLeft)
							sb.Append(" style=\"text-align: ").Append(AlignmentToCss(alignment)).Append("\"");
						sb.Append('>');

						string baseFont = defaultFontName;
						float baseSize = defaultFontSize;
						if (headingLevel > 0) {
							baseSize = BlockStyleHelper.HeadingPointSizes[headingLevel];
						} else if (preformatted) {
							baseFont = BlockStyleHelper.PreformattedFontName;
							baseSize = BlockStyleHelper.PreformattedPointSize;
						}
						if (textEnd > paraStart)
							EmitHtmlInlineContent(doc, richOle, sb, paraStart, textEnd, !htmlFontSizing, headingLevel > 0, baseFont, baseSize);
						sb.Append("</").Append(tag).Append('>');
					}

					pos = paraEnd;
					if (pos <= paraStart) pos = paraStart + 1;
				}
				finally {
					Marshal.ReleaseComObject(paraRange);
				}
			}

			CloseAllLists(sb, listTypeStack, liOpenStack);
			if (completeDocument) sb.Append("</body></html>");
			return sb.ToString();
			}
			finally {
				if (richOle != null) Marshal.ReleaseComObject(richOle);
			}
		}

		#region - HTML Emission -

		private static void EmitHtmlInlineContent(ITextDocument2 doc, IRichEditOle richOle, StringBuilder sb, int start, int end, bool emitFontStyles, bool suppressBold, string defaultFontName, float defaultFontSize) {
			bool curBold = false, curItalic = false, curStrike = false, curSup = false, curSub = false;
			int curUnderline = 0, curFore = tomConstants.tomAutoColor, curBack = tomConstants.tomAutoColor;
			string curUrl = string.Empty;
			string curFontName = string.Empty;
			float curFontSize = 0f;
			bool tagsOpen = false;

			var runRange = doc.Range2(start, start);
			try {
				while (runRange.Start < end) {
					int moved = runRange.MoveEnd(tomConstants.tomCharFormat, 1);
					if (moved == 0) break;
					if (runRange.End > end) runRange.End = end;

					var font = runRange.Font;
					bool newBold, newItalic, newStrike, newSup, newSub;
					int newUnderline, newFore, newBack;
					string newFontName = string.Empty;
					float newFontSize = 0f;
					try {
						newBold = !suppressBold && font.Bold == tomConstants.tomTrue;
						newItalic = font.Italic == tomConstants.tomTrue;
						newStrike = font.StrikeThrough == tomConstants.tomTrue;
						newSup = font.Superscript == tomConstants.tomTrue;
						newSub = font.Subscript == tomConstants.tomTrue;
						newUnderline = font.Underline;
						newFore = font.ForeColor;
						newBack = font.BackColor;
						if (emitFontStyles) {
							newFontName = font.Name ?? string.Empty;
							if (newFontName.Equals(defaultFontName, StringComparison.OrdinalIgnoreCase))
								newFontName = string.Empty;
							newFontSize = font.Size;
							if (newFontSize < 0) newFontSize = 0f;
							if (defaultFontSize > 0 && Math.Abs(newFontSize - defaultFontSize) < 0.5f)
								newFontSize = 0f;
						}
					}
					finally {
						Marshal.ReleaseComObject(font);
					}

					if (newUnderline <= 0)
						newUnderline = 0;
					else if (newUnderline != (int)UnderlineStyle.Single &&
						newUnderline != (int)UnderlineStyle.Double &&
						newUnderline != (int)UnderlineStyle.Dotted &&
						newUnderline != (int)UnderlineStyle.Dashed &&
						newUnderline != (int)UnderlineStyle.Wavy)
						newUnderline = (int)UnderlineStyle.Single;

					string newUrl = GetEmitCleanUrl(runRange.URL);

					string text = runRange.Text ?? string.Empty;
					int runDocStart = runRange.Start;
					int docOffset = 0;
					if (text.StartsWith("HYPERLINK \"", StringComparison.Ordinal)) {
						int endQuote = text.IndexOf('"', 11);
						if (endQuote >= 0) {
							docOffset = endQuote + 1;
							text = text.Substring(docOffset);
						}
					}

					bool formatChanged = newBold != curBold || newItalic != curItalic || newStrike != curStrike ||
						newUnderline != curUnderline || newSup != curSup || newSub != curSub ||
						newFore != curFore || newBack != curBack || newUrl != curUrl ||
						newFontName != curFontName || Math.Abs(newFontSize - curFontSize) > 0.01f;

					if (formatChanged && tagsOpen) {
						CloseHtmlInlineTags(sb, curBold, curItalic, curStrike, curUnderline, curSup, curSub, curFore, curBack, curUrl, curFontName, curFontSize);
						tagsOpen = false;
					}

					curBold = newBold; curItalic = newItalic; curStrike = newStrike;
					curUnderline = newUnderline; curSup = newSup; curSub = newSub;
					curFore = newFore; curBack = newBack; curUrl = newUrl;
					curFontName = newFontName; curFontSize = newFontSize;

					bool hasFormatting = curBold || curItalic || curStrike || curUnderline != 0 ||
						curSup || curSub || curFore != tomConstants.tomAutoColor || curBack != tomConstants.tomAutoColor ||
						curUrl.Length > 0 || NeedsStyleSpan(curFore, curBack, curUnderline, curFontName, curFontSize);

					int segStart = 0;
					for (int idx = 0; idx <= text.Length; idx++) {
						bool atImage = idx < text.Length && text[idx] == '\uFFFC';
						bool atEndOfText = idx == text.Length;
						if (!atImage && !atEndOfText) continue;

						if (idx > segStart) {
							if (hasFormatting && !tagsOpen) {
								OpenHtmlInlineTags(sb, curBold, curItalic, curStrike, curUnderline, curSup, curSub, curFore, curBack, curUrl, curFontName, curFontSize);
								tagsOpen = true;
							}
							sb.Append(WebUtility.HtmlEncode(text.Substring(segStart, idx - segStart)));
						}

						if (atImage) {
							if (tagsOpen) {
								CloseHtmlInlineTags(sb, curBold, curItalic, curStrike, curUnderline, curSup, curSub, curFore, curBack, curUrl, curFontName, curFontSize);
								tagsOpen = false;
							}
							EmitImageTag(doc, richOle, sb, runDocStart + docOffset + idx);
							segStart = idx + 1;
						}
					}

					runRange.Collapse(tomConstants.tomCollapseEnd);
				}

				if (tagsOpen)
					CloseHtmlInlineTags(sb, curBold, curItalic, curStrike, curUnderline, curSup, curSub, curFore, curBack, curUrl, curFontName, curFontSize);
			}
			finally {
				Marshal.ReleaseComObject(runRange);
			}
		}

		private static void OpenHtmlInlineTags(StringBuilder sb, bool bold, bool italic, bool strike,
			int underline, bool sup, bool sub, int foreColor, int backColor, string url,
			string fontName, float fontSize) {
			if (url.Length > 0)
				sb.Append("<a href=\"").Append(WebUtility.HtmlEncode(url)).Append("\">");
			if (bold) sb.Append("<b>");
			if (italic) sb.Append("<i>");
			if (strike) sb.Append("<s>");
			if (underline == (int)UnderlineStyle.Single) sb.Append("<u>");
			if (sup) sb.Append("<sup>");
			if (sub) sb.Append("<sub>");
			if (NeedsStyleSpan(foreColor, backColor, underline, fontName, fontSize))
				sb.Append("<span style=\"").Append(BuildSpanStyle(foreColor, backColor, underline, fontName, fontSize)).Append("\">");
		}

		private static void CloseHtmlInlineTags(StringBuilder sb, bool bold, bool italic, bool strike,
			int underline, bool sup, bool sub, int foreColor, int backColor, string url,
			string fontName, float fontSize) {
			if (NeedsStyleSpan(foreColor, backColor, underline, fontName, fontSize)) sb.Append("</span>");
			if (sub) sb.Append("</sub>");
			if (sup) sb.Append("</sup>");
			if (underline == (int)UnderlineStyle.Single) sb.Append("</u>");
			if (strike) sb.Append("</s>");
			if (italic) sb.Append("</i>");
			if (bold) sb.Append("</b>");
			if (url.Length > 0) sb.Append("</a>");
		}

		private static bool NeedsStyleSpan(int foreColor, int backColor, int underline, string fontName, float fontSize) {
			return foreColor != tomConstants.tomAutoColor || backColor != tomConstants.tomAutoColor ||
				(underline > 0 && underline != (int)UnderlineStyle.Single) ||
				fontName.Length > 0 ||
				fontSize > 0;
		}

		private static string BuildSpanStyle(int foreColor, int backColor, int underline, string fontName, float fontSize) {
			var sb = new StringBuilder();
			if (fontName.Length > 0)
				sb.Append("font-family: '").Append(fontName).Append("'");
			if (fontSize > 0) {
				if (sb.Length > 0) sb.Append("; ");
				sb.Append("font-size: ").Append(fontSize.ToString("0.##", CultureInfo.InvariantCulture)).Append("pt");
			}
			if (foreColor != tomConstants.tomAutoColor) {
				if (sb.Length > 0) sb.Append("; ");
				sb.Append("color: ").Append(OleColorToCss(foreColor));
			}
			if (backColor != tomConstants.tomAutoColor) {
				if (sb.Length > 0) sb.Append("; ");
				sb.Append("background-color: ").Append(OleColorToCss(backColor));
			}
			if (underline > 0 && underline != (int)UnderlineStyle.Single) {
				if (sb.Length > 0) sb.Append("; ");
				sb.Append("text-decoration: underline; text-decoration-style: ").Append(GetCssTextDecorationStyle(underline));
			}
			return sb.ToString();
		}

		private static string GetCssTextDecorationStyle(int underlineValue) {
			switch (underlineValue) {
				case (int)UnderlineStyle.Double: return "double";
				case (int)UnderlineStyle.Dotted: return "dotted";
				case (int)UnderlineStyle.Dashed: return "dashed";
				case (int)UnderlineStyle.Wavy: return "wavy";
				default: return "solid";
			}
		}

		private static string OleColorToCss(int oleColor) {
			var c = ColorTranslator.FromOle(oleColor);
			return "#" + c.R.ToString("x2") + c.G.ToString("x2") + c.B.ToString("x2");
		}

		private static string AlignmentToCss(int alignment) {
			switch (alignment) {
				case tomConstants.tomAlignCenter: return "center";
				case tomConstants.tomAlignRight: return "right";
				case tomConstants.tomAlignJustify: return "justify";
				default: return "left";
			}
		}

		private static string GetEmitListOpenTag(int listType) {
			switch (listType) {
				case tomConstants.tomListBullet: return "<ul>";
				case tomConstants.tomListNumberAsLCLetter: return "<ol type=\"a\">";
				case tomConstants.tomListNumberAsUCLetter: return "<ol type=\"A\">";
				case tomConstants.tomListNumberAsLCRoman: return "<ol type=\"i\">";
				case tomConstants.tomListNumberAsUCRoman: return "<ol type=\"I\">";
				default: return "<ol>";
			}
		}

		private static string GetEmitListCloseTag(int listType) {
			return listType == tomConstants.tomListBullet ? "</ul>" : "</ol>";
		}

		private static void CloseAllLists(StringBuilder sb, List<int> listTypeStack, List<bool> liOpenStack) {
			while (listTypeStack.Count > 0) {
				int top = listTypeStack.Count - 1;
				if (liOpenStack[top]) sb.Append("</li>");
				sb.Append(GetEmitListCloseTag(listTypeStack[top]));
				listTypeStack.RemoveAt(top);
				liOpenStack.RemoveAt(top);
			}
		}

		private static string GetEmitCleanUrl(string url) {
			if (string.IsNullOrEmpty(url)) return string.Empty;
			if (url.Length >= 2 && url[0] == '"' && url[url.Length - 1] == '"')
				return url.Substring(1, url.Length - 2);
			return url;
		}

		private static string StripEmitHyperlinkFieldCode(string text) {
			if (!text.StartsWith("HYPERLINK \"", StringComparison.Ordinal)) return text;
			int endQuote = text.IndexOf('"', 11);
			if (endQuote < 0) return text;
			return text.Substring(endQuote + 1);
		}

		private static void EmitImageTag(ITextDocument2 doc, IRichEditOle richOle, StringBuilder sb, int imgPos) {
			var imgRange = doc.Range2(imgPos, imgPos + 1);
			string url;
			try {
				url = GetEmitCleanUrl(imgRange.URL);
			}
			finally {
				Marshal.ReleaseComObject(imgRange);
			}

			byte[] imageData = null;
			string altText = string.Empty;
			int widthHimetric = 0, heightHimetric = 0;

			if (richOle != null) {
				if (ImageOleObject.TryRead(richOle, imgPos, out var bytes, out var alt, out widthHimetric, out heightHimetric)) {
					imageData = bytes;
					altText = alt ?? string.Empty;
				}
			}

			if (imageData == null || imageData.Length == 0) return;

			int widthPx = HimetricToPixels(widthHimetric);
			int heightPx = HimetricToPixels(heightHimetric);
			byte[] emit = ConvertImageBytesToPng(imageData, ref widthPx, ref heightPx) ?? imageData;

			string mimeType = GetEmitImageMimeType(emit);
			string base64 = ToWebSafeBase64(emit);

			if (url.Length > 0)
				sb.Append("<a href=\"").Append(WebUtility.HtmlEncode(url)).Append("\">");

			sb.Append("<img src=\"data:").Append(mimeType).Append(";base64,").Append(base64).Append('"');
			if (altText.Length > 0)
				sb.Append(" alt=\"").Append(WebUtility.HtmlEncode(altText)).Append('"');
			if (widthPx > 0)
				sb.Append(" width=\"").Append(widthPx).Append('"');
			if (heightPx > 0)
				sb.Append(" height=\"").Append(heightPx).Append('"');
			sb.Append(" />");

			if (url.Length > 0)
				sb.Append("</a>");
		}

		private static byte[] ConvertImageBytesToPng(byte[] data, ref int widthPx, ref int heightPx) {
			try {
				using (var ms = new MemoryStream(data))
				using (var image = Image.FromStream(ms)) {
					if (widthPx <= 0) widthPx = image.Width;
					if (heightPx <= 0) heightPx = image.Height;
					var guid = image.RawFormat.Guid;
					if (guid == ImageFormat.Png.Guid || guid == ImageFormat.Jpeg.Guid || guid == ImageFormat.Gif.Guid)
						return data;
					using (var outMs = new MemoryStream()) {
						image.Save(outMs, ImageFormat.Png);
						return outMs.ToArray();
					}
				}
			}
			catch { return null; }
		}

		private static int HimetricToPixels(int himetric) {
			return (int)((long)himetric * 96 / 2540);
		}

		private static string GetEmitImageMimeType(byte[] data) {
			if (data.Length >= 4) {
				if (data[0] == 0x89 && data[1] == 0x50 && data[2] == 0x4E && data[3] == 0x47)
					return "image/png";
				if (data[0] == 0xFF && data[1] == 0xD8 && data[2] == 0xFF)
					return "image/jpeg";
				if (data[0] == 0x47 && data[1] == 0x49 && data[2] == 0x46)
					return "image/gif";
			}
			return "image/png";
		}

		private static string ToWebSafeBase64(byte[] data) {
			return Convert.ToBase64String(data).Replace('+', '-').Replace('/', '_').TrimEnd('=');
		}

		#endregion

		internal static void FromHtml(ITextDocument2 doc, IRichEditOle richOle, string html, bool strict = false) {
			if (html == null) throw new ArgumentNullException(nameof(html));

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

					var tokens = Tokenize(html);
					var blocks = BuildBlocks(tokens, strict);
					InsertBlocks(doc, richOle, blocks);
				}
				finally {
					doc.EndEditCollection();
				}
			}
			finally {
				doc.Unfreeze();
			}
		}

		#region - Data Types -

		private enum TokenType { Text, OpenTag, CloseTag, SelfClosingTag }

		private struct HtmlToken {
			public TokenType Type;
			public string TagName;
			public string Text;
			public Dictionary<string, string> Attributes;
		}

		private struct StyleOverride {
			public bool? Bold;
			public bool? Italic;
			public int? Underline;
			public bool? StrikeThrough;
			public bool? Superscript;
			public bool? Subscript;
			public int? ForeColor;
			public int? BackColor;
			public string Url;
			public int? Alignment;
			public string FontName;
			public float? FontSize;
		}

		private struct StyleEntry {
			public string TagName;
			public StyleOverride Override;
		}

		private struct EffectiveStyle {
			public bool Bold;
			public bool Italic;
			public int Underline;
			public bool StrikeThrough;
			public bool Superscript;
			public bool Subscript;
			public int ForeColor;
			public int BackColor;
			public string Url;
			public int Alignment;
			public string FontName;
			public float FontSize;
		}

		private struct InlineSpan {
			public string Text;
			public bool Bold;
			public bool Italic;
			public int Underline;
			public bool StrikeThrough;
			public bool Superscript;
			public bool Subscript;
			public int ForeColor;
			public int BackColor;
			public string Url;
			public string FontName;
			public float FontSize;
			public byte[] ImageData;
			public string ImageAltText;
			public int ImageWidth;
			public int ImageHeight;
		}

		private struct ParagraphBlock {
			public List<InlineSpan> Spans;
			public int ListType;
			public int ListLevel;
			public int HeadingLevel;
			public bool Preformatted;
			public int Alignment;
		}

		#endregion

		#region - Tokenizer -

		private static readonly HashSet<string> VoidElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
			"br", "hr", "img", "input", "meta", "link", "area", "base", "col", "embed", "source", "track", "wbr"
		};

		private static List<HtmlToken> Tokenize(string html) {
			var tokens = new List<HtmlToken>();
			int i = 0;
			var textBuf = new StringBuilder();

			while (i < html.Length) {
				if (html[i] == '<') {
					if (textBuf.Length > 0) {
						tokens.Add(new HtmlToken { Type = TokenType.Text, Text = WebUtility.HtmlDecode(textBuf.ToString()) });
						textBuf.Clear();
					}

					if (i + 3 < html.Length && html[i + 1] == '!' && html[i + 2] == '-' && html[i + 3] == '-') {
						int end = html.IndexOf("-->", i + 4, StringComparison.Ordinal);
						i = end >= 0 ? end + 3 : html.Length;
						continue;
					}

					if (i + 1 < html.Length && html[i + 1] == '!') {
						int end = html.IndexOf('>', i + 2);
						i = end >= 0 ? end + 1 : html.Length;
						continue;
					}

					int tagEnd = html.IndexOf('>', i + 1);
					if (tagEnd < 0) { textBuf.Append('<'); i++; continue; }

					string tagContent = html.Substring(i + 1, tagEnd - i - 1);
					i = tagEnd + 1;

					string trimmed = tagContent.Trim();
					if (trimmed.Length == 0) continue;

					if (trimmed[0] == '/') {
						string name = trimmed.Substring(1).Trim().ToLowerInvariant();
						if (name.Length > 0)
							tokens.Add(new HtmlToken { Type = TokenType.CloseTag, TagName = name });
						continue;
					}

					bool isSelfClose = trimmed[trimmed.Length - 1] == '/';
					if (isSelfClose)
						trimmed = trimmed.Substring(0, trimmed.Length - 1).Trim();

					ParseTagContent(trimmed, out string tagName, out var attrs);

					if (!isSelfClose && VoidElements.Contains(tagName))
						isSelfClose = true;

					tokens.Add(new HtmlToken {
						Type = isSelfClose ? TokenType.SelfClosingTag : TokenType.OpenTag,
						TagName = tagName,
						Attributes = attrs
					});
				}
				else {
					textBuf.Append(html[i]);
					i++;
				}
			}

			if (textBuf.Length > 0)
				tokens.Add(new HtmlToken { Type = TokenType.Text, Text = WebUtility.HtmlDecode(textBuf.ToString()) });

			return tokens;
		}

		private static void ParseTagContent(string content, out string tagName, out Dictionary<string, string> attributes) {
			attributes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			int i = 0;
			while (i < content.Length && !char.IsWhiteSpace(content[i])) i++;
			tagName = content.Substring(0, i).ToLowerInvariant();

			while (i < content.Length) {
				while (i < content.Length && char.IsWhiteSpace(content[i])) i++;
				if (i >= content.Length) break;

				int nameStart = i;
				while (i < content.Length && content[i] != '=' && !char.IsWhiteSpace(content[i])) i++;
				string attrName = content.Substring(nameStart, i - nameStart);
				if (attrName.Length == 0) break;

				while (i < content.Length && char.IsWhiteSpace(content[i])) i++;

				if (i < content.Length && content[i] == '=') {
					i++;
					while (i < content.Length && char.IsWhiteSpace(content[i])) i++;

					string attrValue;
					if (i < content.Length && (content[i] == '"' || content[i] == '\'')) {
						char quote = content[i];
						i++;
						int valueStart = i;
						while (i < content.Length && content[i] != quote) i++;
						attrValue = content.Substring(valueStart, i - valueStart);
						if (i < content.Length) i++;
					}
					else {
						int valueStart = i;
						while (i < content.Length && !char.IsWhiteSpace(content[i])) i++;
						attrValue = content.Substring(valueStart, i - valueStart);
					}

					attributes[attrName] = WebUtility.HtmlDecode(attrValue);
				}
				else {
					attributes[attrName] = string.Empty;
				}
			}
		}

		#endregion

		#region - Block Builder -

		private static readonly HashSet<string> BlockElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
			"p", "div", "br", "hr", "blockquote", "pre", "h1", "h2", "h3", "h4", "h5", "h6",
			"ul", "ol", "li", "dl", "dt", "dd", "table", "thead", "tbody", "tfoot", "tr", "td", "th",
			"section", "article", "nav", "aside", "header", "footer", "main", "figure", "figcaption",
			"address", "details", "summary", "center"
		};

		private static readonly HashSet<string> IgnoredContentElements = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
			"style", "script", "head", "title", "noscript"
		};

		private static readonly char[] TrimmableWhitespace = { ' ', '\t', '\r', '\n', '\f' };

		private static readonly HashSet<string> SupportedTags = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
			"b", "strong", "i", "em", "u", "ins", "s", "del", "strike", "sup", "sub",
			"mark", "a", "code", "kbd", "tt", "samp", "font", "span", "img", "br",
			"p", "div", "h1", "h2", "h3", "h4", "h5", "h6", "pre", "ul", "ol", "li", "center", "hr",
			"html", "body",
			"style", "script", "head", "title", "noscript"
		};

		private static readonly HashSet<string> SupportedCssProperties = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
			"font-weight", "font-style", "text-decoration", "text-decoration-line", "text-decoration-style",
			"color", "background-color", "vertical-align", "text-align", "font-family", "font-size",
			"list-style-type", "list-style"
		};

		private static void ValidateToken(HtmlToken token) {
			if (token.Type == TokenType.Text) return;

			string tagName = token.TagName;
			if (!SupportedTags.Contains(tagName))
				throw new NotSupportedException("Unsupported HTML tag: '" + tagName + "'.");

			if (token.Type == TokenType.CloseTag ||
				IgnoredContentElements.Contains(tagName) ||
				tagName == "html" || tagName == "body")
				return;

			if (token.Attributes != null) {
				foreach (var kvp in token.Attributes) {
					if (!IsAttributeSupported(tagName, kvp.Key))
						throw new NotSupportedException("Unsupported attribute '" + kvp.Key + "' on HTML tag '" + tagName + "'.");
				}

				if (token.Attributes.TryGetValue("style", out string style))
					ValidateCssProperties(tagName, style);
			}
		}

		private static bool IsAttributeSupported(string tagName, string attrName) {
			switch (attrName.ToLowerInvariant()) {
				case "style":
				case "align":
					return true;
			}

			switch (tagName) {
				case "a":
					return attrName.Equals("href", StringComparison.OrdinalIgnoreCase);
				case "img":
					switch (attrName.ToLowerInvariant()) {
						case "src": case "alt": case "width": case "height": return true;
						default: return false;
					}
				case "font":
					switch (attrName.ToLowerInvariant()) {
						case "face": case "size": case "color": return true;
						default: return false;
					}
				case "ol":
					return attrName.Equals("type", StringComparison.OrdinalIgnoreCase);
			}

			return false;
		}

		private static void ValidateCssProperties(string tagName, string style) {
			var declarations = style.Split(';');
			foreach (var decl in declarations) {
				int colon = decl.IndexOf(':');
				if (colon < 0) continue;
				string property = decl.Substring(0, colon).Trim().ToLowerInvariant();
				if (property.Length == 0) continue;
				if (!SupportedCssProperties.Contains(property))
					throw new NotSupportedException("Unsupported CSS property '" + property + "' on HTML tag '" + tagName + "'.");
			}
		}

		private static List<ParagraphBlock> BuildBlocks(List<HtmlToken> tokens, bool strict) {
			var blocks = new List<ParagraphBlock>();
			var styleStack = new List<StyleEntry>();
			var listStack = new List<int>();
			var currentSpans = new List<InlineSpan>();
			int currentListType = tomConstants.tomListNone;
			int currentListLevel = 0;
			int currentHeadingLevel = 0;
			bool currentPreformatted = false;
			int preDepth = 0;
			int ignoreDepth = 0;
			string ignoreTag = null;

			for (int t = 0; t < tokens.Count; t++) {
				var token = tokens[t];

				if (ignoreDepth > 0) {
					if (token.Type == TokenType.OpenTag && token.TagName == ignoreTag)
						ignoreDepth++;
					else if (token.Type == TokenType.CloseTag && token.TagName == ignoreTag)
						ignoreDepth--;
					continue;
				}

				if (strict)
					ValidateToken(token);

				switch (token.Type) {
					case TokenType.Text:
						if (preDepth > 0) {
							var style = ComputeEffectiveStyle(styleStack);
							string[] lines = token.Text.Split('\n');
							for (int li = 0; li < lines.Length; li++) {
								string line = lines[li];
								if (line.Length > 0 && line[line.Length - 1] == '\r')
									line = line.Substring(0, line.Length - 1);
								if (li > 0) {
									FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
										currentHeadingLevel, true, ComputeBlockAlignment(styleStack));
								}
								if (line.Length > 0)
									currentSpans.Add(CreateSpan(line, style));
							}
						}
						else {
							string text = CollapseWhitespace(token.Text);
							if (text.Length > 0) {
								var style = ComputeEffectiveStyle(styleStack);
								currentSpans.Add(CreateSpan(text, style));
							}
						}
						break;

					case TokenType.OpenTag:
						ProcessOpenTag(token, blocks, styleStack, listStack, ref currentSpans,
							ref currentListType, ref currentListLevel, ref currentHeadingLevel,
							ref currentPreformatted, ref preDepth, ref ignoreDepth, ref ignoreTag);
						break;

					case TokenType.SelfClosingTag:
						if (token.TagName == "br" || token.TagName == "hr") {
							FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
								currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
							if (listStack.Count == 0) {
								currentListType = tomConstants.tomListNone;
								currentListLevel = 0;
							}
						}
						else if (token.TagName == "img") {
							byte[] imageData = null;
							string altText = "";
							int imgWidth = 320;
							int imgHeight = 240;

							if (token.Attributes != null) {
								if (token.Attributes.TryGetValue("alt", out string alt))
									altText = alt ?? "";
								if (token.Attributes.TryGetValue("width", out string wStr) &&
									int.TryParse(wStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int w) && w > 0)
									imgWidth = w;
								if (token.Attributes.TryGetValue("height", out string hStr) &&
									int.TryParse(hStr, NumberStyles.Integer, CultureInfo.InvariantCulture, out int h) && h > 0)
									imgHeight = h;
								if (token.Attributes.TryGetValue("src", out string src))
									imageData = ParseDataUri(src);
							}

							if (imageData == null || !ValidateImageData(imageData))
								imageData = CreatePlaceholderImage(imgWidth, imgHeight);

							var imgStyle = ComputeEffectiveStyle(styleStack);
							currentSpans.Add(new InlineSpan {
								ImageData = imageData,
								ImageAltText = altText,
								ImageWidth = imgWidth,
								ImageHeight = imgHeight,
								Url = imgStyle.Url ?? string.Empty
							});
						}
						break;

					case TokenType.CloseTag:
						ProcessCloseTag(token, blocks, styleStack, listStack, ref currentSpans,
							ref currentListType, ref currentListLevel, ref currentHeadingLevel,
							ref currentPreformatted, ref preDepth);
						break;
				}
			}

			FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
				currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
			return blocks;
		}

		private static void ProcessOpenTag(HtmlToken token, List<ParagraphBlock> blocks, List<StyleEntry> styleStack,
			List<int> listStack, ref List<InlineSpan> currentSpans, ref int currentListType, ref int currentListLevel,
			ref int currentHeadingLevel, ref bool currentPreformatted, ref int preDepth,
			ref int ignoreDepth, ref string ignoreTag) {

			if (IgnoredContentElements.Contains(token.TagName)) {
				ignoreDepth = 1;
				ignoreTag = token.TagName;
				return;
			}

			if (token.TagName == "html" || token.TagName == "body") return;

			if (token.TagName == "ul" || token.TagName == "ol") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				int lt = token.TagName == "ul" ? tomConstants.tomListBullet : GetOrderedListType(token.Attributes);
				if (token.Attributes != null && token.Attributes.TryGetValue("style", out string listStyle))
					lt = GetListTypeFromCss(listStyle, lt);
				listStack.Add(lt);
				return;
			}

			if (token.TagName == "li") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count > 0) {
					currentListType = listStack[listStack.Count - 1];
					currentListLevel = listStack.Count - 1;
				}
				PushStyleEntry(styleStack, token);
				return;
			}

			if (token.TagName == "br" || token.TagName == "hr") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			if (token.TagName == "pre") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				preDepth++;
				currentPreformatted = true;
				currentHeadingLevel = 0;
				PushStyleEntry(styleStack, token);
				return;
			}

			int headingLevel = GetHeadingLevel(token.TagName);
			if (headingLevel > 0) {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				currentHeadingLevel = headingLevel;
				currentPreformatted = false;
				PushStyleEntry(styleStack, token);
				return;
			}

			if (BlockElements.Contains(token.TagName)) {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				currentHeadingLevel = 0;
				currentPreformatted = preDepth > 0;
				PushStyleEntry(styleStack, token);
				return;
			}

			PushStyleEntry(styleStack, token);
		}

		private static void ProcessCloseTag(HtmlToken token, List<ParagraphBlock> blocks, List<StyleEntry> styleStack,
			List<int> listStack, ref List<InlineSpan> currentSpans, ref int currentListType, ref int currentListLevel,
			ref int currentHeadingLevel, ref bool currentPreformatted, ref int preDepth) {

			if (token.TagName == "html" || token.TagName == "body") return;

			if (token.TagName == "ul" || token.TagName == "ol") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				if (listStack.Count > 0)
					listStack.RemoveAt(listStack.Count - 1);
				if (listStack.Count > 0) {
					currentListType = listStack[listStack.Count - 1];
					currentListLevel = listStack.Count - 1;
				}
				else {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			if (token.TagName == "li") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				PopStyleEntry(styleStack, token.TagName);
				if (listStack.Count > 0) {
					currentListType = listStack[listStack.Count - 1];
					currentListLevel = listStack.Count - 1;
				}
				else {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			if (token.TagName == "pre") {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				PopStyleEntry(styleStack, token.TagName);
				if (preDepth > 0) preDepth--;
				currentPreformatted = preDepth > 0;
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			if (GetHeadingLevel(token.TagName) > 0) {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				PopStyleEntry(styleStack, token.TagName);
				currentHeadingLevel = 0;
				currentPreformatted = preDepth > 0;
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			if (BlockElements.Contains(token.TagName)) {
				FlushBlock(blocks, ref currentSpans, currentListType, currentListLevel,
					currentHeadingLevel, currentPreformatted, ComputeBlockAlignment(styleStack));
				PopStyleEntry(styleStack, token.TagName);
				if (listStack.Count == 0) {
					currentListType = tomConstants.tomListNone;
					currentListLevel = 0;
				}
				return;
			}

			PopStyleEntry(styleStack, token.TagName);
		}

		private static void FlushBlock(List<ParagraphBlock> blocks, ref List<InlineSpan> currentSpans,
			int listType, int listLevel, int headingLevel, bool preformatted, int alignment) {
			if (currentSpans.Count == 0) return;

			if (!preformatted)
				TrimBlockWhitespace(currentSpans);
			if (currentSpans.Count == 0) return;

			blocks.Add(new ParagraphBlock {
				Spans = currentSpans,
				ListType = listType,
				ListLevel = listLevel,
				HeadingLevel = headingLevel,
				Preformatted = preformatted,
				Alignment = alignment
			});
			currentSpans = new List<InlineSpan>();
		}

		private static int ComputeBlockAlignment(List<StyleEntry> styleStack) {
			int alignment = tomConstants.tomAlignLeft;
			for (int i = 0; i < styleStack.Count; i++) {
				if (styleStack[i].Override.Alignment.HasValue)
					alignment = styleStack[i].Override.Alignment.Value;
			}
			return alignment;
		}

		private static int GetHeadingLevel(string tagName) {
			if (tagName.Length == 2 && tagName[0] == 'h' && tagName[1] >= '1' && tagName[1] <= '6')
				return tagName[1] - '0';
			return 0;
		}

		private static void TrimBlockWhitespace(List<InlineSpan> spans) {
			while (spans.Count > 0 && spans[0].Text != null && spans[0].Text.Trim(TrimmableWhitespace).Length == 0)
				spans.RemoveAt(0);
			while (spans.Count > 0 && spans[spans.Count - 1].Text != null && spans[spans.Count - 1].Text.Trim(TrimmableWhitespace).Length == 0)
				spans.RemoveAt(spans.Count - 1);
			if (spans.Count > 0 && spans[0].Text != null) {
				var first = spans[0];
				string trimmedStart = first.Text.TrimStart(TrimmableWhitespace);
				if (trimmedStart != first.Text) {
					first.Text = trimmedStart;
					spans[0] = first;
				}
			}
			if (spans.Count > 0 && spans[spans.Count - 1].Text != null) {
				var last = spans[spans.Count - 1];
				string trimmedEnd = last.Text.TrimEnd(TrimmableWhitespace);
				if (trimmedEnd != last.Text) {
					last.Text = trimmedEnd;
					spans[spans.Count - 1] = last;
				}
			}
			while (spans.Count > 0 && spans[0].Text != null && spans[0].Text.Length == 0)
				spans.RemoveAt(0);
			while (spans.Count > 0 && spans[spans.Count - 1].Text != null && spans[spans.Count - 1].Text.Length == 0)
				spans.RemoveAt(spans.Count - 1);
		}

		private static string CollapseWhitespace(string text) {
			if (text.Length == 0) return text;
			var sb = new StringBuilder(text.Length);
			bool lastWasSpace = false;
			for (int i = 0; i < text.Length; i++) {
				char c = text[i];
				if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\f') {
					if (!lastWasSpace) {
						sb.Append(' ');
						lastWasSpace = true;
					}
				}
				else {
					sb.Append(c);
					lastWasSpace = false;
				}
			}
			return sb.ToString();
		}

		#endregion

		#region - Style Processing -

		private static void PushStyleEntry(List<StyleEntry> stack, HtmlToken token) {
			var over = GetStyleOverrideForTag(token.TagName, token.Attributes);
			stack.Add(new StyleEntry { TagName = token.TagName, Override = over });
		}

		private static void PopStyleEntry(List<StyleEntry> stack, string tagName) {
			for (int i = stack.Count - 1; i >= 0; i--) {
				if (stack[i].TagName == tagName) {
					stack.RemoveAt(i);
					return;
				}
			}
		}

		private static StyleOverride GetStyleOverrideForTag(string tagName, Dictionary<string, string> attrs) {
			var over = new StyleOverride();

			switch (tagName) {
				case "b":
				case "strong":
					over.Bold = true;
					break;
				case "i":
				case "em":
					over.Italic = true;
					break;
				case "u":
				case "ins":
					over.Underline = (int)UnderlineStyle.Single;
					break;
				case "s":
				case "del":
				case "strike":
					over.StrikeThrough = true;
					break;
				case "sup":
					over.Superscript = true;
					break;
				case "sub":
					over.Subscript = true;
					break;
				case "mark":
					over.BackColor = ColorTranslator.ToOle(Color.Yellow);
					break;
				case "a":
					if (attrs != null && attrs.TryGetValue("href", out string href) && !string.IsNullOrEmpty(href))
						over.Url = href;
					break;
				case "center":
					over.Alignment = tomConstants.tomAlignCenter;
					break;
				case "code":
				case "kbd":
				case "tt":
				case "samp":
					over.FontName = BlockStyleHelper.PreformattedFontName;
					break;
				case "font":
					if (attrs != null) {
						if (attrs.TryGetValue("face", out string face) && !string.IsNullOrEmpty(face))
							over.FontName = ParseFontFamilyValue(face);
						if (attrs.TryGetValue("size", out string sizeAttr)) {
							float? fs = ParseHtmlFontSize(sizeAttr);
							if (fs.HasValue) over.FontSize = fs.Value;
						}
						if (attrs.TryGetValue("color", out string fontColor)) {
							int? fc = ParseCssColor(fontColor);
							if (fc.HasValue) over.ForeColor = fc.Value;
						}
					}
					break;
			}

			if (attrs != null) {
				if (attrs.TryGetValue("align", out string alignAttr)) {
					int? al = MapAlignValue(alignAttr.Trim().ToLowerInvariant());
					if (al.HasValue) over.Alignment = al.Value;
				}
				if (attrs.TryGetValue("style", out string style))
					MergeInlineStyle(ref over, style);
			}

			return over;
		}

		private static void MergeInlineStyle(ref StyleOverride over, string style) {
			bool? hasUnderline = null;
			int underlineStyleVal = (int)UnderlineStyle.Single;

			var declarations = style.Split(';');
			foreach (var decl in declarations) {
				int colon = decl.IndexOf(':');
				if (colon < 0) continue;
				string property = decl.Substring(0, colon).Trim().ToLowerInvariant();
				string value = decl.Substring(colon + 1).Trim();
				string valueLower = value.ToLowerInvariant();

				switch (property) {
					case "font-weight":
						if (valueLower == "bold" || valueLower == "bolder" ||
							(int.TryParse(valueLower, NumberStyles.Integer, CultureInfo.InvariantCulture, out int weight) && weight >= 700))
							over.Bold = true;
						else if (valueLower == "normal" || valueLower == "lighter")
							over.Bold = false;
						break;

					case "font-style":
						if (valueLower == "italic" || valueLower == "oblique")
							over.Italic = true;
						else if (valueLower == "normal")
							over.Italic = false;
						break;

					case "text-decoration":
						if (valueLower == "none") {
							hasUnderline = false;
							over.StrikeThrough = false;
						}
						else {
							if (valueLower.Contains("underline"))
								hasUnderline = true;
							if (valueLower.Contains("line-through"))
								over.StrikeThrough = true;
							if (valueLower.Contains("double")) underlineStyleVal = (int)UnderlineStyle.Double;
							else if (valueLower.Contains("dotted")) underlineStyleVal = (int)UnderlineStyle.Dotted;
							else if (valueLower.Contains("dashed")) underlineStyleVal = (int)UnderlineStyle.Dashed;
							else if (valueLower.Contains("wavy")) underlineStyleVal = (int)UnderlineStyle.Wavy;
						}
						break;

					case "text-decoration-line":
						if (valueLower == "none") {
							hasUnderline = false;
							over.StrikeThrough = false;
						}
						else {
							if (valueLower.Contains("underline"))
								hasUnderline = true;
							if (valueLower.Contains("line-through"))
								over.StrikeThrough = true;
						}
						break;

					case "text-decoration-style":
						underlineStyleVal = MapCssUnderlineStyle(valueLower);
						break;

					case "color":
						int? fg = ParseCssColor(value);
						if (fg.HasValue) over.ForeColor = fg.Value;
						break;

					case "background-color":
						int? bg = ParseCssColor(value);
						if (bg.HasValue) over.BackColor = bg.Value;
						break;

					case "vertical-align":
							if (valueLower == "super")
								over.Superscript = true;
							else if (valueLower == "sub")
								over.Subscript = true;
							break;

						case "text-align":
							int? align = MapAlignValue(valueLower);
							if (align.HasValue) over.Alignment = align.Value;
							break;

						case "font-family":
							string fn = ParseFontFamilyValue(value);
							if (!string.IsNullOrEmpty(fn)) over.FontName = fn;
							break;

						case "font-size":
							float? sz = ParseCssFontSize(valueLower);
							if (sz.HasValue) over.FontSize = sz.Value;
							break;
					}
			}

			if (hasUnderline == true)
				over.Underline = underlineStyleVal;
			else if (hasUnderline == false)
				over.Underline = 0;
		}

		private static EffectiveStyle ComputeEffectiveStyle(List<StyleEntry> stack) {
			var style = new EffectiveStyle {
				ForeColor = tomConstants.tomAutoColor,
				BackColor = tomConstants.tomAutoColor,
				Url = string.Empty,
				FontName = string.Empty
			};

			for (int i = 0; i < stack.Count; i++) {
				var over = stack[i].Override;
				if (over.Bold.HasValue) style.Bold = over.Bold.Value;
				if (over.Italic.HasValue) style.Italic = over.Italic.Value;
				if (over.Underline.HasValue) style.Underline = over.Underline.Value;
				if (over.StrikeThrough.HasValue) style.StrikeThrough = over.StrikeThrough.Value;
				if (over.Superscript.HasValue) style.Superscript = over.Superscript.Value;
				if (over.Subscript.HasValue) style.Subscript = over.Subscript.Value;
				if (over.ForeColor.HasValue) style.ForeColor = over.ForeColor.Value;
				if (over.BackColor.HasValue) style.BackColor = over.BackColor.Value;
				if (over.Url != null) style.Url = over.Url;
				if (over.FontName != null) style.FontName = over.FontName;
				if (over.FontSize.HasValue) style.FontSize = over.FontSize.Value;
			}

			return style;
		}

		private static InlineSpan CreateSpan(string text, EffectiveStyle style) {
			return new InlineSpan {
				Text = text,
				Bold = style.Bold,
				Italic = style.Italic,
				Underline = style.Underline,
				StrikeThrough = style.StrikeThrough,
				Superscript = style.Superscript,
				Subscript = style.Subscript,
				ForeColor = style.ForeColor,
				BackColor = style.BackColor,
				Url = style.Url ?? string.Empty,
				FontName = style.FontName ?? string.Empty,
				FontSize = style.FontSize
			};
		}

		#endregion

		#region - CSS Parsing -

		private static int MapCssUnderlineStyle(string value) {
			switch (value) {
				case "double": return (int)UnderlineStyle.Double;
				case "dotted": return (int)UnderlineStyle.Dotted;
				case "dashed": return (int)UnderlineStyle.Dashed;
				case "wavy": return (int)UnderlineStyle.Wavy;
				default: return (int)UnderlineStyle.Single;
			}
		}

		private static int GetOrderedListType(Dictionary<string, string> attrs) {
			if (attrs != null && attrs.TryGetValue("type", out string type)) {
				switch (type) {
					case "a": return tomConstants.tomListNumberAsLCLetter;
					case "A": return tomConstants.tomListNumberAsUCLetter;
					case "i": return tomConstants.tomListNumberAsLCRoman;
					case "I": return tomConstants.tomListNumberAsUCRoman;
				}
			}
			return tomConstants.tomListNumberAsArabic;
		}

		private static int GetListTypeFromCss(string style, int defaultType) {
			var declarations = style.Split(';');
			foreach (var decl in declarations) {
				int colon = decl.IndexOf(':');
				if (colon < 0) continue;
				string property = decl.Substring(0, colon).Trim().ToLowerInvariant();
				if (property != "list-style-type" && property != "list-style") continue;
				string value = decl.Substring(colon + 1).Trim().ToLowerInvariant();
				switch (value) {
					case "disc":
					case "circle":
					case "square":
						return tomConstants.tomListBullet;
					case "decimal":
						return tomConstants.tomListNumberAsArabic;
					case "lower-alpha":
					case "lower-latin":
						return tomConstants.tomListNumberAsLCLetter;
					case "upper-alpha":
					case "upper-latin":
						return tomConstants.tomListNumberAsUCLetter;
					case "lower-roman":
						return tomConstants.tomListNumberAsLCRoman;
					case "upper-roman":
						return tomConstants.tomListNumberAsUCRoman;
				}
			}
			return defaultType;
		}

		private static int? ParseCssColor(string value) {
			value = value.Trim().ToLowerInvariant();
			if (value.Length == 0 || value == "inherit" || value == "initial" || value == "unset" ||
				value == "currentcolor" || value == "transparent")
				return null;

			if (value.StartsWith("rgba(") && value.EndsWith(")")) {
				string inner = value.Substring(5, value.Length - 6);
				var parts = inner.Split(',');
				if (parts.Length >= 3 &&
					int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int r) &&
					int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int g) &&
					int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int b))
					return ColorTranslator.ToOle(Color.FromArgb(Clamp(r), Clamp(g), Clamp(b)));
				return null;
			}

			if (value.StartsWith("rgb(") && value.EndsWith(")")) {
				string inner = value.Substring(4, value.Length - 5);
				var parts = inner.Split(',');
				if (parts.Length == 3 &&
					int.TryParse(parts[0].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int r) &&
					int.TryParse(parts[1].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int g) &&
					int.TryParse(parts[2].Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out int b))
					return ColorTranslator.ToOle(Color.FromArgb(Clamp(r), Clamp(g), Clamp(b)));
				return null;
			}

			if (value.Length == 4 && value[0] == '#') {
				int r = HexCharToInt(value[1]);
				int g = HexCharToInt(value[2]);
				int b = HexCharToInt(value[3]);
				if (r >= 0 && g >= 0 && b >= 0)
					return ColorTranslator.ToOle(Color.FromArgb(r * 17, g * 17, b * 17));
			}

			try {
				var c = ColorTranslator.FromHtml(value);
				return ColorTranslator.ToOle(c);
			}
			catch {
				return null;
			}
		}

		private static int HexCharToInt(char c) {
			if (c >= '0' && c <= '9') return c - '0';
			if (c >= 'a' && c <= 'f') return c - 'a' + 10;
			if (c >= 'A' && c <= 'F') return c - 'A' + 10;
			return -1;
		}

		private static int Clamp(int v) => v < 0 ? 0 : v > 255 ? 255 : v;

		private static int? MapAlignValue(string value) {
			switch (value) {
				case "left": return tomConstants.tomAlignLeft;
				case "center": return tomConstants.tomAlignCenter;
				case "right": return tomConstants.tomAlignRight;
				case "justify": return tomConstants.tomAlignJustify;
				default: return null;
			}
		}

		private static string ParseFontFamilyValue(string value) {
			var families = value.Split(',');
			foreach (var family in families) {
				string name = family.Trim().Trim('"', '\'').Trim();
				if (name.Length == 0) continue;
				string lower = name.ToLowerInvariant();
				if (lower == "serif" || lower == "sans-serif" || lower == "monospace" ||
					lower == "cursive" || lower == "fantasy" || lower == "system-ui" ||
					lower == "ui-serif" || lower == "ui-sans-serif" || lower == "ui-monospace" ||
					lower == "ui-rounded" || lower == "emoji" || lower == "math" || lower == "fangsong")
					continue;
				return name;
			}
			return null;
		}

		private static float? ParseCssFontSize(string value) {
			value = value.Trim();
			if (value.Length == 0) return null;

			switch (value) {
				case "xx-small": return 7f;
				case "x-small": return 8f;
				case "small": return 10f;
				case "medium": return 12f;
				case "large": return 14f;
				case "x-large": return 18f;
				case "xx-large": return 24f;
				case "smaller": return null;
				case "larger": return null;
			}

			if (value.EndsWith("pt")) {
				if (float.TryParse(value.Substring(0, value.Length - 2).Trim(),
					NumberStyles.Float, CultureInfo.InvariantCulture, out float pt) && pt > 0)
					return pt;
			}
			else if (value.EndsWith("px")) {
				if (float.TryParse(value.Substring(0, value.Length - 2).Trim(),
					NumberStyles.Float, CultureInfo.InvariantCulture, out float px) && px > 0)
					return px * 0.75f;
			}
			else if (value.EndsWith("em")) {
				if (float.TryParse(value.Substring(0, value.Length - 2).Trim(),
					NumberStyles.Float, CultureInfo.InvariantCulture, out float em) && em > 0)
					return em * BlockStyleHelper.ParagraphPointSize;
			}
			else if (value.EndsWith("rem")) {
				if (float.TryParse(value.Substring(0, value.Length - 3).Trim(),
					NumberStyles.Float, CultureInfo.InvariantCulture, out float rem) && rem > 0)
					return rem * BlockStyleHelper.ParagraphPointSize;
			}
			else if (value.EndsWith("%")) {
				if (float.TryParse(value.Substring(0, value.Length - 1).Trim(),
					NumberStyles.Float, CultureInfo.InvariantCulture, out float pct) && pct > 0)
					return pct / 100f * BlockStyleHelper.ParagraphPointSize;
			}

			return null;
		}

		private static float? ParseHtmlFontSize(string value) {
			value = value.Trim();
			if (int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int size)) {
				switch (size) {
					case 1: return 8f;
					case 2: return 10f;
					case 3: return 12f;
					case 4: return 14f;
					case 5: return 18f;
					case 6: return 24f;
					case 7: return 36f;
				}
			}
			return ParseCssFontSize(value);
		}

		#endregion

		#region - Image Processing -

		private static byte[] ParseDataUri(string src) {
			if (string.IsNullOrEmpty(src)) return null;
			if (!src.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) return null;

			int commaIndex = src.IndexOf(',');
			if (commaIndex < 0) return null;

			string meta = src.Substring(5, commaIndex - 5);
			if (meta.IndexOf(";base64", StringComparison.OrdinalIgnoreCase) < 0) return null;

			string base64 = src.Substring(commaIndex + 1);
			base64 = base64.Replace('-', '+').Replace('_', '/');

			int mod = base64.Length % 4;
			if (mod == 2) base64 += "==";
			else if (mod == 3) base64 += "=";

			try {
				return Convert.FromBase64String(base64);
			}
			catch {
				return null;
			}
		}

		private static bool ValidateImageData(byte[] data) {
			try {
				using (var ms = new MemoryStream(data))
				using (var image = Image.FromStream(ms)) {
					var guid = image.RawFormat.Guid;
					return guid == ImageFormat.Jpeg.Guid ||
						guid == ImageFormat.Png.Guid ||
						guid == ImageFormat.Gif.Guid;
				}
			}
			catch {
				return false;
			}
		}

		private static byte[] CreatePlaceholderImage(int width, int height) {
			using (var bmp = new Bitmap(width, height))
			using (var g = Graphics.FromImage(bmp)) {
				g.Clear(Color.LightGray);
				using (var pen = new Pen(Color.Red, 2f)) {
					g.DrawRectangle(pen, 1, 1, width - 3, height - 3);
					g.DrawLine(pen, 0, 0, width, height);
					g.DrawLine(pen, width, 0, 0, height);
				}
				using (var ms = new MemoryStream()) {
					bmp.Save(ms, ImageFormat.Png);
					return ms.ToArray();
				}
			}
		}

		private static int PixelsToHimetric(int pixels) {
			return (int)(pixels * 2540L / 96);
		}

		#endregion

		#region - TOM2 Insertion -

		private static void InsertBlocks(ITextDocument2 doc, IRichEditOle richOle, List<ParagraphBlock> blocks) {
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
								if (prevBlock.Alignment != tomConstants.tomAlignLeft)
									prevPara.Alignment = prevBlock.Alignment;
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
					InsertSpans(doc, richOle, range, block.Spans);
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

						var rFont = range.Font;
						try {
							rFont.Bold = tomConstants.tomFalse;
							rFont.Size = BlockStyleHelper.ParagraphPointSize;
						}
						finally {
							Marshal.ReleaseComObject(rFont);
						}
					}
					else if (block.Preformatted && blockEnd > blockStart) {
						var preRange = doc.Range2(blockStart, blockEnd);
						try {
							var pFont = preRange.Font;
							try {
								pFont.Name = BlockStyleHelper.PreformattedFontName;
								pFont.Size = BlockStyleHelper.PreformattedPointSize;
							}
							finally {
								Marshal.ReleaseComObject(pFont);
							}
						}
						finally {
							Marshal.ReleaseComObject(preRange);
						}

						var rFont = range.Font;
						try {
							rFont.Name = string.Empty;
							rFont.Size = BlockStyleHelper.ParagraphPointSize;
						}
						finally {
							Marshal.ReleaseComObject(rFont);
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
						if (lastBlock.Alignment != tomConstants.tomAlignLeft)
							lastPara.Alignment = lastBlock.Alignment;
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

		private static void InsertSpans(ITextDocument2 doc, IRichEditOle richOle, ITextRange2 range, List<InlineSpan> spans) {
			for (int s = 0; s < spans.Count; s++) {
				var span = spans[s];

				if (span.ImageData != null) {
					int imgStart = range.Start;
					if (richOle != null) {
						int hr = ImageOleObject.Insert(richOle, imgStart,
							span.ImageData, span.ImageAltText ?? string.Empty,
							PixelsToHimetric(span.ImageWidth), PixelsToHimetric(span.ImageHeight));
						if (hr >= 0) {
							range.SetRange(imgStart + 1, imgStart + 1);

							if (!string.IsNullOrEmpty(span.Url)) {
								var linkRange = doc.Range2(imgStart, imgStart + 1);
								try {
									linkRange.URL = "\"" + span.Url + "\"";
								}
								finally {
									Marshal.ReleaseComObject(linkRange);
								}
							}
						}
					}
					continue;
				}

				if (string.IsNullOrEmpty(span.Text)) continue;

				int insertStart = range.Start;
				range.Text = span.Text;
				int insertEnd = insertStart + span.Text.Length;
				range.SetRange(insertEnd, insertEnd);

				bool needsFormatting = span.Bold || span.Italic || span.Underline != 0 ||
					span.StrikeThrough || span.Superscript || span.Subscript ||
					span.ForeColor != tomConstants.tomAutoColor || span.BackColor != tomConstants.tomAutoColor ||
					!string.IsNullOrEmpty(span.Url) ||
					!string.IsNullOrEmpty(span.FontName) || span.FontSize > 0;

				if (needsFormatting) {
					var fmtRange = doc.Range2(insertStart, insertEnd);
					try {
						bool needsFontFormat = span.Bold || span.Italic || span.Underline != 0 ||
							span.StrikeThrough || span.Superscript || span.Subscript ||
							span.ForeColor != tomConstants.tomAutoColor || span.BackColor != tomConstants.tomAutoColor ||
							!string.IsNullOrEmpty(span.FontName) || span.FontSize > 0;

						if (needsFontFormat) {
							var font = fmtRange.Font;
							try {
								if (span.Bold) font.Bold = tomConstants.tomTrue;
								if (span.Italic) font.Italic = tomConstants.tomTrue;
								if (span.Underline != 0) font.Underline = span.Underline;
								if (span.StrikeThrough) font.StrikeThrough = tomConstants.tomTrue;
								if (span.Superscript) font.Superscript = tomConstants.tomTrue;
								if (span.Subscript) font.Subscript = tomConstants.tomTrue;
								if (span.ForeColor != tomConstants.tomAutoColor) font.ForeColor = span.ForeColor;
								if (span.BackColor != tomConstants.tomAutoColor) font.BackColor = span.BackColor;
								if (!string.IsNullOrEmpty(span.FontName)) font.Name = span.FontName;
								if (span.FontSize > 0) font.Size = span.FontSize;
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
