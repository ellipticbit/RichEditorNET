using System;
using System.Collections.Generic;
using System.Drawing;

namespace EllipticBit.RichEditorNET
{
	internal sealed class ToolbarIconCache : IDisposable
	{
		private readonly List<Bitmap> _icons = new List<Bitmap>();

		internal float DpiScale { get; }

		internal Bitmap Bold { get; }
		internal Bitmap Italic { get; }
		internal Bitmap Underline { get; }
		internal Bitmap UnderlineSingle { get; }
		internal Bitmap UnderlineDouble { get; }
		internal Bitmap UnderlineDotted { get; }
		internal Bitmap UnderlineDashed { get; }
		internal Bitmap UnderlineWavy { get; }
		internal Bitmap Strikethrough { get; }
		internal Bitmap Subscript { get; }
		internal Bitmap Superscript { get; }
		internal Bitmap FontColor { get; }
		internal Bitmap BackgroundColor { get; }
		internal Bitmap AlignLeft { get; }
		internal Bitmap AlignCenter { get; }
		internal Bitmap AlignRight { get; }
		internal Bitmap AlignJustify { get; }
		internal Bitmap FontSizeIncrease { get; }
		internal Bitmap FontSizeDecrease { get; }
		internal Bitmap Hyperlink { get; }
		internal Bitmap InsertImage { get; }
		internal Bitmap InsertLinkedThumbnail { get; }
		internal Bitmap BulletList { get; }
		internal Bitmap OrderedList { get; }

		internal ToolbarIconCache(float dpiScale)
		{
			DpiScale = dpiScale;

			Bold = Cache(ToolbarIcons.Bold(dpiScale));
			Italic = Cache(ToolbarIcons.Italic(dpiScale));
			Underline = Cache(ToolbarIcons.Underline(dpiScale));
			UnderlineSingle = Cache(ToolbarIcons.UnderlineSingle(dpiScale));
			UnderlineDouble = Cache(ToolbarIcons.UnderlineDouble(dpiScale));
			UnderlineDotted = Cache(ToolbarIcons.UnderlineDotted(dpiScale));
			UnderlineDashed = Cache(ToolbarIcons.UnderlineDashed(dpiScale));
			UnderlineWavy = Cache(ToolbarIcons.UnderlineWavy(dpiScale));
			Strikethrough = Cache(ToolbarIcons.Strikethrough(dpiScale));
			Subscript = Cache(ToolbarIcons.Subscript(dpiScale));
			Superscript = Cache(ToolbarIcons.Superscript(dpiScale));
			FontColor = Cache(ToolbarIcons.FontColor(dpiScale));
			BackgroundColor = Cache(ToolbarIcons.BackgroundColor(dpiScale));
			AlignLeft = Cache(ToolbarIcons.AlignLeft(dpiScale));
			AlignCenter = Cache(ToolbarIcons.AlignCenter(dpiScale));
			AlignRight = Cache(ToolbarIcons.AlignRight(dpiScale));
			AlignJustify = Cache(ToolbarIcons.AlignJustify(dpiScale));
			FontSizeIncrease = Cache(ToolbarIcons.FontSizeIncrease(dpiScale));
			FontSizeDecrease = Cache(ToolbarIcons.FontSizeDecrease(dpiScale));
			Hyperlink = Cache(ToolbarIcons.Hyperlink(dpiScale));
			InsertImage = Cache(ToolbarIcons.InsertImage(dpiScale));
			InsertLinkedThumbnail = Cache(ToolbarIcons.InsertLinkedThumbnail(dpiScale));
			BulletList = Cache(ToolbarIcons.BulletList(dpiScale));
			OrderedList = Cache(ToolbarIcons.OrderedList(dpiScale));
		}

		private Bitmap Cache(Bitmap bmp)
		{
			_icons.Add(bmp);
			return bmp;
		}

		public void Dispose()
		{
			foreach (var icon in _icons)
				icon.Dispose();
			_icons.Clear();
		}
	}
}
