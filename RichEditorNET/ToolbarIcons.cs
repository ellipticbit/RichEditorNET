using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace EllipticBit.RichEditorNET
{
	internal static class ToolbarIcons
	{
		private const int Size = 16;
		private const int IconPadding = 2;
		internal const int TotalIconSize = Size + IconPadding * 2;
		private const int RenderSize = 64;

		private static Bitmap Create(float scale, Action<Graphics> draw)
		{
			int targetSize = Math.Max((int)(TotalIconSize * scale), 1);
			int renderSize = Math.Max(RenderSize, targetSize);
			float renderScale = (float)renderSize / TotalIconSize;

			var hires = new Bitmap(renderSize, renderSize);
			using (var g = Graphics.FromImage(hires))
			{
				g.SmoothingMode = SmoothingMode.AntiAlias;
				g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
				g.ScaleTransform(renderScale, renderScale);
				g.TranslateTransform(IconPadding, IconPadding);
				draw(g);
			}

			if (targetSize == renderSize)
				return hires;

			try
			{
				var result = new Bitmap(targetSize, targetSize);
				using (var g = Graphics.FromImage(result))
				{
					g.InterpolationMode = InterpolationMode.HighQualityBicubic;
					g.SmoothingMode = SmoothingMode.HighQuality;
					g.CompositingQuality = CompositingQuality.HighQuality;
					g.PixelOffsetMode = PixelOffsetMode.HighQuality;
					g.DrawImage(hires, 0, 0, targetSize, targetSize);
				}
				return result;
			}
			finally
			{
				hires.Dispose();
			}
		}

		private static void DrawCenteredText(Graphics g, string text, Font font, Brush brush)
		{
			using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
			{
				g.DrawString(text, font, brush, new RectangleF(0, 1, Size, Size), sf);
			}
		}

		internal static Bitmap Bold(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 6f, FontStyle.Bold))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "B", font, brush);
			});
		}

		internal static Bitmap Italic(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 6f, FontStyle.Italic))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "I", font, brush);
			});
		}

		internal static Bitmap Underline(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
					g.DrawLine(pen, 3, 16, 13, 16);
			});
		}

		internal static Bitmap UnderlineSingle(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f))
					g.DrawLine(pen, 2, 16, 14, 16);
			});
		}

		internal static Bitmap UnderlineDouble(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f))
				{
					g.DrawLine(pen, 2, 16, 14, 16);
					g.DrawLine(pen, 2, 14, 14, 14);
				}
			});
		}

		internal static Bitmap UnderlineDotted(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f) { DashStyle = DashStyle.Dot })
					g.DrawLine(pen, 2, 16, 14, 16);
			});
		}

		internal static Bitmap UnderlineDashed(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f) { DashStyle = DashStyle.Dash })
					g.DrawLine(pen, 2, 16, 14, 16);
			});
		}

		internal static Bitmap UnderlineWavy(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, Size, Size), sf);
				var color = Color.FromArgb(60, 60, 60);
				using (var pen = new Pen(color, 1f))
				{
					var points = new[] {
						new PointF(2, 16), new PointF(4, 14), new PointF(6, 16), new PointF(8, 14),
						new PointF(10, 16), new PointF(12, 14), new PointF(14, 16)
					};
					g.DrawCurve(pen, points, 0.5f);
				}
			});
		}

		internal static Bitmap Strikethrough(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 4f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "ab", font, brush);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
					g.DrawLine(pen, 2, 10, 14, 10);
			});
		}

		internal static Bitmap Subscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 3f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("a", font, brush, 0, -4);
					g.DrawString("2", small, brush, 9, 8);
				}
			});
		}

		internal static Bitmap Superscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 3f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("a", font, brush, 0, -4);
					g.DrawString("2", small, brush, 9, -3);
				}
			});
		}

		internal static Bitmap FontColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 5.5f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, 1, -4);
				using (var brush = new SolidBrush(Color.Red))
					g.FillRectangle(brush, 1, 13, 14, 3);
			});
		}

		internal static Bitmap BackgroundColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var bgBrush = new SolidBrush(Color.Yellow))
					g.FillRectangle(bgBrush, 1, 1, 14, 14);
				using (var font = new Font("Segoe UI", 4f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
						g.DrawString("ab", font, brush, new RectangleF(0, 0, Size, Size), sf);
				}
				using (var pen = new Pen(Color.FromArgb(120, 120, 120), 1f))
					g.DrawRectangle(pen, 0, 0, 15, 15);
			});
		}

		internal static Bitmap AlignLeft(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
				{
					g.DrawLine(pen, 2, 3, 14, 3);
					g.DrawLine(pen, 2, 6, 10, 6);
					g.DrawLine(pen, 2, 9, 13, 9);
					g.DrawLine(pen, 2, 12, 9, 12);
				}
			});
		}

		internal static Bitmap AlignCenter(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
				{
					g.DrawLine(pen, 2, 3, 14, 3);
					g.DrawLine(pen, 4, 6, 12, 6);
					g.DrawLine(pen, 3, 9, 13, 9);
					g.DrawLine(pen, 5, 12, 11, 12);
				}
			});
		}

		internal static Bitmap AlignRight(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
				{
					g.DrawLine(pen, 2, 3, 14, 3);
					g.DrawLine(pen, 6, 6, 14, 6);
					g.DrawLine(pen, 3, 9, 14, 9);
					g.DrawLine(pen, 7, 12, 14, 12);
				}
			});
		}

		internal static Bitmap AlignJustify(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
				{
					g.DrawLine(pen, 2, 3, 14, 3);
					g.DrawLine(pen, 2, 6, 14, 6);
					g.DrawLine(pen, 2, 9, 14, 9);
					g.DrawLine(pen, 2, 12, 14, 12);
				}
			});
		}

		internal static Bitmap FontSizeIncrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 6f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -2, -2);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f))
				{
					g.DrawLine(pen, 12, 2, 12, 8);
					g.DrawLine(pen, 9, 5, 15, 5);
				}
			});
		}

		internal static Bitmap FontSizeDecrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 5f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -2, 0);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1.5f))
					g.DrawLine(pen, 10, 5, 15, 5);
			});
		}

		internal static Bitmap Hyperlink(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 1.5f))
				{
					g.DrawArc(pen, 1, 4, 8, 8, 90, 180);
					g.DrawLine(pen, 5, 4, 8, 4);
					g.DrawLine(pen, 5, 12, 8, 12);

					g.DrawArc(pen, 7, 4, 8, 8, -90, 180);
					g.DrawLine(pen, 8, 4, 11, 4);
					g.DrawLine(pen, 8, 12, 11, 12);
				}
			});
		}

		internal static Bitmap InsertImage(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f))
					g.DrawRectangle(pen, 1, 2, 13, 11);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 2, 3, 12, 10);

				using (var brush = new SolidBrush(Color.FromArgb(255, 200, 0)))
					g.FillEllipse(brush, 3, 4, 4, 4);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(2, 13), new Point(6, 8), new Point(9, 11), new Point(11, 9), new Point(14, 13) };
					g.FillPolygon(brush, points);
				}
			});
		}

		internal static Bitmap InsertLinkedThumbnail(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 1f))
					g.DrawRectangle(pen, 1, 2, 10, 9);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 2, 3, 9, 8);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(2, 11), new Point(5, 7), new Point(7, 9), new Point(9, 7), new Point(11, 11) };
					g.FillPolygon(brush, points);
				}

				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 1.5f))
				{
					g.DrawArc(pen, 10, 9, 4, 5, 90, 180);
					g.DrawLine(pen, 12, 9, 13, 9);
					g.DrawLine(pen, 12, 14, 13, 14);
					g.DrawArc(pen, 12, 9, 4, 5, -90, 180);
				}
			});
		}

		internal static Bitmap BulletList(float scale)
		{
			return Create(scale, g =>
			{
				var bulletColor = Color.FromArgb(60, 60, 60);
				var lineColor = Color.FromArgb(100, 100, 100);
				using (var bulletBrush = new SolidBrush(bulletColor))
				using (var linePen = new Pen(lineColor, 1.5f)) {
					g.Clear(Color.Transparent);
					g.SmoothingMode = SmoothingMode.AntiAlias;
					g.FillEllipse(Brushes.Black, 1, 2, 3, 3);
					g.DrawLine(Pens.Black, 6, 3, 14, 3);
					g.FillEllipse(Brushes.Black, 1, 7, 3, 3);
					g.DrawLine(Pens.Black, 6, 8, 14, 8);
					g.FillEllipse(Brushes.Black, 1, 12, 3, 3);
					g.DrawLine(Pens.Black, 6, 13, 14, 13);

					//g.FillEllipse(bulletBrush, 2, 2, 3, 3);
					//g.DrawLine(linePen, 7, 3, 14, 3);

					//g.FillEllipse(bulletBrush, 2, 6, 3, 3);
					//g.DrawLine(linePen, 7, 7, 14, 7);

					//g.FillEllipse(bulletBrush, 2, 10, 3, 3);
					//g.DrawLine(linePen, 7, 11, 14, 11);
				}
			});
		}

		internal static Bitmap OrderedList(float scale)
		{
			return Create(scale, g =>
			{
				var textColor = Color.FromArgb(60, 60, 60);
				var lineColor = Color.FromArgb(100, 100, 100);
				using (var font = new Font("Segoe UI", 2f, FontStyle.Bold))
				using (var brush = new SolidBrush(textColor))
				using (var linePen = new Pen(lineColor, 2f))
				{
					g.DrawString("1", font, brush, 0, 0);
					g.DrawLine(linePen, 6, 3, 14, 3);

					g.DrawString("2", font, brush, 0, 5);
					g.DrawLine(linePen, 6, 8, 14, 8);

					g.DrawString("3", font, brush, 0, 10);
					g.DrawLine(linePen, 6, 13, 14, 13);
				}
			});
		}
	}
}
