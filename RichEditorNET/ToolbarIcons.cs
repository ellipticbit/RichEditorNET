using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace EllipticBit.RichEditorNET
{
	internal static class ToolbarIcons
	{
		private const int DrawArea = 48;
		private const int DrawPadding = 8;
		private const int DrawSize = DrawArea + DrawPadding * 2;
		internal const int TotalIconSize = 20;

		private static Bitmap Create(float scale, Action<Graphics> draw)
		{
			int targetSize = Math.Max((int)(TotalIconSize * scale), 1);
			int renderSize = Math.Max(DrawSize, targetSize);
			float renderScale = (float)renderSize / DrawSize;

			var hires = new Bitmap(renderSize, renderSize);
			using (var g = Graphics.FromImage(hires))
			{
				g.SmoothingMode = SmoothingMode.AntiAlias;
				g.TextRenderingHint = TextRenderingHint.AntiAliasGridFit;
				g.ScaleTransform(renderScale, renderScale);
				g.TranslateTransform(DrawPadding, DrawPadding);
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
				g.DrawString(text, font, brush, new RectangleF(0, 3, DrawArea, DrawArea), sf);
			}
		}

		internal static Bitmap Bold(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 18f, FontStyle.Bold))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "B", font, brush);
			});
		}

		internal static Bitmap Italic(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 18f, FontStyle.Italic))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "I", font, brush);
			});
		}

		internal static Bitmap Underline(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
					g.DrawLine(pen, 9, 48, 39, 48);
			});
		}

		internal static Bitmap UnderlineSingle(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f))
					g.DrawLine(pen, 6, 48, 42, 48);
			});
		}

		internal static Bitmap UnderlineDouble(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f))
				{
					g.DrawLine(pen, 6, 48, 42, 48);
					g.DrawLine(pen, 6, 42, 42, 42);
				}
			});
		}

		internal static Bitmap UnderlineDotted(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f) { DashStyle = DashStyle.Dot })
					g.DrawLine(pen, 6, 48, 42, 48);
			});
		}

		internal static Bitmap UnderlineDashed(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f) { DashStyle = DashStyle.Dash })
					g.DrawLine(pen, 6, 48, 42, 48);
			});
		}

		internal static Bitmap UnderlineWavy(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				var color = Color.FromArgb(60, 60, 60);
				using (var pen = new Pen(color, 3f))
				{
					var points = new[] {
						new PointF(6, 48), new PointF(12, 42), new PointF(18, 48), new PointF(24, 42),
						new PointF(30, 48), new PointF(36, 42), new PointF(42, 48)
					};
					g.DrawCurve(pen, points, 0.5f);
				}
			});
		}

		internal static Bitmap Strikethrough(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 12f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "ab", font, brush);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
					g.DrawLine(pen, 6, 30, 42, 30);
			});
		}

		internal static Bitmap Subscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 9f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("x", font, brush, 0, -12);
					g.DrawString("2", small, brush, 27, 26);
				}
			});
		}

		internal static Bitmap Superscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 9f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("x", font, brush, 0, -12);
					g.DrawString("2", small, brush, 27, -11);
				}
			});
		}

		internal static Bitmap FontColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 16.5f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, 3, -12);
				using (var brush = new SolidBrush(Color.Red))
					g.FillRectangle(brush, 3, 39, 42, 9);
			});
		}

		internal static Bitmap BackgroundColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var bgBrush = new SolidBrush(Color.Yellow))
					g.FillRectangle(bgBrush, 3, 3, 42, 42);
				using (var font = new Font("Segoe UI", 12f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
						g.DrawString("ab", font, brush, new RectangleF(0, 0, DrawArea, DrawArea), sf);
				}
				using (var pen = new Pen(Color.FromArgb(120, 120, 120), 3f))
					g.DrawRectangle(pen, 0, 0, 45, 45);
			});
		}

		internal static Bitmap AlignLeft(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
				{
					g.DrawLine(pen, 6, 9, 42, 9);
					g.DrawLine(pen, 6, 18, 30, 18);
					g.DrawLine(pen, 6, 27, 39, 27);
					g.DrawLine(pen, 6, 36, 27, 36);
				}
			});
		}

		internal static Bitmap AlignCenter(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
				{
					g.DrawLine(pen, 6, 9, 42, 9);
					g.DrawLine(pen, 12, 18, 36, 18);
					g.DrawLine(pen, 9, 27, 39, 27);
					g.DrawLine(pen, 15, 36, 33, 36);
				}
			});
		}

		internal static Bitmap AlignRight(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
				{
					g.DrawLine(pen, 6, 9, 42, 9);
					g.DrawLine(pen, 18, 18, 42, 18);
					g.DrawLine(pen, 9, 27, 42, 27);
					g.DrawLine(pen, 21, 36, 42, 36);
				}
			});
		}

		internal static Bitmap AlignJustify(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
				{
					g.DrawLine(pen, 6, 9, 42, 9);
					g.DrawLine(pen, 6, 18, 42, 18);
					g.DrawLine(pen, 6, 27, 42, 27);
					g.DrawLine(pen, 6, 36, 42, 36);
				}
			});
		}

		internal static Bitmap FontSizeIncrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 18f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -6, -6);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f))
				{
					g.DrawLine(pen, 36, 6, 36, 24);
					g.DrawLine(pen, 27, 15, 45, 15);
				}
			});
		}

		internal static Bitmap FontSizeDecrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 15f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -6, 0);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4.5f))
					g.DrawLine(pen, 30, 15, 45, 15);
			});
		}

		internal static Bitmap Hyperlink(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 4.5f))
				{
					g.DrawArc(pen, 3, 12, 24, 24, 90, 180);
					g.DrawLine(pen, 15, 12, 24, 12);
					g.DrawLine(pen, 15, 36, 24, 36);

					g.DrawArc(pen, 21, 12, 24, 24, -90, 180);
					g.DrawLine(pen, 24, 12, 33, 12);
					g.DrawLine(pen, 24, 36, 33, 36);
				}
			});
		}

		internal static Bitmap InsertImage(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f))
					g.DrawRectangle(pen, 3, 6, 39, 33);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 6, 9, 36, 30);

				using (var brush = new SolidBrush(Color.FromArgb(255, 200, 0)))
					g.FillEllipse(brush, 9, 12, 12, 12);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(6, 39), new Point(18, 24), new Point(27, 33), new Point(33, 27), new Point(42, 39) };
					g.FillPolygon(brush, points);
				}
			});
		}

		internal static Bitmap InsertLinkedThumbnail(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 3f))
					g.DrawRectangle(pen, 3, 6, 30, 27);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 6, 9, 27, 24);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(6, 33), new Point(15, 21), new Point(21, 27), new Point(27, 21), new Point(33, 33) };
					g.FillPolygon(brush, points);
				}

				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 4.5f))
				{
					g.DrawArc(pen, 30, 27, 12, 15, 90, 180);
					g.DrawLine(pen, 36, 27, 39, 27);
					g.DrawLine(pen, 36, 42, 39, 42);
					g.DrawArc(pen, 36, 27, 12, 15, -90, 180);
				}
			});
		}

		internal static Bitmap BulletList(float scale)
		{
			return Create(scale, g =>
			{
				var color = Color.FromArgb(60, 60, 60);
				using (var bulletBrush = new SolidBrush(color))
				using (var linePen = new Pen(color, 4.5f))
				{
					g.FillEllipse(bulletBrush, 3, 6, 9, 9);
					g.DrawLine(linePen, 18, 9, 42, 9);
					g.FillEllipse(bulletBrush, 3, 21, 9, 9);
					g.DrawLine(linePen, 18, 24, 42, 24);
					g.FillEllipse(bulletBrush, 3, 36, 9, 9);
					g.DrawLine(linePen, 18, 39, 42, 39);
				}
			});
		}

		internal static Bitmap OrderedList(float scale)
		{
			return Create(scale, g =>
			{
				var textColor = Color.FromArgb(60, 60, 60);
				var lineColor = Color.FromArgb(100, 100, 100);
				using (var font = new Font("Segoe UI", 6f, FontStyle.Bold))
				using (var brush = new SolidBrush(textColor))
				using (var linePen = new Pen(lineColor, 6f))
				{
					g.DrawString("1", font, brush, 0, 0);
					g.DrawLine(linePen, 18, 9, 42, 9);

					g.DrawString("2", font, brush, 0, 15);
					g.DrawLine(linePen, 18, 24, 42, 24);

					g.DrawString("3", font, brush, 0, 30);
					g.DrawLine(linePen, 18, 39, 42, 39);
				}
			});
		}
	}
}
