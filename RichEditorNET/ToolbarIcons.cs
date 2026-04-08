using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;

namespace EllipticBit.RichEditorNET
{
	internal static class ToolbarIcons
	{
		private const int DrawSize = 64;
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
				g.DrawString(text, font, brush, new RectangleF(0, 4, DrawSize, DrawSize), sf);
			}
		}

		internal static Bitmap Bold(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 24f, FontStyle.Bold))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "B", font, brush);
			});
		}

		internal static Bitmap Italic(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Verdana", 24f, FontStyle.Italic))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "I", font, brush);
			});
		}

		internal static Bitmap Underline(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 6f))
					g.DrawLine(pen, 12, 64, 52, 64);
			});
		}

		internal static Bitmap UnderlineSingle(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
					g.DrawLine(pen, 8, 64, 56, 64);
			});
		}

		internal static Bitmap UnderlineDouble(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 8, 64, 56, 64);
					g.DrawLine(pen, 8, 56, 56, 56);
				}
			});
		}

		internal static Bitmap UnderlineDotted(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f) { DashStyle = DashStyle.Dot })
					g.DrawLine(pen, 8, 64, 56, 64);
			});
		}

		internal static Bitmap UnderlineDashed(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f) { DashStyle = DashStyle.Dash })
					g.DrawLine(pen, 8, 64, 56, 64);
			});
		}

		internal static Bitmap UnderlineWavy(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
					g.DrawString("U", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				var color = Color.FromArgb(60, 60, 60);
				using (var pen = new Pen(color, 4f))
				{
					var points = new[] {
						new PointF(8, 64), new PointF(16, 56), new PointF(24, 64), new PointF(32, 56),
						new PointF(40, 64), new PointF(48, 56), new PointF(56, 64)
					};
					g.DrawCurve(pen, points, 0.5f);
				}
			});
		}

		internal static Bitmap Strikethrough(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 16f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					DrawCenteredText(g, "ab", font, brush);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
					g.DrawLine(pen, 8, 40, 56, 40);
			});
		}

		internal static Bitmap Subscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 10f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("x", font, brush, 0, -16);
					g.DrawString("2", small, brush, 36, 34);
				}
			});
		}

		internal static Bitmap Superscript(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var small = new Font("Segoe UI", 10f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					g.DrawString("x", font, brush, 0, -16);
					g.DrawString("2", small, brush, 36, -10);
				}
			});
		}

		internal static Bitmap FontColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 22f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, 4, -16);
				using (var brush = new SolidBrush(Color.Red))
					g.FillRectangle(brush, 4, 52, 56, 12);
			});
		}

		internal static Bitmap BackgroundColor(float scale)
		{
			return Create(scale, g =>
			{
				using (var bgBrush = new SolidBrush(Color.Yellow))
					g.FillRectangle(bgBrush, 4, 4, 56, 56);
				using (var font = new Font("Segoe UI", 16f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
				{
					using (var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
						g.DrawString("ab", font, brush, new RectangleF(0, 0, DrawSize, DrawSize), sf);
				}
			});
		}

		internal static Bitmap AlignLeft(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 8, 12, 56, 12);
					g.DrawLine(pen, 8, 24, 40, 24);
					g.DrawLine(pen, 8, 36, 52, 36);
					g.DrawLine(pen, 8, 48, 36, 48);
				}
			});
		}

		internal static Bitmap AlignCenter(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 8, 12, 56, 12);
					g.DrawLine(pen, 16, 24, 48, 24);
					g.DrawLine(pen, 12, 36, 52, 36);
					g.DrawLine(pen, 20, 48, 44, 48);
				}
			});
		}

		internal static Bitmap AlignRight(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 8, 12, 56, 12);
					g.DrawLine(pen, 24, 24, 56, 24);
					g.DrawLine(pen, 12, 36, 56, 36);
					g.DrawLine(pen, 28, 48, 56, 48);
				}
			});
		}

		internal static Bitmap AlignJustify(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 8, 12, 56, 12);
					g.DrawLine(pen, 8, 24, 56, 24);
					g.DrawLine(pen, 8, 36, 56, 36);
					g.DrawLine(pen, 8, 48, 56, 48);
				}
			});
		}

		internal static Bitmap FontSizeIncrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 24f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -8, -8);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
				{
					g.DrawLine(pen, 48, 8, 48, 32);
					g.DrawLine(pen, 36, 20, 60, 20);
				}
			});
		}

		internal static Bitmap FontSizeDecrease(float scale)
		{
			return Create(scale, g =>
			{
				using (var font = new Font("Segoe UI", 20f, FontStyle.Regular))
				using (var brush = new SolidBrush(Color.FromArgb(60, 60, 60)))
					g.DrawString("A", font, brush, -8, 2);
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
					g.DrawLine(pen, 30, 20, 50, 20);
			});
		}

		internal static Bitmap Hyperlink(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 6f))
				{
					g.DrawArc(pen, 4, 16, 32, 32, 90, 180);
					g.DrawLine(pen, 20, 16, 28, 16);
					g.DrawLine(pen, 20, 48, 28, 48);

					g.DrawArc(pen, 28, 16, 32, 32, -90, 180);
					g.DrawLine(pen, 36, 16, 44, 16);
					g.DrawLine(pen, 36, 48, 44, 48);

					g.DrawLine(pen, 20, 32, 44, 32);
				}
			});
		}

		internal static Bitmap InsertImage(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
					g.DrawRectangle(pen, 4, 8, 52, 44);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 8, 12, 48, 40);

				using (var brush = new SolidBrush(Color.FromArgb(255, 200, 0)))
					g.FillEllipse(brush, 12, 16, 16, 16);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(8, 52), new Point(24, 32), new Point(36, 44), new Point(44, 36), new Point(56, 52) };
					g.FillPolygon(brush, points);
				}
			});
		}

		internal static Bitmap InsertLinkedThumbnail(float scale)
		{
			return Create(scale, g =>
			{
				using (var pen = new Pen(Color.FromArgb(60, 60, 60), 4f))
					g.DrawRectangle(pen, 4, 8, 40, 36);

				using (var brush = new SolidBrush(Color.FromArgb(135, 206, 250)))
					g.FillRectangle(brush, 8, 12, 36, 32);

				using (var brush = new SolidBrush(Color.FromArgb(76, 153, 0)))
				{
					var points = new[] { new Point(8, 44), new Point(20, 28), new Point(28, 36), new Point(36, 28), new Point(44, 44) };
					g.FillPolygon(brush, points);
				}

				using (var pen = new Pen(Color.FromArgb(0, 102, 204), 6f))
				{
					g.DrawArc(pen, 40, 36, 16, 20, 90, 180);
					g.DrawLine(pen, 48, 36, 52, 36);
					g.DrawLine(pen, 48, 56, 52, 56);
					g.DrawArc(pen, 48, 36, 16, 20, -90, 180);
				}
			});
		}

		internal static Bitmap BulletList(float scale)
		{
			return Create(scale, g =>
			{
				var textColor = Color.FromArgb(50, 50, 255);
				var lineColor = Color.FromArgb(60, 60, 60);
				using (var bulletBrush = new SolidBrush(textColor))
				using (var linePen = new Pen(lineColor, 6f))
				{
					g.FillEllipse(bulletBrush, 4, 8, 12, 12);
					g.DrawLine(linePen, 24, 12, 56, 12);
					g.FillEllipse(bulletBrush, 4, 28, 12, 12);
					g.DrawLine(linePen, 24, 32, 56, 32);
					g.FillEllipse(bulletBrush, 4, 48, 12, 12);
					g.DrawLine(linePen, 24, 52, 56, 52);
				}
			});
		}

		internal static Bitmap OrderedList(float scale)
		{
			return Create(scale, g =>
			{
				var textColor = Color.FromArgb(50, 50, 255);
				var lineColor = Color.FromArgb(60, 60, 60);
				using (var font = new Font("Segoe UI", 8f, FontStyle.Bold))
				using (var brush = new SolidBrush(textColor))
				using (var linePen = new Pen(lineColor, 6f))
				{
					g.DrawString("1", font, brush, 0, 0);
					g.DrawLine(linePen, 24, 12, 56, 12);

					g.DrawString("2", font, brush, 0, 20);
					g.DrawLine(linePen, 24, 32, 56, 32);

					g.DrawString("3", font, brush, 0, 40);
					g.DrawLine(linePen, 24, 52, 56, 52);
				}
			});
		}
	}
}
