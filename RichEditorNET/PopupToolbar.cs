using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using EllipticBit.RichEditorNET.Formatting;
using EllipticBit.RichEditorNET.Support;
using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET
{
	internal sealed class PopupToolbar : Form
	{
		private readonly RichEditBox _editor;
		private readonly ToolbarIconCache _icons;
		private readonly ToolStrip _topStrip;
		private readonly ToolStrip _bottomStrip;
		private readonly float _dpiScale;
		private Point _targetScreenPoint;
		private bool _shown;
		private int _blockingOperations;
		private int _showGeneration;

		private ToolStripComboBox _fontCombo;
		private ToolStripComboBox _sizeCombo;
		private ToolStripComboBox _blockStyleCombo;
		private Color _lastFontColor = Color.Black;
		private Color _lastBackgroundColor = Color.Yellow;
		private bool _updatingState;

		internal PopupToolbar(RichEditBox editor, ToolbarIconCache icons)
		{
			_editor = editor;
			_icons = icons;
			_dpiScale = icons.DpiScale;

			FormBorderStyle = FormBorderStyle.None;
			ShowInTaskbar = false;
			StartPosition = FormStartPosition.Manual;
			BackColor = Color.FromArgb(245, 245, 245);
			KeyPreview = true;
			DoubleBuffered = true;
			Padding = new Padding(1);

			_topStrip = CreateTopStrip();
			_bottomStrip = CreateBottomStrip();

			_topStrip.Dock = DockStyle.Top;
			_bottomStrip.Dock = DockStyle.Top;

			Controls.Add(_bottomStrip);
			Controls.Add(_topStrip);

			var totalWidth = Math.Max(_topStrip.PreferredSize.Width, _bottomStrip.PreferredSize.Width) + Padding.Horizontal;
			var totalHeight = _topStrip.PreferredSize.Height + _bottomStrip.PreferredSize.Height + Padding.Vertical;
			ClientSize = new Size(totalWidth, totalHeight);
		}

		protected override CreateParams CreateParams
		{
			get
			{
				var cp = base.CreateParams;
				cp.ClassStyle |= 0x00020000; // CS_DROPSHADOW
				return cp;
			}
		}

		protected override void OnLoad(EventArgs e)
		{
			base.OnLoad(e);
			PositionOnScreen();
			RefreshState();
		}

		internal void ShowAt(Point screenPoint)
		{
			_showGeneration++;
			_targetScreenPoint = screenPoint;

			if (!_shown)
			{
				_shown = true;
				Show(_editor.FindForm());
			}
			else
			{
				PositionOnScreen();
				RefreshState();
				Show();
			}
		}

		private void PositionOnScreen()
		{
			var screen = Screen.FromPoint(_targetScreenPoint);
			int x = _targetScreenPoint.X;
			int y = _targetScreenPoint.Y - Height - 4;

			x = Math.Max(screen.WorkingArea.Left, Math.Min(x, screen.WorkingArea.Right - Width));
			if (y < screen.WorkingArea.Top)
				y = _targetScreenPoint.Y + 4;

			Location = new Point(x, y);
		}

		protected override void OnPaint(PaintEventArgs e)
		{
			base.OnPaint(e);
			using (var pen = new Pen(Color.FromArgb(200, 200, 200)))
				e.Graphics.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
		}

		protected override void SetVisibleCore(bool value)
		{
			if (!value && Visible && _editor.IsHandleCreated)
			{
				_editor.TextDocument.Freeze();
				try
				{
					base.SetVisibleCore(value);
				}
				finally
				{
					_editor.TextDocument.Unfreeze();
				}
			}
			else
			{
				base.SetVisibleCore(value);
			}
		}

		protected override void OnDeactivate(EventArgs e)
		{
			base.OnDeactivate(e);
			if (_blockingOperations > 0) return;
			int gen = _showGeneration;
			BeginInvoke(new Action(() =>
			{
				if (gen == _showGeneration && _blockingOperations == 0 && !Disposing && !IsDisposed)
					Hide();
			}));
		}

		protected override void OnKeyDown(KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				Hide();
				e.Handled = true;
				return;
			}
			base.OnKeyDown(e);
		}

			#region - Top Strip -

		private ToolStrip CreateTopStrip()
		{
			var strip = CreateStrip();

			_fontCombo = new ToolStripComboBox
			{
				DropDownStyle = ComboBoxStyle.DropDown,
				AutoSize = false,
				Width = 250,
				FlatStyle = FlatStyle.System,
				Visible = _editor.EnableFontName,
			};
			foreach (var family in FontFamily.Families)
				_fontCombo.Items.Add(family.Name);
			_fontCombo.SelectedIndexChanged += OnFontFamilyChanged;
			_fontCombo.KeyDown += OnFontComboKeyDown;
			_fontCombo.DropDown += (s, e) => _blockingOperations++;
			_fontCombo.DropDownClosed += (s, e) =>
			{
				_blockingOperations--;
				BeginInvoke(new Action(() =>
				{
					if (_blockingOperations == 0 && !ContainsFocus && !Disposing && !IsDisposed)
						Hide();
				}));
			};
			strip.Items.Add(_fontCombo);

			_blockStyleCombo = new ToolStripComboBox
			{
				DropDownStyle = ComboBoxStyle.DropDownList,
				AutoSize = false,
				Width = 150,
				FlatStyle = FlatStyle.System,
				Visible = _editor.EnableHtmlFontSizing,
			};
			foreach (var name in BlockStyleHelper.BlockStyleNames)
				_blockStyleCombo.Items.Add(name);
			_blockStyleCombo.SelectedIndexChanged += OnBlockStyleChanged;
			_blockStyleCombo.DropDown += (s, e) => _blockingOperations++;
			_blockStyleCombo.DropDownClosed += (s, e) =>
			{
				_blockingOperations--;
				BeginInvoke(new Action(() =>
				{
					if (_blockingOperations == 0 && !ContainsFocus && !Disposing && !IsDisposed)
						Hide();
				}));
			};
			strip.Items.Add(_blockStyleCombo);

			_sizeCombo = new ToolStripComboBox
			{
				DropDownStyle = ComboBoxStyle.DropDown,
				AutoSize = false,
				Width = 100,
				FlatStyle = FlatStyle.System,
				Visible = _editor.EnableFontSize,
			};
			foreach (var size in new[] { "8", "9", "10", "11", "12", "14", "16", "18", "20", "22", "24", "26", "28", "36", "48", "72" })
				_sizeCombo.Items.Add(size);
			_sizeCombo.SelectedIndexChanged += OnFontSizeChanged;
			_sizeCombo.KeyDown += OnSizeComboKeyDown;
			_sizeCombo.DropDown += (s, e) => _blockingOperations++;
			_sizeCombo.DropDownClosed += (s, e) =>
			{
				_blockingOperations--;
				BeginInvoke(new Action(() =>
				{
					if (_blockingOperations == 0 && !ContainsFocus && !Disposing && !IsDisposed)
						Hide();
				}));
			};
			strip.Items.Add(_sizeCombo);

			strip.Items.Add(CreateButton(_icons.FontSizeIncrease, "Increase Font Size", _editor.EnableFontSize, (s, e) => _editor.IncreaseFontSize()));
			strip.Items.Add(CreateButton(_icons.FontSizeDecrease, "Decrease Font Size", _editor.EnableFontSize, (s, e) => _editor.DecreaseFontSize()));

			bool hasFontGroup = _editor.EnableFontName || _editor.EnableFontSize || _editor.EnableHtmlFontSizing;
			bool hasInsertGroup = _editor.EnableHyperlinks || _editor.EnableImages;
			bool hasListGroup = _editor.EnableLists;
			AddSeparator(strip, hasFontGroup && (hasInsertGroup || hasListGroup));

			strip.Items.Add(CreateButton(_icons.Hyperlink, "Insert Hyperlink", _editor.EnableHyperlinks, (s, e) => _editor.RaiseInsertHyperlinkClicked()));

			strip.Items.Add(CreateButton(_icons.InsertImage, "Insert Image", _editor.EnableImages && _editor.RequireEmbeddedImage, (s, e) => _editor.RaiseInsertImageClicked()));
			strip.Items.Add(CreateButton(_icons.InsertLinkedThumbnail, "Insert Linked Thumbnail", _editor.EnableImages && !_editor.RequireEmbeddedImage, (s, e) => _editor.RaiseInsertLinkedThumbnailClicked()));

			AddSeparator(strip, hasInsertGroup && hasListGroup);

			strip.Items.Add(CreateButton(_icons.BulletList, "Bullet List", _editor.EnableLists, (s, e) => { _editor.ToggleList(ListStyle.Bullet); RefreshState(); }));
			strip.Items.Add(CreateButton(_icons.OrderedList, "Ordered List", _editor.EnableLists, (s, e) => { _editor.ToggleList(ListStyle.Decimal); RefreshState(); }));

			return strip;
		}

		#endregion

		#region - Bottom Strip -

		private ToolStrip CreateBottomStrip()
		{
			var strip = CreateStrip();

			strip.Items.Add(CreateToggleButton(_icons.Bold, "Bold", _editor.EnableBold, (s, e) => { _editor.ToggleBold(); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.Italic, "Italic", _editor.EnableItalic, (s, e) => { _editor.ToggleItalic(); RefreshState(); }));
			strip.Items.Add(CreateUnderlineSplitButton());
			strip.Items.Add(CreateToggleButton(_icons.Strikethrough, "Strikethrough", _editor.EnableStrikeThrough, (s, e) => { _editor.ToggleStrikeThrough(); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.Subscript, "Subscript", _editor.EnableSubscript, (s, e) => { _editor.ToggleSubscript(); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.Superscript, "Superscript", _editor.EnableSuperscript, (s, e) => { _editor.ToggleSuperscript(); RefreshState(); }));

			bool hasTextGroup = _editor.EnableBold || _editor.EnableItalic || _editor.EnableUnderline || _editor.EnableStrikeThrough || _editor.EnableSubscript || _editor.EnableSuperscript;
			bool hasColorGroup = _editor.EnableBackgroundColor || _editor.EnableFontColor;
			bool hasAlignGroup = _editor.EnableAlignment;
			AddSeparator(strip, hasTextGroup && (hasColorGroup || hasAlignGroup));

			strip.Items.Add(CreateColorSplitButton(_icons.BackgroundColor, "Background Color", _editor.EnableBackgroundColor,
				color => { _lastBackgroundColor = color; _editor.SetBackgroundColor(color); },
				() => _lastBackgroundColor));
			strip.Items.Add(CreateColorSplitButton(_icons.FontColor, "Font Color", _editor.EnableFontColor,
				color => { _lastFontColor = color; _editor.SetFontColor(color); },
				() => _lastFontColor));

			AddSeparator(strip, hasColorGroup && hasAlignGroup);

			strip.Items.Add(CreateToggleButton(_icons.AlignLeft, "Align Left", _editor.EnableAlignment, (s, e) => { _editor.SetAlignment(ParagraphAlignment.Left); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.AlignCenter, "Align Center", _editor.EnableAlignment, (s, e) => { _editor.SetAlignment(ParagraphAlignment.Center); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.AlignRight, "Align Right", _editor.EnableAlignment, (s, e) => { _editor.SetAlignment(ParagraphAlignment.Right); RefreshState(); }));
			strip.Items.Add(CreateToggleButton(_icons.AlignJustify, "Align Justify", _editor.EnableAlignment, (s, e) => { _editor.SetAlignment(ParagraphAlignment.Justify); RefreshState(); }));

			return strip;
		}

		private ToolStripItem CreateUnderlineSplitButton()
		{
			var button = new ToolStripSplitButton
			{
				Image = _icons.Underline,
				ToolTipText = "Underline",
				DisplayStyle = ToolStripItemDisplayStyle.Image,
				Visible = _editor.EnableUnderline,
			};

			button.ButtonClick += (s, e) =>
			{
				_editor.ToggleUnderline(UnderlineStyle.Single);
				RefreshState();
			};

			button.DropDownItems.Add(CreateUnderlineDropDownItem("Single", _icons.UnderlineSingle, UnderlineStyle.Single));
			button.DropDownItems.Add(CreateUnderlineDropDownItem("Double", _icons.UnderlineDouble, UnderlineStyle.Double));
			button.DropDownItems.Add(CreateUnderlineDropDownItem("Dotted", _icons.UnderlineDotted, UnderlineStyle.Dotted));
			button.DropDownItems.Add(CreateUnderlineDropDownItem("Dashed", _icons.UnderlineDashed, UnderlineStyle.Dashed));
			button.DropDownItems.Add(CreateUnderlineDropDownItem("Wavy", _icons.UnderlineWavy, UnderlineStyle.Wavy));

			return button;
		}

		private ToolStripMenuItem CreateUnderlineDropDownItem(string text, Bitmap icon, UnderlineStyle style)
		{
			var item = new ToolStripMenuItem(text, icon, (s, e) =>
			{
				_editor.ToggleUnderline(style);
				RefreshState();
			});
			return item;
		}

		#endregion

		#region - State Refresh -

		private void RefreshState()
		{
			var range = _editor.TextDocument.Range2(_editor.SelectionStart, _editor.SelectionStart + _editor.SelectionLength);
			try
			{
				var font = range.Font;
				if (font != null)
				{
					try
					{
						_updatingState = true;
						try
						{
							if (_editor.EnableFontName)
							{
								try { _fontCombo.Text = font.Name; }
								catch { /* ignore COM errors for undefined state */ }
							}

							if (_editor.EnableFontSize)
							{
								try
								{
									float size = font.Size;
									_sizeCombo.Text = size > 0 ? size.ToString("0.##") : "";
								}
								catch { /* ignore COM errors for undefined state */ }
							}
						}
						finally
						{
							_updatingState = false;
						}

						if (_editor.EnableHtmlFontSizing)
						{
							try
							{
								var blockStyle = _editor.GetBlockStyle();
								int idx = (int)blockStyle;
								_updatingState = true;
								try { _blockStyleCombo.SelectedIndex = idx; }
								finally { _updatingState = false; }
							}
							catch { /* ignore COM errors for undefined state */ }
						}

						SetChecked("Bold", font.Bold == tomConstants.tomTrue);
						SetChecked("Italic", font.Italic == tomConstants.tomTrue);
						SetChecked("Underline", font.Underline != tomConstants.tomNone && font.Underline != tomConstants.tomUndefined && font.Underline > 0);
						SetChecked("Strikethrough", font.StrikeThrough == tomConstants.tomTrue);
						SetChecked("Subscript", font.Subscript == tomConstants.tomTrue);
						SetChecked("Superscript", font.Superscript == tomConstants.tomTrue);
					}
					finally
					{
						Marshal.ReleaseComObject(font);
					}
				}

				var para = range.Para;
				if (para != null)
				{
					try
					{
						int alignment = para.Alignment;
						SetChecked("Align Left", alignment == (int)ParagraphAlignment.Left);
						SetChecked("Align Center", alignment == (int)ParagraphAlignment.Center);
						SetChecked("Align Right", alignment == (int)ParagraphAlignment.Right);
						SetChecked("Align Justify", alignment == (int)ParagraphAlignment.Justify);

						int listType = para.ListType & 0xFFFF;
						SetChecked("Bullet List", listType == (int)ListStyle.Bullet);
						SetChecked("Ordered List", listType >= (int)ListStyle.Decimal);
					}
					finally
					{
						Marshal.ReleaseComObject(para);
					}
				}
			}
			finally
			{
				Marshal.ReleaseComObject(range);
			}
		}

		private void SetChecked(string toolTipText, bool isChecked)
		{
			SetCheckedInStrip(_topStrip, toolTipText, isChecked);
			SetCheckedInStrip(_bottomStrip, toolTipText, isChecked);
		}

		private static void SetCheckedInStrip(ToolStrip strip, string toolTipText, bool isChecked)
		{
			foreach (ToolStripItem item in strip.Items)
			{
				if (item.ToolTipText == toolTipText)
				{
					if (item is ToolStripButton btn)
						btn.Checked = isChecked;
					else if (item is ToolStripSplitButton split)
						split.BackColor = isChecked ? SystemColors.GradientActiveCaption : strip.BackColor;
					break;
				}
			}
		}

		#endregion

		#region - Event Handlers -

		private void OnFontFamilyChanged(object sender, EventArgs e)
		{
			if (_updatingState) return;
			if (_fontCombo.SelectedItem is string name && !string.IsNullOrEmpty(name))
			{
				_editor.SetFontName(name);
				_editor.Focus();
			}
		}

		private void OnFontComboKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				string name = _fontCombo.Text;
				if (!string.IsNullOrEmpty(name))
				{
					_editor.SetFontName(name);
					_editor.Focus();
				}
				e.Handled = true;
			}
		}

		private void OnFontSizeChanged(object sender, EventArgs e)
		{
			if (_updatingState) return;
			ApplyFontSize();
		}

		private void OnSizeComboKeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Enter)
			{
				ApplyFontSize();
				e.Handled = true;
			}
		}

		private void ApplyFontSize()
		{
			if (float.TryParse(_sizeCombo.Text, out float size) && size > 0)
			{
				_editor.SetFontSize(size);
				_editor.Focus();
			}
		}

		private void OnBlockStyleChanged(object sender, EventArgs e)
		{
			if (_updatingState) return;
			if (_blockStyleCombo.SelectedIndex >= 0)
			{
				_editor.SetBlockStyle((BlockStyle)_blockStyleCombo.SelectedIndex);
				RefreshState();
				_editor.Focus();
			}
		}

		#endregion

		#region - Helpers -

		private ToolStrip CreateStrip()
		{
			int scaledSize = Math.Max((int)(ToolbarIcons.TotalIconSize * _dpiScale), 1);
			return new ToolStrip
			{
				GripStyle = ToolStripGripStyle.Hidden,
				BackColor = Color.FromArgb(245, 245, 245),
				Renderer = new ToolStripProfessionalRenderer(new PopupToolbarColorTable()),
				Padding = new Padding(2, 1, 2, 1),
				ImageScalingSize = new Size(scaledSize, scaledSize),
			};
		}

		private ToolStripButton CreateButton(Bitmap icon, string toolTip, bool visible, EventHandler onClick)
		{
			var button = new ToolStripButton
			{
				Image = icon,
				ToolTipText = toolTip,
				DisplayStyle = ToolStripItemDisplayStyle.Image,
				Visible = visible,
			};
			button.Click += onClick;
			return button;
		}

		private ToolStripButton CreateToggleButton(Bitmap icon, string toolTip, bool visible, EventHandler onClick)
		{
			var button = new ToolStripButton
			{
				Image = icon,
				ToolTipText = toolTip,
				DisplayStyle = ToolStripItemDisplayStyle.Image,
				CheckOnClick = false,
				Visible = visible,
			};
			button.Click += onClick;
			return button;
		}

		private static void AddSeparator(ToolStrip strip, bool visible = true)
		{
			strip.Items.Add(new ToolStripSeparator { Visible = visible });
		}

		private ToolStripSplitButton CreateColorSplitButton(Bitmap icon, string toolTip, bool visible,
			Action<Color> applyColor, Func<Color> getLastColor)
		{
			var button = new ToolStripSplitButton
			{
				Image = icon,
				ToolTipText = toolTip,
				DisplayStyle = ToolStripItemDisplayStyle.Image,
				Visible = visible,
			};

			button.ButtonClick += (s, e) => applyColor(getLastColor());

			button.DropDownOpening += (s, e) => _blockingOperations++;
			button.DropDownClosed += (s, e) =>
			{
				_blockingOperations--;
				BeginInvoke(new Action(() =>
				{
					if (_blockingOperations == 0 && !ContainsFocus && !Disposing && !IsDisposed)
						Hide();
				}));
			};

			AddColorItem(button, "Black", Color.Black, applyColor);
			AddColorItem(button, "Dark Gray", Color.FromArgb(64, 64, 64), applyColor);
			AddColorItem(button, "Gray", Color.Gray, applyColor);
			AddColorItem(button, "Silver", Color.Silver, applyColor);
			AddColorItem(button, "White", Color.White, applyColor);
			AddColorItem(button, "Dark Red", Color.FromArgb(192, 0, 0), applyColor);
			AddColorItem(button, "Red", Color.Red, applyColor);
			AddColorItem(button, "Orange", Color.FromArgb(255, 165, 0), applyColor);
			AddColorItem(button, "Yellow", Color.Yellow, applyColor);
			AddColorItem(button, "Light Green", Color.FromArgb(146, 208, 80), applyColor);
			AddColorItem(button, "Green", Color.Green, applyColor);
			AddColorItem(button, "Cyan", Color.Cyan, applyColor);
			AddColorItem(button, "Light Blue", Color.FromArgb(0, 176, 240), applyColor);
			AddColorItem(button, "Blue", Color.Blue, applyColor);
			AddColorItem(button, "Dark Blue", Color.FromArgb(0, 32, 96), applyColor);
			AddColorItem(button, "Purple", Color.FromArgb(112, 48, 160), applyColor);
			AddColorItem(button, "Magenta", Color.Magenta, applyColor);
			AddColorItem(button, "Pink", Color.FromArgb(255, 102, 153), applyColor);
			AddColorItem(button, "Teal", Color.Teal, applyColor);
			AddColorItem(button, "Navy", Color.Navy, applyColor);

			button.DropDownItems.Add(new ToolStripSeparator());
			button.DropDownItems.Add(new ToolStripMenuItem("More Colors...", null, (s, e) =>
			{
				_blockingOperations++;
				try
				{
					using (var dialog = new ColorDialog { FullOpen = true })
					{
						if (dialog.ShowDialog(this) == DialogResult.OK)
							applyColor(dialog.Color);
					}
				}
				finally
				{
					_blockingOperations--;
				}
			}));

			return button;
		}

		private static void AddColorItem(ToolStripSplitButton button, string name, Color color, Action<Color> applyColor)
		{
			button.DropDownItems.Add(new ToolStripMenuItem(name, CreateColorSwatch(color), (s, e) => applyColor(color)));
		}

		private static Bitmap CreateColorSwatch(Color color)
		{
			var swatch = new Bitmap(16, 16);
			using (var g = Graphics.FromImage(swatch))
			{
				using (var brush = new SolidBrush(color))
					g.FillRectangle(brush, 0, 0, 16, 16);
				using (var pen = new Pen(Color.FromArgb(180, 180, 180)))
					g.DrawRectangle(pen, 0, 0, 15, 15);
			}
			return swatch;
		}

		#endregion

		private sealed class PopupToolbarColorTable : ProfessionalColorTable
		{
			public override Color ToolStripGradientBegin => Color.FromArgb(245, 245, 245);
			public override Color ToolStripGradientMiddle => Color.FromArgb(245, 245, 245);
			public override Color ToolStripGradientEnd => Color.FromArgb(245, 245, 245);
			public override Color ToolStripBorder => Color.Transparent;
		}
	}
}
