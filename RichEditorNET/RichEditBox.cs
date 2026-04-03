using System;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using EllipticBit.RichEditorNET.TextObjectModel2;

namespace EllipticBit.RichEditorNET
{
	public class RichEditBox : RichTextBox
	{
		static RichEditBox()
		{
			PInvoke.LoadLibrary(PInvoke.MSFTEDIT_DLL);
		}

		protected override CreateParams CreateParams
		{
			get
			{
				var cp = base.CreateParams;
				cp.ClassName = PInvoke.RICHEDIT_CLASS;
				return cp;
			}
		}

		[DefaultValue(true)]
		[Category("Behavior")]
		[Description("Enables or disables the system spell checker for this editor.")]
		public bool EnableSpellCheck
		{
			get => _enableSpellCheck;
			set
			{
				_enableSpellCheck = value;
				if (IsHandleCreated) {
					ApplySpellCheckOption();
				}
			}
		}
		private bool _enableSpellCheck = true;

		private ITextDocument2? _textDocument;

		[Browsable(false)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		public ITextDocument2 TextDocument => _textDocument ?? throw new InvalidOperationException("The control handle has not been created.");

		protected override void OnHandleCreated(EventArgs e)
		{
			base.OnHandleCreated(e);

			IntPtr pOle = IntPtr.Zero;
			PInvoke.SendMessage(Handle, PInvoke.EM_GETOLEINTERFACE, IntPtr.Zero, ref pOle);
			if (pOle != IntPtr.Zero)
			{
				_textDocument = Marshal.GetObjectForIUnknown(pOle) as ITextDocument2;
				Marshal.Release(pOle);
			}

			PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_AND, (IntPtr)(~PInvoke.ECO_AUTOWORDSELECTION));
			ApplySpellCheckOption();
		}

		protected override void OnHandleDestroyed(EventArgs e)
		{
			if (_textDocument != null)
			{
				Marshal.ReleaseComObject(_textDocument);
				_textDocument = null;
			}

			base.OnHandleDestroyed(e);
		}

		protected override void OnDoubleClick(EventArgs e)
		{
			base.OnDoubleClick(e);

			string selected = SelectedText;
			if (selected.Length > 0 && char.IsWhiteSpace(selected[selected.Length - 1]))
			{
				base.SelectionLength = selected.TrimEnd().Length;
			}
		}

		private void ApplySpellCheckOption()
		{
			var options = (int)PInvoke.SendMessage(Handle, PInvoke.EM_GETLANGOPTIONS, IntPtr.Zero, IntPtr.Zero);
			if (_enableSpellCheck)
				options |= PInvoke.IMF_SPELLCHECKING;
			else
				options &= ~PInvoke.IMF_SPELLCHECKING;
			PInvoke.SendMessage(Handle, PInvoke.EM_SETLANGOPTIONS, IntPtr.Zero, (IntPtr)options);
		}

		protected override void OnMouseUp(MouseEventArgs e)
		{
			base.OnMouseUp(e);

			if (e.Button == MouseButtons.Left && ModifierKeys.HasFlag(Keys.Control))
			{
				PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_AND, (IntPtr)(~PInvoke.ECO_AUTOWORDSELECTION));

				if (base.SelectionLength > 0)
				{
					string selected = SelectedText;
					if (selected.Length > 0 && char.IsWhiteSpace(selected[selected.Length - 1]))
					{
						base.SelectionLength = selected.TrimEnd().Length;
					}
				}
			}
		}

		protected override void OnMouseDown(MouseEventArgs e)
		{
			if (e.Button == MouseButtons.Right)
			{
				int charIndex = GetCharIndexFromPosition(e.Location);
				SelectionStart = charIndex;
				SelectionLength = 0;
			}
			else if (e.Button == MouseButtons.Left && ModifierKeys.HasFlag(Keys.Control))
			{
				PInvoke.SendMessage(Handle, PInvoke.EM_SETOPTIONS, (IntPtr)PInvoke.ECOOP_OR, (IntPtr)PInvoke.ECO_AUTOWORDSELECTION);
			}

			base.OnMouseDown(e);
		}

		#region - Hidden Properties -
		// Hide RTF properties

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new string? Rtf
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new string? SelectedRtf
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		// Hide text formatting properties

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new HorizontalAlignment SelectionAlignment
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Color SelectionBackColor
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new bool SelectionBullet
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionCharOffset
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Color SelectionColor
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new Font? SelectionFont
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionHangingIndent
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionIndent
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new bool SelectionProtected
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int SelectionRightIndent
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int[] SelectionTabs
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		[Browsable(false)]
		[EditorBrowsable(EditorBrowsableState.Never)]
		[DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
		private new int BulletIndent
		{
			get => throw new NotSupportedException();
			set => throw new NotSupportedException();
		}

		// Hide file save/load methods

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(string path)
		{
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(string path, RichTextBoxStreamType fileType)
		{
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void LoadFile(Stream data, RichTextBoxStreamType fileType)
		{
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(string path)
		{
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(string path, RichTextBoxStreamType fileType)
		{
			throw new NotSupportedException();
		}

		[EditorBrowsable(EditorBrowsableState.Never)]
		private new void SaveFile(Stream data, RichTextBoxStreamType fileType)
		{
			throw new NotSupportedException();
		}

		#endregion
	}
}
