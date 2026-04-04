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

		#region - Text Formatting -

		/// <summary>
		/// Toggles bold formatting on the current selection.
		/// </summary>
		public void ToggleBold() {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Bold = font.Bold == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Toggles italic formatting on the current selection.
		/// </summary>
		public void ToggleItalic() {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Italic = font.Italic == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Toggles underline formatting on the current selection.
		/// </summary>
		/// <param name="style">The underline style to apply. Defaults to <see cref="UnderlineStyle.Single"/>.</param>
		public void ToggleUnderline(UnderlineStyle style = UnderlineStyle.Single) {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Underline = font.Underline == (int)style ? tomConstants.tomNone : (int)style;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Toggles strikethrough formatting on the current selection.
		/// </summary>
		public void ToggleStrikeThrough() {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.StrikeThrough = font.StrikeThrough == tomConstants.tomTrue ? tomConstants.tomFalse : tomConstants.tomTrue;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Sets the font color on the current selection.
		/// </summary>
		/// <param name="color">The color to apply. Use <see cref="Color.Empty"/> to reset to the automatic color.</param>
		public void SetFontColor(Color color) {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.ForeColor = color.IsEmpty ? tomConstants.tomAutoColor : ColorTranslator.ToOle(color);
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Sets the background color on the current selection.
		/// </summary>
		/// <param name="color">The color to apply. Use <see cref="Color.Empty"/> to reset to the automatic color.</param>
		public void SetBackgroundColor(Color color) {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.BackColor = color.IsEmpty ? tomConstants.tomAutoColor : ColorTranslator.ToOle(color);
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Sets the font name on the current selection.
		/// </summary>
		/// <param name="name">The font family name to apply.</param>
		public void SetFontName(string name) {
			if (string.IsNullOrEmpty(name)) throw new ArgumentException("Font name cannot be null or empty.", nameof(name));
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Name = name;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Sets the font size on the current selection.
		/// </summary>
		/// <param name="size">The font size in points. Must be greater than zero.</param>
		public void SetFontSize(float size) {
			if (size <= 0) throw new ArgumentOutOfRangeException(nameof(size), "Font size must be greater than zero.");
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Size = size;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Increases the font size on the current selection by the specified number of points.
		/// </summary>
		/// <param name="points">The number of points to increase by. Defaults to 1.</param>
		public void IncreaseFontSize(float points = 1f) {
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Point increment must be greater than zero.");
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					font.Size += points;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Decreases the font size on the current selection by the specified number of points.
		/// The resulting size will not go below 1 point.
		/// </summary>
		/// <param name="points">The number of points to decrease by. Defaults to 1.</param>
		public void DecreaseFontSize(float points = 1f) {
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Point decrement must be greater than zero.");
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					float newSize = font.Size - points;
					font.Size = newSize < 1f ? 1f : newSize;
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Toggles superscript formatting on the current selection.
		/// Disables subscript if superscript is being enabled.
		/// </summary>
		public void ToggleSuperscript() {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					if (font.Superscript == tomConstants.tomTrue) {
						font.Superscript = tomConstants.tomFalse;
					}
					else {
						font.Superscript = tomConstants.tomTrue;
						font.Subscript = tomConstants.tomFalse;
					}
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Toggles subscript formatting on the current selection.
		/// Disables superscript if subscript is being enabled.
		/// </summary>
		public void ToggleSubscript() {
			var selection = TextDocument.Selection2;
			try {
				var font = selection.Font;
				try {
					if (font.Subscript == tomConstants.tomTrue) {
						font.Subscript = tomConstants.tomFalse;
					}
					else {
						font.Subscript = tomConstants.tomTrue;
						font.Superscript = tomConstants.tomFalse;
					}
				}
				finally {
					Marshal.ReleaseComObject(font);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		#endregion

		#region - Paragraph Formatting -

		/// <summary>
		/// Toggles a list on the current selection. If the selection already uses the specified
		/// list style, the list is removed. Otherwise the specified style is applied.
		/// </summary>
		/// <param name="style">The list style to apply.</param>
		public void ToggleList(ListStyle style) {
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					int currentType = para.ListType & 0xFFFF;
					if (currentType == (int)style) {
						para.ListType = tomConstants.tomListNone;
						para.ListLevelIndex = 0;
					}
					else {
						int listType = (int)style;
						if (style != ListStyle.Bullet) {
							listType |= tomConstants.tomListPeriod;
						}
						para.ListType = listType;
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Increases the list nesting level of the current selection by one.
		/// Has no effect if the selection is not in a list.
		/// </summary>
		public void IncreaseListLevel() {
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					if (para.ListType != tomConstants.tomListNone) {
						para.ListLevelIndex++;
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Decreases the list nesting level of the current selection by one.
		/// Has no effect if the selection is not in a list or is already at the top level.
		/// </summary>
		public void DecreaseListLevel() {
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					if (para.ListType != tomConstants.tomListNone && para.ListLevelIndex > 0) {
						para.ListLevelIndex--;
					}
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Sets the paragraph alignment on the current selection.
		/// </summary>
		/// <param name="alignment">The alignment to apply.</param>
		public void SetAlignment(ParagraphAlignment alignment) {
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					para.Alignment = (int)alignment;
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Increases the left indent of the current selection by the specified number of points.
		/// </summary>
		/// <param name="points">The number of points to indent by. Defaults to 36 (0.5 inch).</param>
		public void Indent(float points = 36f) {
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Indent amount must be greater than zero.");
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					para.SetIndents(para.FirstLineIndent, para.LeftIndent + points, para.RightIndent);
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		/// <summary>
		/// Decreases the left indent of the current selection by the specified number of points.
		/// The resulting indent will not go below zero.
		/// </summary>
		/// <param name="points">The number of points to outdent by. Defaults to 36 (0.5 inch).</param>
		public void Outdent(float points = 36f) {
			if (points <= 0) throw new ArgumentOutOfRangeException(nameof(points), "Outdent amount must be greater than zero.");
			var selection = TextDocument.Selection2;
			try {
				var para = selection.Para;
				try {
					float newLeft = para.LeftIndent - points;
					para.SetIndents(para.FirstLineIndent, newLeft < 0f ? 0f : newLeft, para.RightIndent);
				}
				finally {
					Marshal.ReleaseComObject(para);
				}
			}
			finally {
				Marshal.ReleaseComObject(selection);
			}
		}

		#endregion

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
