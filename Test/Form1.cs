namespace Test
{
	public partial class Form1 : Form
	{
		public Form1() {
			InitializeComponent();
		}

		private void Form1_Load(object sender, EventArgs e) {

			var markdown =
				"This is a **bold** word and an *italic* word.\n" +
				"Here is ***bold and italic*** together.\n" +
				"And some ~~strikethrough~~ text.\n" +
				"\n" +
				"This is a second paragraph with a [hyperlink](https://github.com) in it.\n" +
				"\n" +
				"- Bullet item one\n" +
				"- Bullet item two\n" +
				"    - Nested bullet\n" +
				"    - Another nested bullet\n" +
				"- Bullet item three\n" +
				"\n" +
				"1. First ordered item\n" +
				"2. Second ordered item\n" +
				"    1. Nested ordered item\n" +
				"3. Third ordered item\n" +
				"\n" +
				"A paragraph with **bold and *nested italic* inside** formatting.\n" +
				"\n" +
				"An image link: ![Alt text](https://example.com/image.png)";

			richTextBox1.Markdown = markdown;

			var mdoutput = richTextBox1.Markdown;

			if (mdoutput.Equals(markdown, StringComparison.Ordinal)) {
				MessageBox.Show("Markdown matches!");
			}

			// --- HTML formatting test ---
			var html =
				"<h1>Heading 1</h1>" +
				"<h2>Heading 2</h2>" +
				"<h3>Heading 3</h3>" +
				"<h4>Heading 4</h4>" +
				"<h5>Heading 5</h5>" +
				"<h6>Heading 6</h6>" +

				// Bold variants
				"<p>This is <b>bold with b</b> and <strong>bold with strong</strong> " +
				"and <span style=\"font-weight: bold\">bold with CSS</span>.</p>" +

				// Italic variants
				"<p>This is <i>italic with i</i> and <em>italic with em</em> " +
				"and <span style=\"font-style: italic\">italic with CSS</span>.</p>" +

				// Underline variants (5 styles)
				"<p><u>Single underline</u> and <ins>underline with ins</ins> " +
				"and <span style=\"text-decoration: underline\">underline with CSS</span>.</p>" +
				"<p><span style=\"text-decoration-style: double; text-decoration-line: underline\">Double underline</span> " +
				"<span style=\"text-decoration-style: dotted; text-decoration-line: underline\">Dotted underline</span> " +
				"<span style=\"text-decoration-style: dashed; text-decoration-line: underline\">Dashed underline</span> " +
				"<span style=\"text-decoration-style: wavy; text-decoration-line: underline\">Wavy underline</span>.</p>" +

				// Strikethrough variants
				"<p><s>Strikethrough with s</s> and <del>strikethrough with del</del> " +
				"and <span style=\"text-decoration: line-through\">strikethrough with CSS span</span> " +
				"and <span style=\"text-decoration: line-through\">strikethrough with CSS</span>.</p>" +

				// Combined text-decoration
				"<p><span style=\"text-decoration: underline line-through\">Underline and strikethrough combined</span>.</p>" +

				// Font color
				"<p><span style=\"color: red\">Red with color style</span> " +
				"<span style=\"color: #0000FF\">Blue with hex CSS</span> " +
				"<span style=\"color: rgb(0, 128, 0)\">Green with rgb CSS</span>.</p>" +

				// Background color
				"<p><mark>Highlighted with mark</mark> " +
				"<span style=\"background-color: yellow\">Yellow background with CSS</span> " +
				"<span style=\"background-color: #FFC0CB\">Pink background with hex</span>.</p>" +

				// Font name
				"<p><span style=\"font-family: Arial\">Arial with font-family style</span> " +
				"<span style=\"font-family: 'Times New Roman'\">Times New Roman with CSS</span>.</p>" +
				"<p>Monospace tags: <code>code</code> <kbd>kbd</kbd> <span style=\"font-family: monospace\">tt</span> <samp>samp</samp>.</p>" +

				// Font size - HTML font size equivalents via CSS
				"<p><span style=\"font-size: 8pt\">Size 1</span> <span style=\"font-size: 10pt\">Size 2</span> " +
				"<span style=\"font-size: 12pt\">Size 3</span> <span style=\"font-size: 14pt\">Size 4</span> " +
				"<span style=\"font-size: 18pt\">Size 5</span> <span style=\"font-size: 24pt\">Size 6</span> " +
				"<span style=\"font-size: 36pt\">Size 7</span>.</p>" +

				// Font size - CSS units
				"<p><span style=\"font-size: 8pt\">8pt</span> " +
				"<span style=\"font-size: 16px\">16px</span> " +
				"<span style=\"font-size: 1.5em\">1.5em</span> " +
				"<span style=\"font-size: 1.2rem\">1.2rem</span> " +
				"<span style=\"font-size: 150%\">150%</span>.</p>" +

				// Font size - CSS named sizes
				"<p><span style=\"font-size: xx-small\">xx-small</span> " +
				"<span style=\"font-size: x-small\">x-small</span> " +
				"<span style=\"font-size: small\">small</span> " +
				"<span style=\"font-size: medium\">medium</span> " +
				"<span style=\"font-size: large\">large</span> " +
				"<span style=\"font-size: x-large\">x-large</span> " +
				"<span style=\"font-size: xx-large\">xx-large</span>.</p>" +

				// Superscript and Subscript
				"<p>Normal text with <sup>superscript</sup> and <sub>subscript</sub>.</p>" +

				// Preformatted text
				"<pre>Preformatted text block\n  with preserved   spacing\n    and line breaks</pre>" +

				// Hyperlinks
				"<p>A <a href=\"https://github.com\">simple hyperlink</a> " +
				"and a <a href=\"https://example.com\"><b>bold hyperlink</b></a> " +
				"and an <a href=\"https://example.com\"><i>italic hyperlink</i></a>.</p>" +

				// Alignment
				"<p style=\"text-align: left\">Left aligned paragraph.</p>" +
				"<p style=\"text-align: center\">Center aligned paragraph.</p>" +
				"<p style=\"text-align: right\">Right aligned paragraph.</p>" +
				"<p style=\"text-align: justify\">Justified paragraph with enough text to show the justification effect on a wide enough display.</p>" +

				// Unordered list with nesting
				"<ul>" +
				"<li>Bullet item one</li>" +
				"<li>Bullet item two" +
				"<ul>" +
				"<li>Nested bullet A</li>" +
				"<li>Nested bullet B</li>" +
				"</ul>" +
				"</li>" +
				"<li>Bullet item three</li>" +
				"</ul>" +

				// Ordered list with nesting
				"<ol>" +
				"<li>First ordered item</li>" +
				"<li>Second ordered item" +
				"<ol>" +
				"<li>Nested ordered item</li>" +
				"</ol>" +
				"</li>" +
				"<li>Third ordered item</li>" +
				"</ol>" +

				// Ordered list type variants
				"<ol type=\"a\"><li>Lowercase alpha</li><li>Second alpha</li></ol>" +
				"<ol type=\"A\"><li>Uppercase alpha</li><li>Second alpha</li></ol>" +
				"<ol type=\"i\"><li>Lowercase roman</li><li>Second roman</li></ol>" +
				"<ol type=\"I\"><li>Uppercase roman</li><li>Second roman</li></ol>" +

				// Nested/combined formatting
				"<p><b><i>Bold and italic</i></b> " +
				"<b><u>Bold and underline</u></b> " +
				"<i><s>Italic and strikethrough</s></i>.</p>" +
				"<p><span style=\"color: red; font-weight: bold; font-style: italic\">Red bold italic via CSS</span>.</p>" +
				"<p><b>Bold with <i>nested italic</i> inside</b> formatting.</p>" +

				// Embedded image (colorful 4x4 checkerboard PNG, displayed at 64x64)
				"<p>Embedded image: <img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAAoSURBVBhXVckxDQAACMCwaUITIpCIqxE+OJtCqY1GSKc8lHKxycXmACitJyWNw1liAAAAAElFTkSuQmCC\" alt=\"Colorful PNG\" width=\"64\" height=\"64\" /></p>" +

				// Embedded JPEG (4x4 checkerboard, displayed at 64x64)
				"<p>Embedded JPEG: <img src=\"data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAAEAAQDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDyn9ov9ovxD+zR4xtvDnhy2+1Wc/8AaO5/7b1jSv8Ajz1i/wBKi/daXe2kJzDpsLcxnZu8qLyraK2t4CiigD//2Q==\" alt=\"Checkerboard JPEG\" width=\"64\" height=\"64\" /></p>" +

				// Linked thumbnail (colorful 4x4 pattern, displayed at 64x64, wrapped in hyperlink)
				"<p>Linked thumbnail: <a href=\"https://example.com/fullsize.png\"><img src=\"data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAQAAAAECAYAAACp8Z5+AAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAAApSURBVBhXYzjBcOI/w507///3MPxP+Z/ynwGZA5ZE4dy5858BmQOSBACWFypZKMSEbwAAAABJRU5ErkJggg==\" alt=\"Linked thumb\" width=\"64\" height=\"64\" /></a></p>" +

				// Corrupted image (invalid data → should produce placeholder)
				"<p>Corrupted image: <img src=\"data:image/png;base64,dGhpcyBpcyBub3QgYSB2YWxpZCBpbWFnZQ==\" alt=\"Bad image\" width=\"64\" height=\"64\" /></p>";

			richTextBox1.Html = html;

			var htmloutput = richTextBox1.Html;

			if (htmloutput.Equals(html, StringComparison.Ordinal)) {
				MessageBox.Show("HTML matches!");
			}
		}

	}
}
