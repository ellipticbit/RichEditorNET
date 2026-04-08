namespace Test
{
	public partial class Form1 : Form
	{
		public Form1() {
			InitializeComponent();
		}

		private void Form1_Activated(object sender, EventArgs e) {

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

			var output = richTextBox1.Markdown;

			if (output.Equals(markdown, StringComparison.Ordinal)) {
				MessageBox.Show("Markdown matches!");
			}
		}

	}
}
