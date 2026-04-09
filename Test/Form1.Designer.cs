using EllipticBit.RichEditorNET;

namespace Test
{
	partial class Form1
	{
		/// <summary>
		///  Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		///  Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		///  Required method for Designer support - do not modify
		///  the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			richTextBox1 = new RichEditBox();
			SuspendLayout();
			// 
			// richTextBox1
			// 
			richTextBox1.Dock = DockStyle.Fill;
			richTextBox1.HideSelection = false;
			richTextBox1.EnableCommonMarkdown = true;
			richTextBox1.EnableGithubMarkdown = true;
			richTextBox1.Location = new Point(0, 0);
			richTextBox1.Name = "richTextBox1";
			richTextBox1.Size = new Size(1574, 929);
			richTextBox1.TabIndex = 0;
			richTextBox1.Text = "";
			// 
			// Form1
			// 
			AutoScaleDimensions = new SizeF(13F, 32F);
			AutoScaleMode = AutoScaleMode.Font;
			ClientSize = new Size(1574, 929);
			Controls.Add(richTextBox1);
			Name = "Form1";
			StartPosition = FormStartPosition.CenterScreen;
			Text = "Form1";
			Activated += Form1_Activated;
			ResumeLayout(false);
		}

		#endregion

		private RichEditBox richTextBox1;
	}
}
