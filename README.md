# RichEditor.NET

[![NuGet](https://img.shields.io/nuget/v/RichEditorNET.svg)](https://www.nuget.org/packages/RichEditorNET)
[![License](https://img.shields.io/badge/license-BSL--1.0-blue.svg)](https://opensource.org/licenses/BSL-1.0)

A modern, feature-rich replacement for the traditional `RichTextBox` control in Windows Forms applications. RichEditor.NET leverages the Windows Text Object Model 2 (TOM2) via the RichEdit 4.1 control (`msftedit.dll`) to provide advanced text and image formatting, built-in spell checking, a context-menu popup toolbar, and first-class HTML and Markdown serialization.

## Features

### Text Formatting

- **Bold** / **Italic** / **Strikethrough**
- **Underline** with multiple styles: Single, Double, Dotted, Dashed, and Wavy
- **Font color** and **background (highlight) color**
- **Font name** and **font size** (free-form or HTML heading sizes)
- **Superscript** and **Subscript**

### Block & Paragraph Formatting

- **Block styles**: Paragraph, Heading 1-6, and Preformatted
- **Paragraph alignment**: Left, Center, Right, and Justify
- **Indentation**: Indent and Outdent with configurable point increments
- **Lists** with multiple styles:
  - Bullet (unordered)
  - Decimal, Lower Alpha, Upper Alpha, Lower Roman, Upper Roman (ordered)
- **Nested list levels** with increase/decrease support

### Image Support

- **Embedded images**: Insert JPEG, PNG, or GIF images directly into the document
- **Linked thumbnails**: Insert a downsampled thumbnail that links to the original full-size image URL
- **Alt text** support for inserted images
- Configurable default image dimensions (`DefaultImageWidth` / `DefaultImageHeight`)

### Serialization

- **HTML**: Read and write document content as HTML with full round-trip support for all formatting features. Optionally enable strict mode (`EnableStrictHtml`) to throw on unsupported tags.
- **Markdown**: Read and write CommonMark or GitHub Flavored Markdown (GFM) via the `EnableCommonMarkdown` and `EnableGithubMarkdown` properties.

### Other Features

- **Built-in spell checking** via the Windows system spell checker
- **Popup toolbar** that appears on right-click with formatting controls
- **Hyperlink insertion** with custom display text
- **Per-feature enable/disable properties** for fine-grained control over available formatting options
- **Events** for hyperlink, image, and linked thumbnail insertion via the popup toolbar

## Requirements

- Windows 8 or later (for spell checking and TOM2 support)
- .NET 8+ or .NET Framework 4.8

## Installation

Install via the .NET CLI:

```
dotnet add package RichEditor.NET
```

Or via the Package Manager Console in Visual Studio:

```
Install-Package RichEditor.NET
```

## Usage

### Basic Setup

Drop a `RichEditBox` onto your form or create one in code:

```csharp
using EllipticBit.RichEditorNET;

var editor = new RichEditBox
{
    Dock = DockStyle.Fill,
    EnableSpellCheck = true,
};
this.Controls.Add(editor);
```

### Applying Text Formatting

```csharp
// Toggle bold on the current selection
editor.ToggleBold();

// Toggle italic
editor.ToggleItalic();

// Apply a wavy underline
editor.ToggleUnderline(UnderlineStyle.Wavy);

// Set font color to red
editor.SetFontColor(Color.Red);

// Set background highlight to yellow
editor.SetBackgroundColor(Color.Yellow);

// Change font name and size
editor.SetFontName("Georgia");
editor.SetFontSize(14f);
```

### Working with Block Styles and Lists

```csharp
// Apply Heading 1 block style (requires EnableHtmlFontSizing = true)
editor.EnableHtmlFontSizing = true;
editor.SetBlockStyle(BlockStyle.Heading1);

// Toggle a bullet list
editor.ToggleList(ListStyle.Bullet);

// Increase nesting level
editor.IncreaseListLevel();

// Set paragraph alignment to center
editor.SetAlignment(ParagraphAlignment.Center);
```

### Inserting Hyperlinks

```csharp
editor.InsertHyperlink("https://github.com", "Visit GitHub");
```

### Inserting Images

```csharp
// Insert an embedded image from a file
using (var stream = File.OpenRead("photo.png"))
{
    editor.InsertImage(stream, "A sample photo");
}

// Insert a linked thumbnail that opens the full image URL
using (var stream = File.OpenRead("photo.png"))
{
    editor.InsertLinkedThumbnail(
        "https://example.com/photo-full.png",
        stream,
        "Thumbnail preview");
}
```

### HTML Round-Trip

```csharp
// Load HTML content
editor.Html = "<p>Hello <b>World</b></p>";

// Retrieve the document as HTML
string html = editor.Html;
```

### Markdown Round-Trip

```csharp
// Enable GitHub Flavored Markdown
editor.EnableGithubMarkdown = true;

// Load Markdown
editor.Markdown = "# Hello\n\nThis is **bold** and ~~strikethrough~~ text.";

// Retrieve the document as Markdown
string md = editor.Markdown;
```

### Markdown Mode

When `EnableCommonMarkdown` or `EnableGithubMarkdown` is set to `true`, formatting options that are not representable in the target Markdown flavor are automatically disabled. This ensures the content stays portable.

## API Reference

### Properties

| Property | Type | Default | Description |
|---|---|---|---|
| `EnableSpellCheck` | `bool` | `true` | Enables the system spell checker |
| `EnableBold` | `bool` | `true` | Enables bold formatting |
| `EnableItalic` | `bool` | `true` | Enables italic formatting |
| `EnableUnderline` | `bool` | `true` | Enables underline formatting |
| `EnableStrikeThrough` | `bool` | `true` | Enables strikethrough formatting |
| `EnableFontColor` | `bool` | `true` | Enables font color formatting |
| `EnableBackgroundColor` | `bool` | `true` | Enables background color formatting |
| `EnableFontName` | `bool` | `true` | Enables font name formatting |
| `EnableFontSize` | `bool` | `true` | Enables font size formatting |
| `EnableSuperscript` | `bool` | `true` | Enables superscript formatting |
| `EnableSubscript` | `bool` | `true` | Enables subscript formatting |
| `EnableLists` | `bool` | `true` | Enables list formatting |
| `EnableAlignment` | `bool` | `true` | Enables paragraph alignment |
| `EnableIndent` | `bool` | `true` | Enables paragraph indentation |
| `EnableHtmlFontSizing` | `bool` | `false` | Uses HTML heading sizes instead of free-form font size/name |
| `EnableHyperlinks` | `bool` | `true` | Enables hyperlink insertion |
| `EnableImages` | `bool` | `true` | Enables image insertion |
| `EnableCommonMarkdown` | `bool` | `false` | Enables CommonMark Markdown mode |
| `EnableGithubMarkdown` | `bool` | `false` | Enables GitHub Flavored Markdown mode |
| `EnableStrictHtml` | `bool` | `false` | Throws on unsupported HTML tags/attributes when loading HTML |
| `DefaultImageWidth` | `int` | `320` | Default display width in pixels for inserted images |
| `DefaultImageHeight` | `int` | `240` | Default display height in pixels for inserted images |
| `Html` | `string` | - | Gets or sets document content as HTML |
| `Markdown` | `string` | - | Gets or sets document content as Markdown |

### Events

| Event | Description |
|---|---|
| `InsertHyperlinkClicked` | Raised when the Insert Hyperlink button is clicked on the popup toolbar |
| `InsertImageClicked` | Raised when the Insert Image button is clicked on the popup toolbar |
| `InsertLinkedThumbnailClicked` | Raised when the Insert Linked Thumbnail button is clicked on the popup toolbar |

## Contributing

Contributions are welcome! To get started:

1. **Fork** this repository and create a feature branch from `master`.
2. **Make your changes** and ensure the project builds successfully for both `net8.0-windows` and `net48` targets.
3. **Add or update tests** as appropriate.
4. **Submit a Pull Request** with a clear description of what you changed and why.

### Coding Standards

- Follow the existing code style and conventions in the repository.
- Use the latest C# language features where appropriate.
- Ensure all COM objects are properly released via `Marshal.ReleaseComObject`.

### LLM-Generated Contributions

If any part of your contribution was generated with the assistance of a Large Language Model (LLM) such as ChatGPT, GitHub Copilot, Claude, or similar tools, **you must include the full prompt(s) used to generate the contribution** in the Pull Request description. This is required for transparency and review purposes.

### Reporting Issues

Please use [GitHub Issues](../../issues) to report bugs or request features. Include steps to reproduce the problem and, if possible, a minimal code sample.

## License

This project is licensed under the [Boost Software License 1.0 (BSL-1.0)](https://opensource.org/licenses/BSL-1.0).

Copyright (c) EllipticBit LLC 2026
