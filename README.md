# Excel to Clean HTML Converter for AI Analysis

A powerful VBA macro for Microsoft Excel that transforms a selected cell range into a clean, minimalist HTML table. It is specifically designed to prepare complex data structures for analysis by AI models.

## The Problem

When you use Excel's built-in "Save as Web Page" feature, it generates cluttered HTML filled with excessive styling, metadata, and Microsoft-specific tags. This "bloated" code is difficult for AI models to parse efficiently and can lead to inaccurate data interpretation.

## The Solution

This macro solves the problem by stripping away all non-essential formatting and attributes. It preserves only the core structure of the table (`<table>`, `<tr>`, `<td>`) and structural attributes (`rowspan`, `colspan`). The result is a pure HTML representation of your data, optimized for clarity and machine readability.

## Key Features

- **Minimalist Output**: Removes all styling, comments, and metadata.
- **Structure Preservation**: Keeps `rowspan` and `colspan` attributes to maintain complex table layouts.
- **Whitespace Cleaning**: Normalizes all spacing for a compact and clean output.
- **Clipboard Ready**: The final HTML code is automatically copied to your clipboard for immediate use.

## How to Use

1.  **Open Excel** and press `Alt + F11` to open the VBA editor.
2.  In the VBA editor, go to **Insert > Module** to create a new module.
3.  **Copy the code** from the `CopySelectionAsCleanHTML.vba` file in this repository.
4.  **Paste the code** into the new module window.
5.  Close the VBA editor and return to your Excel sheet.
6.  **Select the range** of cells you want to convert.
7.  Press `Alt + F8` to open the Macro dialog, select `CopySelectionAsCleanHTML`, and click **Run**.
8.  The clean HTML code is now in your clipboard, ready to be pasted into an AI prompt or any other application.

## License

This project is licensed under the MIT License.
