# VBscript for Word: Image Formatter

A simple VBScript macro for Microsoft Word that automatically formats all images from page 2 onwards.

## Features

- **Automated Filtering**: Only targets images on page 2 or later.
- **Centering**: Automatically centers the images horizontally.
- **Borders**: Adds a subtle black border (0.25pt) to each image.
- **Safe Conversion**: Converts `InlineShapes` to `Shapes` with proper wrapping to ensure consistent formatting.

## Usage

1. Open your Word document.
2. Press `Alt + F11` to open the VBA Editor.
3. Insert a new module (`Insert > Module`).
4. Copy and paste the content of `FormatImage.vb` into the module.
5. Run the macro `FormatImagesFromPage2`.

## Project Structure

- `FormatImage.vb`: The core VBScript/VBA code.
- `README.md`: Project documentation.
