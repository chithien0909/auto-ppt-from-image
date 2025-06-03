# Auto PowerPoint Presentation Generator

A powerful tool that automatically creates professional PowerPoint presentations from images and converts them to PDF with flexible image fitting options.

## Features

- **Interactive Menu Interface**: User-friendly menu system for easy navigation
- **Custom Folder Selection**: Choose any folder containing images to create presentations
- **Multiple Slide Formats**: Support for both 4:3 and 16:9 slide formats
- **Flexible Image Fitting**: Three fitting modes to suit your needs:
  - **Contain**: Show entire image with possible margins
  - **Cover**: Fill entire slide with possible image cropping
  - **Stretch**: Fill entire slide by stretching image
- **PDF Conversion**: Convert presentations to PDF using multiple methods
- **Cross-Platform Support**: Works on Windows, macOS, and Linux
- **Smart Natural Sorting**: Processes files in logical order (1.png, 2.png, 10.png)
- **Multiple Image Formats**: Supports PNG, JPG, JPEG, GIF, BMP, TIFF formats

## Installation

### Prerequisites

- Python 3.6 or higher
- pip (Python package installer)

### Step 1: Clone the Repository

```bash
git clone https://github.com/yourusername/auto-ppt-from-image.git
cd auto-ppt-from-image
```

### Step 2: Install Required Packages

```bash
pip install -r requirements.txt
```

This will install the following dependencies:
- python-pptx: For PowerPoint file creation
- Pillow: For image processing
- reportlab: For PDF generation

### Step 3: Install LibreOffice (Optional, for PDF Conversion)

For the best PDF conversion quality, install LibreOffice:

#### Windows
- Download and install from [LibreOffice website](https://www.libreoffice.org/download/download/)

#### macOS
- Using Homebrew: `brew install --cask libreoffice`
- Or download from [LibreOffice website](https://www.libreoffice.org/download/download/)

#### Linux
- Ubuntu/Debian: `sudo apt install libreoffice`
- Fedora: `sudo dnf install libreoffice`

## Usage

### Running the Application

Start the application with:

```bash
python main.py
```

This will open the main menu with the following options:
1. Create presentation from images
2. Convert presentation to PDF
3. Exit

### Creating a Presentation

1. Select option 1 from the main menu
2. Choose to use the default 'images' folder or select a different folder
3. Enter the desired output filename (default: presentation.pptx)
4. Select slide format (4:3 or 16:9)
5. Choose image fit mode:
   - Contain: Show entire image (may have margins)
   - Cover: Fill entire slide (may crop image)
   - Stretch: Fill entire slide (may distort image)

### Converting to PDF

1. Select option 2 from the main menu
2. Choose a PowerPoint file from the list
3. The application will attempt multiple conversion methods:
   - LibreOffice (if installed)
   - Platform-specific methods (unoconv, PowerPoint, Keynote)
   - Pure Python conversion (basic)

## Project Structure

```
auto-ppt-from-image/
├── main.py                  # Main entry point
├── create_presentation.py   # PowerPoint generation
├── ppt_converter.py         # PDF conversion functionality
├── ui_manager.py            # User interface handling
├── requirements.txt         # Python dependencies
├── images/                  # Default folder for images
│   └── (your images here)
└── README.md                # This file
```

## Image Fitting Modes

### Contain Mode
- Shows the entire image without cropping
- May have margins on sides or top/bottom
- Preserves aspect ratio
- Best when you need to see the whole image

### Cover Mode
- Image fills the entire slide with no margins
- May crop parts of the image to fit
- Preserves aspect ratio
- Best for a professional, full-screen look

### Stretch Mode
- Image fills the entire slide with no margins
- Stretches or compresses the image to fit exactly
- May distort the image
- Best when exact dimensions are critical

## PDF Conversion Methods

The application attempts multiple methods for PDF conversion in this order:

1. **LibreOffice**: High-quality conversion (recommended)
2. **Platform-specific**: 
   - Windows: PowerPoint automation
   - macOS: Keynote/AppleScript
   - Linux: unoconv
3. **Pure Python**: Basic conversion with placeholders

## Troubleshooting

### LibreOffice Not Found

If LibreOffice is installed but not detected:

#### Windows
- Ensure LibreOffice is installed in the standard location
- Add LibreOffice to your PATH environment variable

#### macOS
- If installed with Homebrew, run: `brew link --force libreoffice`
- Add to PATH: `export PATH=$PATH:/Applications/LibreOffice.app/Contents/MacOS`

#### Linux
- Install with: `sudo apt install libreoffice`
- Check if installed: `which libreoffice` or `which soffice`

### Image Folder Issues

- Ensure your images folder contains supported image formats
- Check file permissions
- For relative paths, run the script from the correct directory

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request. 