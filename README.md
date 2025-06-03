# Auto PowerPoint Presentation Generator & Google Slides Converter

This project contains Python scripts that automatically create Microsoft PowerPoint presentations from images and convert them to Google Slides **with perfect image fitting**.

## Features

- **All Images Processed**: Automatically finds and processes ALL images in the specified folder
- **One Image Per Slide**: Creates exactly one slide for each image found
- **Full Slide Coverage**: Images now COVER the entire slide area (no borders or margins)
- **Perfect Google Slides Fitting**: Direct Google Slides creation ensures images fit perfectly
- **Smart Natural Sorting**: Sorts filenames properly (1.png, 2.png, 10.png instead of 1.png, 10.png, 2.png)
- **Aspect Ratio Preserved**: Images maintain proportions while covering the slide
- **Smart Centering**: Images are perfectly centered on each slide
- **Multiple Formats**: Supports PNG, JPG, JPEG, GIF, BMP, TIFF formats
- **Processing Order Display**: Shows the exact order images will be processed
- **Professional Look**: No white space or borders around images

## Setup

1. Install the required dependencies:
   ```bash
   pip3 install -r requirements.txt
   ```

2. For Google Slides, follow the [Google API Setup Guide](GOOGLE_SETUP.md)

## Available Scripts

### 1. Direct Google Slides Creation (⭐ RECOMMENDED)

#### `create_google_slides_direct.py` - Perfect Image Fitting
Creates Google Slides presentations directly with guaranteed image fitting.

**Why use this?** ✅ No conversion issues ✅ Perfect image coverage ✅ Faster process

Usage:
```bash
# Create Google Slides directly from images
python3 create_google_slides_direct.py

# With custom name
python3 create_google_slides_direct.py --name "My Photo Album"

# From specific folder
python3 create_google_slides_direct.py --images-folder photos

# Help and options
python3 create_google_slides_direct.py --help
```

### 2. PowerPoint Generation

#### `create_presentation.py` - Local PowerPoint Creation
Creates PowerPoint presentations locally.

Usage:
```bash
python3 create_presentation.py
```

### 3. PowerPoint to Google Slides Conversion

#### `convert_to_google_slides.py` - Convert Existing PowerPoint
Converts existing PowerPoint files to Google Slides.

**Note**: May have image fitting issues - use direct creation instead!

Usage:
```bash
# Convert a PowerPoint file
python3 convert_to_google_slides.py presentation.pptx

# Convert all PowerPoint files
python3 convert_to_google_slides.py --all
```

## Recommended Workflow

```bash
# BEST: Create Google Slides directly from images
python3 create_google_slides_direct.py --name "My Photo Album"

# Alternative: Create PowerPoint first, then convert
python3 create_presentation.py
python3 convert_to_google_slides.py presentation.pptx
```

## Image Fitting Solution

### Problem with PowerPoint Conversion
When PowerPoint presentations are converted to Google Slides, image positioning and scaling may not translate properly, causing:
- Images not covering entire slides
- White borders around images
- Incorrect aspect ratios

### Solution: Direct Google Slides Creation
The `create_google_slides_direct.py` script solves this by:
1. **Skipping PowerPoint entirely**
2. **Using Google Slides API directly**
3. **Calculating precise cover dimensions**
4. **Ensuring perfect image fitting**

## Natural Sorting Feature

All scripts use **natural sorting** for filenames:

**Old alphabetical sorting:**
```
1.png, 10.png, 2.png, 3.png
```

**New natural sorting:**
```
1.png, 2.png, 3.png, 10.png
```

## Image Coverage Algorithm

The direct Google Slides script uses a **perfect cover algorithm**:
1. **Analyzes** image dimensions vs Google Slides dimensions (10" x 7.5")
2. **Calculates** scale factor using `max(width_scale, height_scale)`
3. **Ensures** complete slide coverage with no white space
4. **Centers** image perfectly
5. **Maintains** aspect ratio while covering entire slide

## File Structure

```
auto-gg-slide/
├── images/                          # Place your images here
│   ├── 1.png
│   └── 2.png
├── create_google_slides_direct.py   # ⭐ RECOMMENDED: Direct Google Slides
├── create_presentation.py           # PowerPoint generation
├── convert_to_google_slides.py      # PowerPoint conversion
├── requirements.txt                 # Python dependencies
├── README.md                       # This file
├── GOOGLE_SETUP.md                 # Google API setup guide
├── .gitignore                      # Protects sensitive credentials
├── credentials.json                # Google API credentials (you create this)
├── token.json                      # Auto-generated auth token
└── presentation.pptx               # Generated PowerPoint files
```

## Current Images

The scripts will process ALL images found in your `images` folder:
- 1.png
- 2.png

Processing order: **1.png → 2.png** (natural sorting)

## Output Comparison

### Direct Google Slides Creation ⭐
- **Perfect image fitting**: Images cover entire slides
- **No conversion issues**: Direct API control
- **Professional appearance**: No white space or borders
- **Fast creation**: Skip PowerPoint step

### PowerPoint → Google Slides Conversion
- **May have fitting issues**: Conversion can change positioning
- **Potential white borders**: Scaling may not translate properly
- **Extra step required**: Create PowerPoint first, then convert

## Google Slides Benefits

✅ **Cloud Access**: Access presentations anywhere  
✅ **Real-time Collaboration**: Share and edit with others  
✅ **Auto-save**: Never lose your work  
✅ **Version History**: Track all changes  
✅ **Easy Sharing**: Share via link or email  
✅ **Cross-platform**: Works on any device with internet  
✅ **Perfect Image Fitting**: With direct creation method  

## Security Notes

- Google API credentials are automatically protected by `.gitignore`
- Never share `credentials.json` or `token.json` files
- See [GOOGLE_SETUP.md](GOOGLE_SETUP.md) for detailed security information

## Key Improvements

✅ **Direct Google Slides Creation**: Perfect image fitting guaranteed  
✅ **Bypasses Conversion Issues**: No PowerPoint-to-Slides problems  
✅ **Full Slide Coverage**: Images cover entire slides with no borders  
✅ **Natural Filename Sorting**: Processes files in logical numeric order  
✅ **Processing Order Display**: Shows exact order before processing  
✅ **Aspect Ratio Maintained**: No distortion while covering slides  
✅ **Professional Appearance**: Clean, full-screen image presentation  
✅ **All Images Processed**: Every image in folder becomes a slide

## Coverage Behavior

- **Portrait images on landscape slides**: Top/bottom may be cropped for full coverage
- **Landscape images on portrait slides**: Left/right may be cropped for full coverage
- **Square images**: Minimal or no cropping
- **All cases**: Image covers 100% of slide area with no white space

## Customization

You can modify the scripts to:
- Change image positioning and sizing
- Add text or captions to slides
- Use different slide layouts
- Add transitions or animations
- Process images from multiple folders 