#!/usr/bin/env python3

import os
import glob
import re
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches
from PIL import Image

def natural_sort_key(filename):
    """
    Generate a key for natural sorting of filenames with numbers.
    This ensures '2.png' comes before '10.png'
    """
    def convert(text):
        return int(text) if text.isdigit() else text.lower()
    
    return [convert(c) for c in re.split('([0-9]+)', filename)]

def create_presentation_from_images(images_folder="images", output_file="presentation.pptx"):
    """
    Create a PowerPoint presentation from images in a folder.
    Each image will cover the entire slide area.
    
    Args:
        images_folder (str): Path to the folder containing images
        output_file (str): Name of the output PowerPoint file
    """
    
    prs = Presentation()
    
    image_extensions = ['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp', '*.tiff']
    image_files = []
    
    for extension in image_extensions:
        image_files.extend(glob.glob(os.path.join(images_folder, extension)))
        image_files.extend(glob.glob(os.path.join(images_folder, extension.upper())))
    
    # Natural sorting by filename
    image_files.sort(key=lambda x: natural_sort_key(os.path.basename(x)))
    
    if not image_files:
        print(f"No image files found in {images_folder} folder.")
        return
    
    print(f"Found {len(image_files)} image(s). Creating presentation...")
    print("Processing order:")
    for i, img_file in enumerate(image_files):
        print(f"  {i+1}. {os.path.basename(img_file)}")
    print()
    
    for i, image_file in enumerate(image_files):
        print(f"Processing slide {i+1}: {os.path.basename(image_file)} - COVERING ENTIRE SLIDE")
        
        slide_layout = prs.slide_layouts[6]  # Blank layout
        slide = prs.slides.add_slide(slide_layout)
        
        slide_width = prs.slide_width
        slide_height = prs.slide_height
        
        try:
            # Get image dimensions
            with Image.open(image_file) as img:
                img_width, img_height = img.size
            
            # Calculate aspect ratios
            slide_ratio = slide_width / slide_height
            img_ratio = img_width / img_height
            
            # Calculate scaling to cover entire slide (may crop image)
            if img_ratio > slide_ratio:
                # Image is wider relative to slide, scale by height
                scale_factor = slide_height / img_height
                final_height = slide_height
                final_width = int(img_width * scale_factor)
                
                # Center horizontally
                left = (slide_width - final_width) // 2
                top = 0
            else:
                # Image is taller relative to slide, scale by width
                scale_factor = slide_width / img_width
                final_width = slide_width
                final_height = int(img_height * scale_factor)
                
                # Center vertically
                left = 0
                top = (slide_height - final_height) // 2
            
            # Ensure minimum coverage by using the larger dimension if needed
            if final_width < slide_width or final_height < slide_height:
                width_scale = slide_width / img_width
                height_scale = slide_height / img_height
                scale_factor = max(width_scale, height_scale)
                
                final_width = int(img_width * scale_factor)
                final_height = int(img_height * scale_factor)
                
                left = (slide_width - final_width) // 2
                top = (slide_height - final_height) // 2
            
            # Add the picture to cover the entire slide
            picture = slide.shapes.add_picture(
                image_file,
                left=left,
                top=top,
                width=final_width,
                height=final_height
            )
            
            print(f"  Image scaled to {final_width}x{final_height} at position ({left}, {top})")
            
        except Exception as e:
            print(f"Error adding image {image_file}: {str(e)}")
            continue
    
    try:
        prs.save(output_file)
        print(f"Presentation saved successfully as '{output_file}'")
        print(f"Created {len(prs.slides)} slides from {len(image_files)} images")
        print("Images now properly cover entire slides with aspect ratio preservation")
    except Exception as e:
        print(f"Error saving presentation: {str(e)}")

def main():
    """Main function to run the script"""
    
    if not os.path.exists("images"):
        print("Error: 'images' folder not found in current directory.")
        return
    
    create_presentation_from_images()

if __name__ == "__main__":
    main() 