#!/usr/bin/env python3

import sys
from create_presentation import create_presentation_from_images
from ppt_converter import convert_pptx_to_pdf
from ui_manager import create_slides_menu, convert_pdf_menu, main_menu

def main():
    """Main entry point for the application."""
    try:
        # Create wrapper functions to pass to the UI manager
        def create_slides():
            create_slides_menu(create_presentation_from_images)
        
        def convert_pdf():
            convert_pdf_menu(convert_pptx_to_pdf)
        
        # Start the main menu
        main_menu(create_slides, convert_pdf)
        
    except KeyboardInterrupt:
        print("\n\nProgram terminated by user.")
        sys.exit(0)

if __name__ == "__main__":
    main() 