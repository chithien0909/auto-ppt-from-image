#!/usr/bin/env python3

import sys
from create_presentation import create_presentation_from_images
from ppt_converter import convert_pptx_to_pdf, find_libreoffice_executable
from ui_manager import create_slides_menu, convert_pdf_menu, main_menu

def check_libreoffice():
    """Check if LibreOffice is installed and display appropriate message."""
    libreoffice_path = find_libreoffice_executable()
    if not libreoffice_path:
        print("\nNote: LibreOffice is not installed. For best PDF conversion results, please install LibreOffice:")
        print("- Windows: Download from https://www.libreoffice.org/download/")
        print("- macOS: Run 'brew install --cask libreoffice' or download from https://www.libreoffice.org/download/")
        print("- Linux: Run 'sudo apt install libreoffice' or equivalent for your distribution")
        print("\nThe program will still work, but PDF conversion quality may be limited.\n")
        input("Press Enter to continue...")

def main():
    """Main entry point for the application."""
    try:
        # Check for LibreOffice installation
        check_libreoffice()
        
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