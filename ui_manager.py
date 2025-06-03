#!/usr/bin/env python3

import os
import sys
import glob

def clear_screen():
    """Clear the terminal screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def print_header():
    """Print the application header."""
    print("=" * 60)
    print("                AUTO PRESENTATION GENERATOR")
    print("=" * 60)
    print()

def list_directories(base_path="."):
    """
    List all directories in the specified path.
    
    Args:
        base_path (str): Base path to list directories from
        
    Returns:
        list: List of directory paths
    """
    try:
        dirs = [d for d in os.listdir(base_path) if os.path.isdir(os.path.join(base_path, d))]
        return sorted(dirs)
    except Exception as e:
        print(f"Error listing directories: {str(e)}")
        return []

def select_folder_text_based(start_path="."):
    """
    Text-based folder selection interface.
    
    Args:
        start_path (str): Starting directory path
        
    Returns:
        str: Selected folder path or empty string if canceled
    """
    current_path = os.path.abspath(start_path)
    
    while True:
        clear_screen()
        print_header()
        print("FOLDER SELECTION")
        print("-" * 60)
        print(f"Current directory: {current_path}")
        print("\nAvailable directories:")
        
        # Add special options
        dirs = [".", ".."] + list_directories(current_path)
        
        # Print directories with numbers
        for i, dir_name in enumerate(dirs, 1):
            if dir_name == ".":
                print(f"{i}. [Select current directory]")
            elif dir_name == "..":
                print(f"{i}. [Go up one level]")
            else:
                print(f"{i}. {dir_name}")
        
        print("\n0. Cancel selection")
        
        # Get user choice
        try:
            choice = input("\nSelect a directory [0-{}]: ".format(len(dirs))).strip()
            
            if not choice:
                continue
                
            if choice == "0":
                return ""
            
            idx = int(choice) - 1
            if idx < 0 or idx >= len(dirs):
                print("Invalid selection. Press Enter to continue...")
                input()
                continue
            
            selected = dirs[idx]
            
            if selected == ".":
                return current_path
            elif selected == "..":
                current_path = os.path.dirname(current_path)
            else:
                current_path = os.path.join(current_path, selected)
                
        except ValueError:
            print("Invalid input. Press Enter to continue...")
            input()

def create_slides_menu(create_presentation_func):
    """
    Display the create slides menu and handle user input.
    
    Args:
        create_presentation_func: Function to call for creating the presentation
    """
    clear_screen()
    print_header()
    print("CREATE PRESENTATION FROM IMAGES")
    print("-" * 60)
    
    # Image folder selection
    print("Select image folder:")
    print("1. Use default 'images' folder")
    print("2. Select a different folder")
    folder_choice = input("Choose option [1]: ").strip() or "1"
    
    if folder_choice == "1":
        # Use default folder
        images_folder = "images"
        # Check if images folder exists
        if not os.path.exists(images_folder):
            print(f"Error: '{images_folder}' folder not found in current directory.")
            print("Please create an 'images' folder and add your images to it.")
            input("\nPress Enter to continue...")
            return
    else:
        # Let user select a folder
        print("\nStarting folder selection...")
        images_folder = select_folder_text_based()
        
        if not images_folder:
            print("Folder selection canceled.")
            input("\nPress Enter to continue...")
            return
        
        if not os.path.exists(images_folder):
            print(f"Error: Selected folder '{images_folder}' does not exist.")
            input("\nPress Enter to continue...")
            return
        
        print(f"\nSelected folder: {images_folder}")
    
    # Check if there are images in the selected folder
    image_extensions = ['*.png', '*.jpg', '*.jpeg', '*.gif', '*.bmp', '*.tiff']
    image_files = []
    
    for extension in image_extensions:
        image_files.extend(glob.glob(os.path.join(images_folder, extension)))
        image_files.extend(glob.glob(os.path.join(images_folder, extension.upper())))
    
    if not image_files:
        print(f"\nNo image files found in '{images_folder}'.")
        print("Please select a folder containing images.")
        input("\nPress Enter to continue...")
        return
    
    print(f"\nFound {len(image_files)} image(s) in the selected folder.")
    
    # Get presentation options
    output_file = input("\nOutput filename [presentation.pptx]: ").strip() or "presentation.pptx"
    if not output_file.endswith(".pptx"):
        output_file += ".pptx"
    
    slide_format_options = {"1": "4:3", "2": "16:9"}
    print("\nSelect slide format:")
    print("1. Standard (4:3)")
    print("2. Widescreen (16:9)")
    slide_format_choice = input("Choose format [1]: ").strip() or "1"
    slide_format = slide_format_options.get(slide_format_choice, "4:3")
    
    fit_mode_options = {"1": "contain", "2": "cover", "3": "stretch"}
    print("\nSelect image fit mode:")
    print("1. Contain - Show entire image (may have margins)")
    print("2. Cover - Fill entire slide (may crop image)")
    print("3. Stretch - Fill entire slide (may distort image)")
    fit_mode_choice = input("Choose fit mode [1]: ").strip() or "1"
    fit_mode = fit_mode_options.get(fit_mode_choice, "contain")
    
    print("\nCreating presentation...")
    create_presentation_func(
        images_folder=images_folder,
        output_file=output_file,
        slide_format=slide_format,
        fit_mode=fit_mode
    )
    
    print("\nPresentation created successfully!")
    input("\nPress Enter to continue...")

def convert_pdf_menu(convert_pdf_func):
    """
    Display the convert to PDF menu and handle user input.
    
    Args:
        convert_pdf_func: Function to call for converting PPTX to PDF
    """
    clear_screen()
    print_header()
    print("CONVERT PRESENTATION TO PDF")
    print("-" * 60)
    
    # List available PPTX files
    pptx_files = [f for f in os.listdir(".") if f.endswith(".pptx")]
    
    if not pptx_files:
        print("No PowerPoint files found in the current directory.")
        input("\nPress Enter to continue...")
        return
    
    print("Available presentations:")
    for i, file in enumerate(pptx_files, 1):
        print(f"{i}. {file}")
    
    # Get file selection
    try:
        selection = input("\nSelect presentation to convert [1]: ").strip() or "1"
        idx = int(selection) - 1
        if idx < 0 or idx >= len(pptx_files):
            raise ValueError("Invalid selection")
        
        input_file = pptx_files[idx]
        output_file = input_file.rsplit(".", 1)[0] + ".pdf"
        
        # Call the conversion function
        success = convert_pdf_func(input_file, output_file)
        
        if not success:
            print("\nConversion failed. Please check the error messages above.")
    
    except (ValueError, IndexError) as e:
        print(f"\nError: {str(e)}")
    
    input("\nPress Enter to continue...")

def main_menu(create_slides_func, convert_pdf_func):
    """
    Display the main menu and handle user selection.
    
    Args:
        create_slides_func: Function to call for the create slides menu
        convert_pdf_func: Function to call for the convert PDF menu
    """
    while True:
        clear_screen()
        print_header()
        print("MAIN MENU")
        print("-" * 60)
        print("1. Create presentation from images")
        print("2. Convert presentation to PDF")
        print("3. Exit")
        print()
        
        choice = input("Select an option [1-3]: ").strip()
        
        if choice == "1":
            create_slides_func()
        elif choice == "2":
            convert_pdf_func()
        elif choice == "3" or choice.lower() in ("q", "quit", "exit"):
            print("\nThank you for using Auto Presentation Generator!")
            sys.exit(0)
        else:
            print("\nInvalid choice. Please try again.")
            input("\nPress Enter to continue...") 