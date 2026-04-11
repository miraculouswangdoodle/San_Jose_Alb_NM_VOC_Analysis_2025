"""
CYL07-A2 File Organizer
=======================
Collects CYL07-A2 calibration cylinder .Chrom files from a nested monthly
folder structure and copies them into a single flat destination folder for
easier batch processing.

Usage:
    Run the script and enter the path to the root source folder when prompted.

        python CYL07A2_organizer.py

Expected folder structure:
    root_source/
        01/          <- main folder (day or run number, zero-padded 01-31)
            01/      <- subfolder (zero-padded 01-31)
                YYYYMMDD_HHMM_CYL07-A2_*.Chrom
            02/
                ...
        02/
            ...

Output:
    A new folder named "CYL07-A2 files_<root_folder_name>" is created as a
    sibling of root_source. All matching .Chrom files are copied there in a
    flat structure, sorted by filename (chronological order). Files that
    already exist in the destination are skipped.

Dependencies:
    Standard library only (os, shutil, pathlib). No additional packages required.
"""

import os
import shutil
from pathlib import Path

def organize_cyl07a2_files():
    """
    Organizes CYL07-A2 .Chrom files from a nested folder structure into a single destination folder.
    
    Process:
    1. Searches through root_folder/01-31/01-31/ structure
    2. Finds all .Chrom files containing "CYL07-A2" in filename
    3. Copies them to a flat folder structure named "CYL07-A2 files_[rootfoldername]"
    4. Skips files that already exist and reports on all actions
    """
    
    # Phase 1: Setup and Validation
    print("=== CYL07-A2 File Organization Script ===\n")
    
    # Get root source folder from user
    while True:
        root_source = input("Enter the path to the root source folder: ").strip()
        if os.path.exists(root_source):
            break
        else:
            print("Error: Folder does not exist. Please try again.\n")
    
    # Extract root folder name and create destination folder name
    root_folder_name = Path(root_source).name
    destination_folder = os.path.join(Path(root_source).parent, f"CYL07-A2 files_{root_folder_name}")
    
    # Create destination folder if it doesn't exist
    try:
        os.makedirs(destination_folder, exist_ok=True)
        print(f"Destination folder: {destination_folder}\n")
    except Exception as e:
        print(f"Error creating destination folder: {e}")
        return
    
    # Initialize tracking lists
    found_files = []
    copied_files = []
    skipped_files = []
    error_files = []
    
    # Phase 2: File Discovery
    print("Scanning for CYL07-A2 files...")
    
    # Loop through main folders (01-31)
    for main_folder_num in range(1, 32):
        main_folder = os.path.join(root_source, f"{main_folder_num:02d}")
        
        if not os.path.exists(main_folder):
            continue
            
        # Loop through subfolders (01-31)
        for sub_folder_num in range(1, 32):
            sub_folder = os.path.join(main_folder, f"{sub_folder_num:02d}")
            
            if not os.path.exists(sub_folder):
                continue
            
            try:
                # Find all .Chrom files in this subfolder
                for file in os.listdir(sub_folder):
                    if file.endswith('.Chrom') and 'CYL07-A2' in file:
                        full_path = os.path.join(sub_folder, file)
                        found_files.append((full_path, file))
                        print(f"Found: {file} in {main_folder_num:02d}/{sub_folder_num:02d}")
            
            except PermissionError:
                print(f"Warning: Permission denied accessing folder {main_folder_num:02d}/{sub_folder_num:02d}")
            except Exception as e:
                print(f"Warning: Error accessing folder {main_folder_num:02d}/{sub_folder_num:02d}: {e}")
    
    if not found_files:
        print("No CYL07-A2 .Chrom files found in the specified folder structure.")
        return
    
    print(f"\nFound {len(found_files)} CYL07-A2 files total.\n")
    
    # Phase 3: File Processing
    print("Processing files...")
    
    # Sort files by filename to maintain date order
    found_files.sort(key=lambda x: x[1])
    
    for file_path, filename in found_files:
        destination_path = os.path.join(destination_folder, filename)
        
        # Check if file already exists in destination
        if os.path.exists(destination_path):
            skipped_files.append(filename)
            print(f"Skipped (already exists): {filename}")
            continue
        
        # Copy file to destination
        try:
            shutil.copy2(file_path, destination_path)
            copied_files.append(filename)
            print(f"Copied: {filename}")
        
        except Exception as e:
            error_files.append((filename, str(e)))
            print(f"Error copying {filename}: {e}")
    
    # Phase 4: Completion Report
    print("\n" + "="*50)
    print("PROCESSING COMPLETE")
    print("="*50)
    
    print(f"Total CYL07-A2 files found: {len(found_files)}")
    print(f"Files successfully copied: {len(copied_files)}")
    print(f"Files skipped (already existed): {len(skipped_files)}")
    print(f"Files with errors: {len(error_files)}")
    
    if skipped_files:
        print(f"\nSkipped files:")
        for filename in skipped_files:
            print(f"  - {filename}")
    
    if error_files:
        print(f"\nFiles with errors:")
        for filename, error in error_files:
            print(f"  - {filename}: {error}")
    
    print(f"\nDestination folder: {destination_folder}")
    print("\nFile organization complete!")

if __name__ == "__main__":
    try:
        organize_cyl07a2_files()
    except KeyboardInterrupt:
        print("\n\nScript interrupted by user.")
    except Exception as e:
        print(f"\nUnexpected error: {e}")
    
    input("\nPress Enter to exit...")
