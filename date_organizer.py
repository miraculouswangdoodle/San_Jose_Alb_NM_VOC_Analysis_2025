"""
Date Range File Organizer
=========================
Collects .Chrom files within a user-specified date/time range from a nested
GC data folder structure and copies them into date-range-named subfolders,
organized by compound type.

Usage:
    Run the script and follow the prompts to enter the data folder path and
    date range.

        python date_organizer.py

Expected folder structure:
    base_path/
        GC/
            C2C6/        <- compound type folder (e.g., C2C6, C6C12)
                MM/      <- month folder (zero-padded)
                    DD/  <- day folder (zero-padded)
                        YYYYMMDD_HHMM_*.Chrom
            C6C12/
                ...

Output:
    For each compound type with matching files, a new subfolder is created
    inside that compound type's GC folder:

        base_path/GC/<compound_type>/YYYYMMDD_HH_to_YYYYMMDD_HH/

    Matching .Chrom files are copied there. Files that already exist in the
    destination are skipped.

Notes:
    - Date range filtering uses hour-level precision (YYYY-MM-DD HH).
      Files are matched based on the YYYYMMDD_HHMM timestamp in their
      filename; a file timestamped 23:59 will be included when the end
      hour is 23, since only the hour portion of end_dt is used for
      comparison.
    - Files are copied (not moved), so originals remain in place.

Dependencies:
    Standard library only (os, shutil, datetime, pathlib).
    No additional packages required.
"""

import os
import shutil
from datetime import datetime
from pathlib import Path


def parse_filename_datetime(filename):
    """
    Parse datetime from filename format: YYYYMMDD_HHMM_*.Chrom

    Args:
        filename (str): The .Chrom filename to parse.

    Returns:
        datetime or None: Parsed datetime object, or None if parsing fails.
    """
    try:
        parts = filename.split("_")
        if len(parts) < 2:
            return None

        date_str = parts[0]  # YYYYMMDD
        time_str = parts[1]  # HHMM

        datetime_str = f"{date_str}{time_str}"
        return datetime.strptime(datetime_str, "%Y%m%d%H%M")
    except (ValueError, IndexError):
        return None


def get_datetime_input(prompt_text):
    """
    Prompt the user to enter a datetime in YYYY-MM-DD HH format.

    Args:
        prompt_text (str): The prompt string to display.

    Returns:
        datetime: Parsed datetime object.
    """
    while True:
        user_input = input(prompt_text)
        try:
            dt = datetime.strptime(user_input, "%Y-%m-%d %H")
            return dt
        except ValueError:
            print("  Invalid format. Please use YYYY-MM-DD HH (e.g., 2025-01-01 14)\n")


def create_destination_folder_name(start_dt, end_dt):
    """
    Build a folder name string from the start and end datetimes.

    Args:
        start_dt (datetime): Start of the date range.
        end_dt (datetime): End of the date range.

    Returns:
        str: Folder name in the format YYYYMMDD_HH_to_YYYYMMDD_HH.
    """
    start_str = start_dt.strftime("%Y%m%d_%H")
    end_str = end_dt.strftime("%Y%m%d_%H")
    return f"{start_str}_to_{end_str}"


def find_chrom_files(base_path, start_dt, end_dt):
    """
    Search for .Chrom files within the specified date range across all
    compound type folders under base_path/GC/.

    The date range comparison uses the full datetime parsed from the
    filename (YYYYMMDD_HHMM), compared against start_dt and end_dt at
    hour precision. A file timestamped 23:59 will match an end_dt of
    23:00 on the same date, since end_dt has minutes set to 00.

    Args:
        base_path (Path): Path to the root data folder containing the GC subfolder.
        start_dt (datetime): Start of the date/time range (inclusive).
        end_dt (datetime): End of the date/time range (inclusive).

    Returns:
        list of tuples: Each tuple is (full_path, compound_type, filename, file_datetime).
    """
    base_path = Path(base_path)
    gc_path = base_path / "GC"

    if not gc_path.exists():
        print(f"  Warning: GC folder not found at {gc_path}")
        return []

    matching_files = []

    # Iterate through compound type folders (e.g., C2C6, C6C12)
    for compound_folder in gc_path.iterdir():
        if not compound_folder.is_dir():
            continue

        compound_type = compound_folder.name

        # Iterate through month folders (MM)
        for month_folder in compound_folder.iterdir():
            if not month_folder.is_dir():
                continue

            # Iterate through day folders (DD)
            for day_folder in month_folder.iterdir():
                if not day_folder.is_dir():
                    continue

                # Find all .Chrom files in this day folder
                for file_path in day_folder.glob("*.Chrom"):
                    filename = file_path.name
                    file_dt = parse_filename_datetime(filename)

                    # Include file if its timestamp falls within the date range
                    if file_dt and start_dt <= file_dt <= end_dt:
                        matching_files.append((file_path, compound_type, filename, file_dt))

    return matching_files


def organize_files(base_path, start_dt, end_dt):
    """
    Find and copy matching .Chrom files into date-range-named destination
    subfolders, organized by compound type.

    Args:
        base_path (Path): Path to the root data folder containing the GC subfolder.
        start_dt (datetime): Start of the date/time range (inclusive).
        end_dt (datetime): End of the date/time range (inclusive).
    """
    base_path = Path(base_path)

    print(f"\nSearching for .Chrom files between {start_dt} and {end_dt}...")

    # Find all matching files
    matching_files = find_chrom_files(base_path, start_dt, end_dt)

    if not matching_files:
        print("No matching files found.")
        return

    print(f"Found {len(matching_files)} matching file(s).")

    # Build destination folder name from date range
    dest_folder_name = create_destination_folder_name(start_dt, end_dt)

    # Group files by compound type
    files_by_compound = {}
    for file_path, compound_type, filename, file_dt in matching_files:
        if compound_type not in files_by_compound:
            files_by_compound[compound_type] = []
        files_by_compound[compound_type].append((file_path, filename))

    # -------------------------------------------------------------------------
    # Copy files to destination folders
    # -------------------------------------------------------------------------
    total_copied = 0
    total_skipped = 0
    skipped_files = []

    for compound_type, files in files_by_compound.items():
        # Destination is a new subfolder inside the compound type folder
        dest_path = base_path / "GC" / compound_type / dest_folder_name
        dest_path.mkdir(parents=True, exist_ok=True)

        print(f"\nProcessing {compound_type} ({len(files)} file(s))...")
        print(f"  Destination: {dest_path}")

        for file_path, filename in files:
            dest_file = dest_path / filename

            # Skip files that already exist in the destination
            if dest_file.exists():
                total_skipped += 1
                skipped_files.append((filename, compound_type, "File already exists"))
                print(f"  Skipped (already exists): {filename}")
                continue

            try:
                shutil.copy2(file_path, dest_file)
                total_copied += 1
                print(f"  Copied: {filename}")
            except Exception as e:
                total_skipped += 1
                skipped_files.append((filename, compound_type, str(e)))
                print(f"  Error copying {filename}: {e}")

    # -------------------------------------------------------------------------
    # Summary report
    # -------------------------------------------------------------------------
    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"  Total files found:               {len(matching_files)}")
    print(f"  Files copied:                    {total_copied}")
    print(f"  Files skipped:                   {total_skipped}")

    if skipped_files:
        print("\nSkipped files:")
        for filename, compound_type, reason in skipped_files:
            print(f"  - {compound_type}/{filename}: {reason}")

    print("\nOrganization complete!")


def main():
    """
    Main entry point. Prompts the user for the data folder path and date range,
    then runs the file organization workflow.
    """
    print("=" * 60)
    print("Date Range File Organizer for .Chrom Files")
    print("=" * 60)

    # Get source directory
    source_dir = input("\nEnter the path to the Monthly Data folder: ").strip()

    if not os.path.exists(source_dir):
        print(f"Error: Directory '{source_dir}' does not exist.")
        return

    # Get date range
    print("\nEnter date range (format: YYYY-MM-DD HH, where HH is 0-23)")
    start_dt = get_datetime_input("  Start date and hour: ")
    end_dt = get_datetime_input("  End date and hour: ")

    # Validate date range order
    if start_dt > end_dt:
        print("Error: Start date/time must be before end date/time.")
        return

    # Run file organization
    organize_files(source_dir, start_dt, end_dt)


if __name__ == "__main__":
    main()
