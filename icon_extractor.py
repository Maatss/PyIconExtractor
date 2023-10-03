"""Utility script for extracting icons from executable files using 7zip.

Accepts .exe files, links to .exe files, and directories containing .exe files.

Example usage:
    python icon_extractor.py 'my/exe.exe'
    python icon_extractor.py 'my/exe/folder'
    python icon_extractor.py 'my/exe.lnk'


Author: Maatss
Date: 2023-10-03
GitHub: https://github.com/Maatss/PyIconExtractor
Version: 1.0
"""

import argparse
import imghdr
import re
import shutil
import subprocess
from subprocess import CompletedProcess
from typing import Generator, List, Optional, Set, Tuple
from pathlib import Path
from send2trash import send2trash
from win32com.client import Dispatch
from win32com.client.dynamic import CDispatch


def try_find_7zip() -> Optional[Path]:
    """Try find 7zip program.

    Returns:
        Optional[Path]: Path to 7zip program, or None if not found.
    """

    # Attempt 1: Path environment variable
    path_7zip: Optional[str] = shutil.which("7z")
    if path_7zip:
        return Path(path_7zip)

    # Attempt 2: 64-bit
    path_7zip = Path("C:/Program Files/7-Zip/7z.exe")
    if path_7zip.exists():
        return path_7zip

    # Attempt 3: 32-bit
    path_7zip = Path("C:/Program Files (x86)/7-Zip/7z.exe")
    if path_7zip.exists():
        return path_7zip

    # Failed
    return None

def find_files_with_extensions(paths: List[Path],
                               file_extensions: Set[str],
                               recursive_search: bool = True,
                               resolve_shortcuts: bool = True) -> Set[Path]:
    """Find file paths with specific file types. Search always includes contents of top level directories.

    Args:
        paths (List[Path]): List of paths to search.
        file_extensions (Set[str]): Set of file extensions to search for.
        recursive_search (bool, optional): Search directories recursively. Defaults to True.
        resolve_shortcuts (bool, optional): Resolve shortcuts. Defaults to True.

    Returns:
        List[Path]: List of file paths.
    """

    # Create shell object for resolving shortcuts
    shell: CDispatch = Dispatch("WScript.Shell")

    # Process file paths
    filtered_file_paths: Set[Path] = set()
    for path in paths:
        # Skip if path does not exist
        if not path.exists():
            print(f"No file exists at given path. Skipping '{path}'.")
            continue

        # Search any directories
        filtered_generator: Generator[Path, None, None] = [path] # Default to path itself
        if path.is_dir():
            directory_generator: Generator[Path, None, None] = path.glob("**/*") if recursive_search else path.glob("*")
            filtered_generator = (subpath for subpath in directory_generator if subpath.is_file())

        # Filter paths
        for subpath in filtered_generator:
            # Collect files with sought after extensions
            if subpath.suffix.lower() in file_extensions:
                filtered_file_paths.add(subpath)
                continue

            # Handle shortcuts
            if resolve_shortcuts and subpath.suffix.lower() == ".lnk":
                # Resolve shortcut
                # https://stackoverflow.com/questions/397125/reading-the-target-of-a-lnk-file-in-python
                shortcut = shell.CreateShortCut(str(subpath))
                target_path: Path = Path(shortcut.Targetpath).resolve()

                # Collect file if it has sought after type
                if target_path.suffix.lower() in file_extensions:
                    filtered_file_paths.add(target_path)
                    continue

    return filtered_file_paths

def find_icons_in_exe(path_7zip: Path, exe_path: Path) -> List[Tuple[str, int]]:
    """Find icons in an executable file.

    Args:
        exe_path (Path): Path to executable file.

    Returns:
        Optional[List[Tuple[str, int]]]: List of icon file paths and sizes.
    """

    # List file contents using 7zip program
    list_cmd: List[str] = [str(path_7zip), "l", str(exe_path)]
    result: CompletedProcess[str] = subprocess.run(list_cmd, capture_output=True, text=True, check=False)
    output: str = result.stdout
    if result.returncode != 0:
        print(f"\t\t- Failed to list file contents, skipping '{exe_path}'.")
        return []

    # Find icon files
    icon_files: List[Tuple[str, int]] = []
    for line in output.split("\n"):
        # Search lines containing: "*\ICON\*"
        # I.e. files in the ICON directory
        match_ico_files: Optional[re.Match[str]] = re.search(r".+[\/\\]ICON[\/\\]\S+", line)

        # Skip if no match
        if not match_ico_files:
            continue

        # Extract icon file name and size
        # Example:
        #                                    Size   Compressed  Name
        # "                    .....       270398       270376  .rsrc\1033\ICON\1.ico"
        # "                    .....         2398         1376  .rsrc\1033\ICON\2"
        fields: List[str] = line.split()
        icon_file_path: str = fields[-1]
        icon_file_size: int = int(fields[-3])
        icon_files.append((icon_file_path, icon_file_size))

    # Return list sorted by size (large to small)
    icon_files.sort(key=lambda entry: entry[1], reverse=True)
    return icon_files


def extract_icon_from_exe(path_7zip: Path, exe_path: Path, inner_icon_file_path: str, output_dir_path: Path) -> Optional[Path]:
    """Extract an icon from an executable file.

    Args:
        exe_path (Path): Path to executable file.
        inner_icon_file_path (str): Path to icon file inside executable file.
        output_path (Path): Path to where to save the icon file.

    Returns:
        bool: True if successful, else False.
    """

    # Extract icon file using 7zip program
    extract_cmd: List[str] = [str(path_7zip), "e", str(exe_path), inner_icon_file_path, f"-o{str(output_dir_path)}"]
    result: CompletedProcess[str] = subprocess.run(extract_cmd, capture_output=True, text=True, check=False)
    if result.returncode != 0:
        print(f"\t\t\t- Failed to extract icon file '{inner_icon_file_path}' from '{exe_path}'.")
        return None

    # If missing a file extension, guess it from the file header
    extracted_file_path: Path = output_dir_path / Path(inner_icon_file_path).name
    if extracted_file_path.suffix.lower() == "":
        print("\t\t\t- Guessing file extension from file header.")
        extension: str = imghdr.what(extracted_file_path)
        if extension:
            extracted_file_path = extracted_file_path.rename(extracted_file_path.with_suffix(f".{extension}"))
        else:
            print(f"\t\t\t\t- Failed to guess file extension for '{extracted_file_path}'.")
            return None

    # Return the path to the extracted icon file
    return extracted_file_path


def extract_icons(path_7zip: Path,
                  exe_file_path: Path,
                  icon_files: List[Tuple[str, int]],
                  output_dir_path: Path,
                  temp_output_dir_path: Path,
                  extract_largest_only: bool) -> int:
    """Extract icons from executable files, and save them to output directory.

    Args:
        path_7zip (Path): Path to 7zip program.
        exe_file_path (Path): Path to executable file.
        icon_files (List[Tuple[str, int]]): List of icon file paths and sizes.
        output_dir_path (Path): Path to where to store extracted icons.
        temp_output_dir_path (Path): Path to temporary output directory.
        extract_largest_only (bool): Extract only the largest icon file.

    Returns:
        bool: True if successful, else False.
    """

    num_icons: int = 0
    for file_index, (inner_icon_file_path, icon_file_size) in enumerate(icon_files, start=1):
        print(f"\t\t- ({file_index}) Extracting icon named '{inner_icon_file_path}' ({icon_file_size} bytes).")
        icon_path: Optional[Path] = extract_icon_from_exe(path_7zip,
                                                            exe_file_path,
                                                            inner_icon_file_path,
                                                            temp_output_dir_path)

        # Skip if extraction failed
        if not icon_path:
            print("\t\t\t- Failed to extract icon.")
            continue

        # If extraction is successful, move icon to output directory
        target_path: Path = output_dir_path / f"{exe_file_path.stem}{icon_path.suffix}"

        # If target path already exists, add a number to the end of the file name
        i: int = 1
        original_stem: str = target_path.stem
        while target_path.exists():
            i += 1
            target_path = target_path.parent / (original_stem + f" ({i}){target_path.suffix}")
        print(f"\t\t\t- Saving icon to: '{target_path}'")
        shutil.move(icon_path, target_path)

        # Increment number of icons extracted
        num_icons += 1

        # Stop if only extract largest icon file
        if extract_largest_only:
            break

    return num_icons


def run_extraction(path_7zip: Path, input_paths: List[Path], output_dir_path: Path, extract_largest_only: bool) -> None:
    """Run extraction of icons from executable files.

    Args:
        path_7zip (Path): Path to 7zip program.
        input_paths (List[Path]): List of paths to executable files or directories containing executable files.
        output_dir_path (Path): Path to where to store extracted icons.
    """

    # Resolve paths
    input_paths: List[Path] = [path.resolve() for path in input_paths]
    print("Input paths:")
    for i, path in enumerate(input_paths, start=1):
        print(f"\t{i}. {path}")

    # Find executable files
    exe_file_paths: Set[Path] = find_files_with_extensions(input_paths, {".exe"})
    print("Found executable paths:")
    for i, file_path in enumerate(exe_file_paths, start=1):
        print(f"\t{i}. {file_path}")

    # Create temporary output directory (remove if already exists)
    temp_output_dir_path: Path = output_dir_path / "temp-output"
    if temp_output_dir_path.exists():
        send2trash(temp_output_dir_path)
    temp_output_dir_path.mkdir(parents=True)

    # Find icons in executable files
    print("Extracting icons from files...")
    num_icons: int = 0
    for file_index, exe_file_path in enumerate(exe_file_paths, start=1):
        print(f"\t{file_index}. {exe_file_path}")
        icon_files: List[Tuple[str, int]] = find_icons_in_exe(path_7zip, exe_file_path)

        # Skip if no icon files found
        if len(icon_files) == 0:
            print("\t\t- No icon files found.")
            continue

        # Extract all icons from executable file
        num_icons += extract_icons(path_7zip,
                                   exe_file_path,
                                   icon_files,
                                   output_dir_path,
                                   temp_output_dir_path,
                                   extract_largest_only)

    # Remove temporary output directory
    temp_output_dir_path.rmdir()

    # Print summary
    print(f"\nExtracted {num_icons} icons from {len(exe_file_paths)} files.")


# -----------------------------
# --          Main           --
# -----------------------------

if __name__ == "__main__":

    # Setup program arguments
    supported_types: Set[str] = {".exe", ".lnk"}
    parser = argparse.ArgumentParser(
        description="Utility script for extracting icons from executable files." +
        "Accepts .exe files, links to .exe files, and directories containing .exe files.",
        usage="python icon_extractor.py 'my/exe.exe'")
    # - 7zip path
    parser.add_argument("-z", "--7zip", dest="seven_zip", type=Path, default=None,
                        help="absolute path to 7zip program. Defaults to guessing the path.")
    # - output directory
    parser.add_argument("-o", "--output", type=Path, default=None,
                        help="absolute path to where to store extracted icons. Defaults to './icons'.")
    # - recursive search
    parser.add_argument("-r", "--recursive", action="store_true",  # Flag, if present it is set to True, else False
                        help="include to search recursively for icons in directories" +
                        " (to include sub-directories)")
    # - extract largest only
    parser.add_argument("-l", "--largest", action="store_true",  # Flag, if present it is set to True, else False
                        help="include to only extract the largest icon file from each executable file")
    # - paths
    parser.add_argument("paths", type=Path, nargs="+",
                        help="paths to executable files or directories containing executable" +
                        " files which will be processed")

    # Parse arguments
    args = parser.parse_args()

    print("Program started.")

    # Setup 7zip path
    path_7zip: Path = args.seven_zip
    if not path_7zip:
        path_7zip: Optional[Path] = try_find_7zip()
        if not path_7zip:
            print("Failed to find 7zip program. Please provide the path to it using the '-z' argument.")
            exit(1)

    # Setup output directory
    output_dir_path: Path = args.output
    if not output_dir_path:
        output_dir_path: Path = Path("./icons").resolve()

    # Setup recursive search
    recursive_search: bool = args.recursive

    # Setup extract largest only
    extract_largest_only: bool = args.largest

    # Setup input paths
    input_paths: List[Path] = args.paths

    # Run extraction
    run_extraction(path_7zip, input_paths, output_dir_path, extract_largest_only)

    print("Program finished.")
