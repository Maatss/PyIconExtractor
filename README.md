# PyIconExtractor - *Get those icons*

PyIconExtractor is a Python script that uses [7zip](https://7-zip.org/) to extract icons from executable files. It is a simple script with only 3 dependencies: [7zip](https://7-zip.org/), [pywin32](), and [send2trash](https://pypi.org/project/Send2Trash/).

## Usage

1. Clone the repository or download the latest release from the [releases page](https://github.com/Maatss/PyIconExtractor).
2. Run the script using one of the following methods:
    1. Drag-and-drop your files or folder directly on the `run_script_in_virtual_env.bat` script.
    2. Run the `build_virtual_env_only.bat` and then invoke the python script with your arguments in the provided terminal.
    3. Run the `icon_extractor.py` script directly with your global python installation.

## Details

The python script will create a folder named `icons` in the same directory as the script, and will extract all the icons into that folder.

The `run_script_in_virtual_env.bat` script creates a local virtual environment, installs the dependencies, and runs the script. It is recommended to use a virtual environment, as it will not interfere with your global python installation.

| Argument | Description |
| --- | --- |
| `-h` or `--help` | Show the help message and exit |
| `-z` or `--7zip` | Absolute path to 7zip program. Defaults to guessing the path. |
| `-o` or `--output` | Absolute path to where to store extracted icons. Defaults to `./icons`. |
| `-r` or `--recursive` | Include to search recursively for icons in directories (to include sub-directories) |
| `-l` or `--largest` | Include to only extract the largest icon file from each executable file |
| `paths` | Paths to executable files or directories containing executable files which will be processed |

## Extract icons manually

If you want to extract the icons manually, you can directly use 7zip to open any executable files as an archive. The icons are stored in the `ICON` folder in the archive.

You can also use the `icon_extractor.py` script directly, but you will need to install the dependencies yourself. You can do this by running `pip install -r requirements.txt` in the root directory of the project.

The benefit of using the python script instead of 7zip directly is that the script will handle bulk processing of files and directories, as well as automatically guessing the image format of icon files without extensions.

## License

This project is licensed under the GNU GPLv3 license. See the included license file for more details.
