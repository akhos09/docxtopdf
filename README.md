# Vagrant Utilities

A Python script to convert `.docx` files to `.pdf` using the `spire.doc` **FREE LICENSE** (IT LEAVES A WATERMARK) module. This tool can process individual or multiple files in a directory.

## Important Licensing Notice

This project uses the `spire.doc` library for Word to PDF conversion. The library is licensed under a commercial license, and this project is using a **free evaluation version**. 

The free evaluation version may have limitations, such as watermarks on generated PDFs, reduced functionality, or time-limited access. If you are planning to use this project in production or for commercial purposes, you will need to obtain a valid commercial license for the `spire.doc` library.

You can find more details about licensing here: [Spire.doc Licensing](https://www.e-iceblue.com/Introduce/spire-doc.html)

## Table of Contents
- [Dependencies](#dependencies)
- [Installation](#installation)
- [Usage](#usage)
- [License](#license)

## Dependencies

The script requires the following libraries and tools:

1. **spire.doc**: A library for working with `.docx` files and converting them to PDF.

2. **os**: A standard Python module for interacting with the operating system (used for file handling).

3. **subprocess**: A standard Python module to execute shell commands, used for managing the installation of dependencies.

4. **tkinter**: A Python library for GUI-based file dialogs, allowing users to select files and folders easily.

5. **pip**: Python's package manager for installing dependencies.

## Installation

1. **Clone this repository**:

   ```
   git clone https://github.com/akhos09/vagrant-utilities
   cd vagrant-utilities
   ```

2. **Install dependencies**:

   If `spire.doc` is not installed, the script will prompt you to install it. To install it manually, run:

   ```
   pip install spire.doc
   ```

## Usage

1. **Run the script**:

   Open a terminal or command prompt and navigate to the directory containing the script. Run the following command:

   ```
   python convert-docx-to-pdf.py
   ```

2. **Choose your conversion method**:

   You will be prompted to choose whether to convert a single file or multiple files from a folder.

   - **For a single file**: Select the `.docx` file and enter the desired PDF filename.
   - **For multiple files**: Select the folder containing `.docx` files, and the script will convert all `.docx` files in the folder to PDFs.

3. **Converted PDFs**:

   The converted PDFs will be saved in a folder named `pdf_converted` in the current directory.

## Troubleshooting

   - Ensure that `Python` and `spire.doc` are installed and updated.
   - If the script cannot find the `.docx` file, check the file path and try again.
   - If the PDF conversion fails, make sure the `.docx` file is valid and does not contain any corrupt elements.
   - If `spire.doc` installation fails, ensure that `pip` is installed and try again.

## License

See the [LICENSE](LICENSE) file for details.

## Credits

Thanks to [spire.doc](https://www.e-iceblue.com) for providing the document conversion functionality.
