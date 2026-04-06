# 📁 Folder2Clipboard

**Folder2Clipboard** is a powerful Command Line Interface (CLI) tool written in Python. It scans a specified directory, extracts text content from files (including source code, plain text, and complex document formats), and merges everything directly into your system's Clipboard.

This is an incredibly useful tool for developers and users who need to quickly gather context from multiple files to feed into AI/LLMs (like ChatGPT, Claude, Gemini).

## ✨ Key Features

- **Recursive Directory Scanning:** Reads all files within subdirectories.
- **Multi-Format Support:**
  - Default: Reads plain Text files, Source Code, Markdown, etc.
  - Optional Flags: Extracts content from `.pdf`, `.docx`, `.xlsx`, and `.pptx` files.
- **Powerful Exclusion Syntax:** Supports ignoring files/folders similar to `.gitignore` (e.g., `node_modules/`, `*.pyc`).
- **Image Replacement:** Automatically inserts an `[IMAGE PLACEHOLDER]` tag wherever an image is found in supported documents.
- **Table Extraction:** Converts tables in Word/Excel into basic Markdown Table formats.

## ⚙️ Installation (From Source)

1. Clone the repository:
   ```bash
   git clone [https://github.com/your-username/Folder2Clipboard.git](https://github.com/your-username/Folder2Clipboard.git)
   cd Folder2Clipboard
    ```

2.  (Recommended) Create a virtual environment:

    ```bash
    python -m venv venv
    source venv/bin/activate  # On Windows: venv\Scripts\activate
    ```

3.  Install the required dependencies:

    ```bash
    pip install -r requirements.txt
    ```

## 🚀 Usage

Basic syntax:

```bash
python main.py <folder_path> [options]
```

### Supported Flags:

  - `-e` or `--exclude`: List of patterns to exclude (separated by spaces).
  - `-d` or `--docx`: Enable scanning for `.docx` files.
  - `-p` or `--pdf`: Enable scanning for `.pdf` files.
  - `-x` or `--excel`: Enable scanning for `.xlsx` files.
  - `-t` or `--pptx`: Enable scanning for `.pptx` files.

*(Tip: You can combine flags together, e.g., `-dpxt`)*

### Examples:

**1. Copy all text/code files in a project, ignoring `node_modules` and `.env`:**

```bash
python main.py ./my_project -e node_modules/ .env
```

**2. Copy a documents folder, enabling PDF and Word extraction:**

```bash
python main.py ./documents -d -p
```

**3. Copy everything, enable ALL document formats, and exclude Python cache files:**

```bash
python main.py ./data_folder -e __pycache__/ *.pyc -dpxt
```

## 📦 Standalone Executable (Release)

If you don't want to install Python or dependencies, you can download the standalone executable (an `.exe` file for Windows or a binary for macOS/Linux) from the **[Releases](https://github.com/nguyenduc13475/Folder2Clipboard/releases/tag/v1.0.0)** page.

How to run the standalone build via Terminal:

```bash
./folder2clip ./my_folder -dpxt
```

-----

**License:** MIT License