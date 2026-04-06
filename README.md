# 📁 Folder2Clipboard

**Folder2Clipboard** is a powerful Command Line Interface (CLI) tool written in Python. It scans a specified directory, extracts text content from various file types (including source code, plain text, Jupyter Notebooks, and complex document formats), and merges everything directly into your system's Clipboard.

This is an incredibly useful tool for developers, researchers, and users who need to quickly gather context from multiple files to feed into AI/LLMs (like ChatGPT, Claude, Gemini) without blowing up their token limits.

## ✨ Key Features

- **Recursive Directory Scanning:** Reads all files within subdirectories.
- **Smart Formatting & Multi-Format Support:**
  - **Default:** Reads plain Text files, Source Code, Markdown, etc.
  - **Document Extraction (Flags):** Extracts text from `.pdf`, `.docx`, `.xlsx`, and `.pptx` files.
  - **Jupyter Notebooks (`.ipynb`):** Parses JSON structure into clean Markdown, formatting code blocks and preserving text outputs.
- **Advanced Line Limiting:** Target specific files (like massive `.csv` or `.log` files) using `.gitignore`-style syntax and cap the number of lines extracted to save token space.
- **Powerful Exclusion Syntax:** Ignore unwanted files/folders (e.g., `node_modules/`, `*.pyc`).
- **Image Placeholder:** Automatically detects images/base64 strings in documents or notebooks and replaces them with an `[IMAGE PLACEHOLDER]` tag to prevent massive clipboard bloat.
- **Table Extraction:** Converts tables in Word/Excel into basic Markdown Table formats.

## ⚙️ Installation (From Source)

1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/Folder2Clipboard.git
   cd Folder2Clipboard
   ```

2. (Recommended) Create a virtual environment:
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. Install the required dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## 🚀 Usage

Basic syntax:

```bash
python main.py <folder_path> [options]
```

### Supported Flags:

- `-e` or `--exclude`: List of patterns to exclude (separated by spaces, e.g., `node_modules/ .env`).
- `-l` or `--limit`: Limit the number of lines extracted for specific files. Syntax: `-l <pattern> <lines>`. Can be used multiple times.
- `-d` or `--docx`: Enable scanning for `.docx` files.
- `-p` or `--pdf`: Enable scanning for `.pdf` files.
- `-x` or `--excel`: Enable scanning for `.xlsx` files.
- `-t` or `--pptx`: Enable scanning for `.pptx` files.
- `-i` or `--ipynb`: Enable clean scanning and markdown conversion for `.ipynb` files.

*(Tip: You can combine boolean flags together, e.g., `-dpxti`)*

### Examples:

**1. Copy all standard text/code files, ignoring `node_modules` and `.env`:**
```bash
python main.py ./my_project -e node_modules/ .env
```

**2. Copy a data folder, but limit all `.csv` files to the first 100 lines and `.json` to 50 lines:**
```bash
python main.py ./data_folder -l *.csv 100 -l *.json 50
```

**3. Copy AI project files, cleanly parsing Jupyter Notebooks and ignoring checkpoints:**
```bash
python main.py ./ai_project -i -e .ipynb_checkpoints/
```

**4. The "Kitchen Sink" - Copy everything, enable ALL document formats, exclude Python caches, and limit log files:**
```bash
python main.py ./mega_folder -e __pycache__/ *.pyc -dpxti -l *.log 200
```

## 📦 Standalone Executable (Release)

If you don't want to install Python or dependencies, you can download the standalone executable (an `.exe` file for Windows or a binary for macOS/Linux) from the **[Releases](https://github.com/nguyenduc13475/Folder2Clipboard/releases/tag/v1.0.0)** page.

How to run the standalone build via Terminal:

```bash
./folder2clip ./my_folder -dpxti -l *.csv 50
```

-----

**License:** MIT License