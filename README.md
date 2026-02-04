# Word to LaTeX Converter

Converts Word documents (.docx) to LaTeX format using pandoc.

## Requirements

- Python 3.10+
- [pandoc](https://pandoc.org/)

### Installing pandoc

**macOS (Homebrew):**
```bash
brew install pandoc
```

**Ubuntu/Debian:**
```bash
sudo apt install pandoc
```

**Windows:**
```bash
choco install pandoc
```

Or download from https://pandoc.org/installing.html

## Usage

1. Edit `convert.py` and set the input file at the top of the file:
   ```python
   INPUT_FILE = "your_document.docx"
   ```

2. Run the script:
   ```bash
   python convert.py
   ```

## Output

For an input file `document.docx`, the script creates:

```
document/
├── document.tex    # LaTeX file
├── document.zip    # Archive of .tex and media
└── media/          # Extracted images (if any)
```
