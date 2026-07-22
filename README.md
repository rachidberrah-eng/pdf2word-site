# PDF to Word Converter (Arabic Math & Physics)

Convert PDF files (containing text and math/physics equations) to editable Word documents with proper formatting.

## Features

- Extract text from PDF files
- Convert mathematical equations to LaTeX format
- Generate Word documents with editable equations (OMML)
- Customize title and text formatting (font, size, color)
- Arabic RTL support

## Requirements

- Python 3.11+
- Pandoc
- Gemini API key

## Installation

```bash
# Clone the repository
git clone https://github.com/rachidberrah-eng/pdf2word-site.git
cd pdf2word-site

# Install Python dependencies
pip install -r requirements.txt

# Install Pandoc (Linux/macOS)
# Ubuntu/Debian: sudo apt-get install pandoc
# macOS: brew install pandoc
# Windows: Download from https://pandoc.org/installing.html
```

## Configuration

Set your Gemini API key:

```bash
# Linux/macOS
export GEMINI_API_KEY="your-api-key"

# Windows (PowerShell)
$env:GEMINI_API_KEY = "your-api-key"
```

## Running

```bash
python app.py
```

Then open http://localhost:5000 in your browser.

## Docker

```bash
docker build -t pdf2word .
docker run -p 5000:5000 -e GEMINI_API_KEY="your-api-key" pdf2word
```

## License

MIT
