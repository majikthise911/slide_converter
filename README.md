# slide_converter

Converts PDF and PPTX lecture slides to structured, single-file HTML with embedded images. Built for aerospace engineering course materials with heavy math, diagrams, and code.

## Install

```bash
pip install git+https://github.com/majikthise911/slide_converter.git
```

Or clone and install in editable mode (changes take effect immediately):

```bash
git clone https://github.com/majikthise911/slide_converter.git
cd slide_converter
pip install -e .
```

## Usage

```bash
convert lecture.pdf                         # auto-renders pages with diagrams/math
convert lecture.pdf --render                # render ALL pages as images (largest file)
convert lecture.pdf --no-render             # text only, no page renders (smallest file)
convert lecture.pdf -o output.html          # custom output name
convert week1.pdf week2.pdf -o combined.html  # merge multiple files
```

### Render modes

| Mode | Flag | Description | Size (77-page PDF) |
|---|---|---|---|
| **Auto** (default) | *(none)* | Renders only pages with diagrams or equations | ~5 MB |
| **All** | `--render` | Renders every page as an image | ~10 MB |
| **None** | `--no-render` | Text extraction only | ~0.4 MB |

## What it does

- Detects slide titles, bullet lists, and sub-bullets from font size analysis
- Identifies equations (CambriaMath/MT-Extra fonts) and styles them separately
- Detects MATLAB/code blocks from consecutive small-font lines with code patterns
- Extracts and embeds all raster images as base64
- Auto-renders pages with vector diagrams or math as full-page images (fixes missing diagrams and garbled Unicode)
- Generates a clickable table of contents from slide titles
- Handles both PDF and PPTX formats
- Outputs a single self-contained HTML file (no external dependencies)

## Requirements

Installed automatically by pip:

- Python >= 3.9
- pymupdf
- python-pptx
- Pillow
