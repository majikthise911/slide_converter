"""
convert — PDF/PPTX to structured HTML with embedded images.

Uses font analysis for heading/body/math detection, bullet structure,
and equation identification. Auto-renders pages with diagrams or equations
as images to guarantee visual accuracy. Outputs single-file HTML with
base64-embedded images.

Usage:
    convert lecture.pdf                     # smart render (auto-detects pages)
    convert lecture.pdf --render            # render ALL pages as images
    convert lecture.pdf --no-render         # text-only (smallest file)
    convert lecture.pdf -o out.html         # custom output name
    convert file1.pdf file2.pdf -o all.html # merge multiple files
"""
import sys
import os
import base64
import html as html_mod
from pathlib import Path
from collections import Counter

try:
    import pymupdf
except ImportError:
    print("Installing pymupdf...")
    os.system(f"{sys.executable} -m pip install -q pymupdf")
    import pymupdf

try:
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False

# Fonts that indicate math/equation content
MATH_FONTS = ("CambriaMath", "Cambria Math", "MT-Extra", "Symbol")

# Render modes
RENDER_AUTO = "auto"     # render pages with diagrams/math (default)
RENDER_ALL = "all"       # render every page
RENDER_NONE = "none"     # no renders

CSS = """\
body { font-family: -apple-system, 'Segoe UI', Arial, sans-serif; max-width: 960px; margin: 40px auto; padding: 20px; line-height: 1.7; color: #1a1a1a; }
h1 { border-bottom: 2px solid #333; padding-bottom: 10px; margin-bottom: 10px; }
nav { margin-bottom: 30px; padding: 15px; background: #f8f9fa; border-radius: 6px; }
nav summary { font-weight: bold; cursor: pointer; }
nav ol { columns: 2; column-gap: 2em; margin: 10px 0 0 0; padding-left: 1.5em; }
nav li { margin: 2px 0; font-size: 0.92em; break-inside: avoid; }
nav a { color: #2471a3; text-decoration: none; }
nav a:hover { text-decoration: underline; }
.slide { margin-bottom: 2em; }
.slide-title { color: #1a5276; margin-top: 2em; padding: 4px 0; border-bottom: 1px solid #ddd; }
ul { margin: 0.3em 0; padding-left: 1.8em; }
li { margin: 0.2em 0; line-height: 1.6; }
.math { font-family: 'Cambria Math', 'STIX Two Math', 'Latin Modern Math', serif; }
.eq { display: block; margin: 0.6em 1.5em; padding: 6px 12px; background: #f0f4f8; border-left: 3px solid #2980b9; font-family: 'Cambria Math', serif; font-size: 1.05em; white-space: pre-wrap; }
img { max-width: 100%; height: auto; margin: 1em 0; display: block; }
.slide-img { max-width: 100%; border: 1px solid #ddd; border-radius: 4px; margin: 0.5em 0; }
.label { font-size: 0.88em; color: #555; margin: 0.3em 0; }
pre, code { background: #f5f5f5; font-family: Menlo, Consolas, monospace; font-size: 0.9em; }
pre { padding: 12px; border-radius: 4px; overflow-x: auto; }
code { padding: 2px 5px; border-radius: 3px; }
table { border-collapse: collapse; margin: 1em 0; width: 100%; }
th, td { border: 1px solid #ccc; padding: 6px 10px; text-align: left; }
th { background: #f0f0f0; font-weight: bold; }
details.render { margin: 0.5em 0; }
details.render summary { cursor: pointer; color: #888; font-size: 0.82em; }
"""


def esc(text):
    return html_mod.escape(text)


def is_math(font_name):
    return any(m in font_name for m in MATH_FONTS)


def spans_to_html(spans):
    """Render a list of text spans to HTML with math/bold/italic detection."""
    parts = []
    for s in spans:
        t = s["text"]
        if not t:
            continue
        e = esc(t)
        font = s["font"]
        if is_math(font):
            parts.append(f'<span class="math">{e}</span>')
        elif "Bold" in font and "Italic" in font:
            parts.append(f"<strong><em>{e}</em></strong>")
        elif "Bold" in font:
            parts.append(f"<strong>{e}</strong>")
        elif "Italic" in font:
            parts.append(f"<em>{e}</em>")
        else:
            parts.append(e)
    return "".join(parts)


def strip_bullet(html_text, char):
    """Remove first occurrence of a bullet character from HTML text."""
    for c in [char, esc(char)]:
        if c in html_text:
            return html_text.replace(c, "", 1).strip()
    return html_text.strip()


def render_page(page, page_num):
    """Render a page to a base64-encoded PNG image tag."""
    pix = page.get_pixmap(dpi=120)
    b64 = base64.b64encode(pix.tobytes("png")).decode()
    return (
        f'<img class="slide-img" src="data:image/png;base64,{b64}" '
        f'alt="Slide {page_num + 1}">'
    )


def page_needs_render(page):
    """Check if a page has significant vector content or math that needs rendering."""
    # Check for vector drawings (diagrams)
    drawings = page.get_drawings()
    has_diagrams = len(drawings) > 4

    # Check for math fonts (garbled Unicode in browsers)
    has_math = False
    for block in page.get_text("dict")["blocks"]:
        if block["type"] != 0:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                if span["text"].strip() and is_math(span["font"]):
                    has_math = True
                    break
            if has_math:
                break
        if has_math:
            break

    return has_diagrams or has_math


# ---------------------------------------------------------------------------
# PDF extraction
# ---------------------------------------------------------------------------

def analyze_pdf_fonts(doc):
    """First pass: determine adaptive font-size thresholds from the document."""
    size_chars = Counter()

    for page in doc:
        for block in page.get_text("dict")["blocks"]:
            if block["type"] != 0:
                continue
            for line in block["lines"]:
                for span in line["spans"]:
                    txt = span["text"].strip()
                    if txt:
                        size_chars[round(span["size"])] += len(txt)

    if not size_chars:
        return {"title": 40, "body": 24}

    sorted_sizes = sorted(size_chars.keys(), reverse=True)
    body = size_chars.most_common(1)[0][0]
    title = sorted_sizes[0]

    return {"title": title, "body": body}


def postprocess_parts(parts):
    """Post-process HTML parts: merge equations, detect code blocks."""
    result = []
    i = 0
    while i < len(parts):
        # --- Merge consecutive equation divs ---
        if parts[i].startswith('<div class="eq">'):
            eq_contents = []
            j = i
            while j < len(parts) and parts[j].startswith('<div class="eq">'):
                inner = parts[j][len('<div class="eq">'):-len("</div>")]
                eq_contents.append(inner)
                j += 1
            if len(eq_contents) > 1:
                result.append('<div class="eq">' + " ".join(eq_contents) + "</div>")
            else:
                result.append(parts[i])
            i = j
            continue

        # --- Detect code blocks in consecutive label lines ---
        if parts[i].startswith('<p class="label">') and parts[i].endswith("</p>"):
            j = i
            label_lines = []
            while (
                j < len(parts)
                and parts[j].startswith('<p class="label">')
                and parts[j].endswith("</p>")
            ):
                inner = parts[j][len('<p class="label">'):-len("</p>")]
                label_lines.append(inner)
                j += 1
            code_indicators = sum(
                1 for line in label_lines if ";" in line or "=" in line or "(" in line
            )
            if len(label_lines) >= 3 and code_indicators >= 2:
                result.append(
                    "<pre><code>" + "\n".join(label_lines) + "</code></pre>"
                )
            else:
                result.extend(parts[i:j])
            i = j
            continue

        result.append(parts[i])
        i += 1
    return result


def pdf_page_html(page, page_num, doc, fs):
    """Convert a single PDF page to structured HTML parts (text only)."""
    parts = []
    d = page.get_text("dict")
    page_h = d["height"]

    list_depth = 0
    current_li = None
    title_emitted = False
    slide_title = None

    def close_li():
        nonlocal current_li
        if current_li is not None:
            parts[current_li] += "</li>"
            current_li = None

    def close_lists():
        nonlocal list_depth
        close_li()
        while list_depth > 0:
            parts.append("</ul>")
            list_depth -= 1

    for block in d["blocks"]:
        if block["type"] != 0:
            continue

        for line in block["lines"]:
            spans = [s for s in line["spans"] if s["text"].strip()]
            if not spans:
                continue

            text = "".join(s["text"] for s in spans).strip()
            if not text:
                continue

            max_sz = max(s["size"] for s in spans)
            y = line["bbox"][1]
            all_math = all(is_math(s["font"]) for s in spans)
            html_t = spans_to_html(spans)

            # --- Skip page numbers ---
            if max_sz < 15 and y > page_h * 0.80 and text.strip().isdigit():
                continue

            # --- Slide title ---
            if max_sz >= fs["title"] - 2 and not title_emitted:
                close_lists()
                slide_title = text
                parts.append(
                    f'<h2 class="slide-title" id="slide-{page_num + 1}">{html_t}</h2>'
                )
                title_emitted = True
                continue

            # --- Top-level bullet ---
            if len(text) > 0 and text[0] in "\u2022\u2023\u25cf\u25cb":
                close_li()
                while list_depth > 1:
                    parts.append("</ul>")
                    list_depth -= 1
                if list_depth == 0:
                    parts.append("<ul>")
                    list_depth = 1
                current_li = len(parts)
                parts.append(f"<li>{strip_bullet(html_t, text[0])}")
                continue

            # --- Sub-bullet ---
            if len(text) > 0 and text[0] in "\u2013\u2014\u2012":
                close_li()
                if list_depth == 0:
                    parts.append("<ul>")
                    list_depth = 1
                if list_depth == 1:
                    parts.append("<ul>")
                    list_depth = 2
                current_li = len(parts)
                parts.append(f"<li>{strip_bullet(html_t, text[0])}")
                continue

            # --- Equation block ---
            if all_math:
                close_li()
                parts.append(f'<div class="eq">{html_t}</div>')
                continue

            # --- Small text ---
            if max_sz < fs["body"] - 6 and max_sz < fs["title"] - 10:
                close_lists()
                parts.append(f'<p class="label">{html_t}</p>')
                continue

            # --- Body / continuation ---
            if current_li is not None:
                parts[current_li] += " " + html_t
            elif list_depth > 0:
                current_li = len(parts)
                parts.append(f"<li>{html_t}")
            else:
                parts.append(f"<p>{html_t}</p>")

    close_lists()

    # --- Embed extracted raster images ---
    for idx, img_info in enumerate(page.get_images()):
        try:
            xref = img_info[0]
            bi = doc.extract_image(xref)
            b64 = base64.b64encode(bi["image"]).decode()
            parts.append(
                f'<img src="data:image/{bi["ext"]};base64,{b64}" '
                f'alt="Slide {page_num + 1} Figure {idx + 1}">'
            )
        except Exception:
            pass

    parts = postprocess_parts(parts)
    return parts, slide_title


def convert_pdf(pdf_path, render_mode=RENDER_AUTO):
    """Convert a PDF file to HTML content."""
    doc = pymupdf.open(str(pdf_path))
    fs = analyze_pdf_fonts(doc)
    n = len(doc)
    print(f"  {n} pages | title={fs['title']}pt body={fs['body']}pt")

    toc_entries = []
    body_parts = []
    rendered_count = 0

    for pn in range(n):
        page = doc[pn]
        page_html, title = pdf_page_html(page, pn, doc, fs)

        # Determine if this page needs a visual render
        should_render = False
        if render_mode == RENDER_ALL:
            should_render = True
        elif render_mode == RENDER_AUTO:
            should_render = page_needs_render(page)

        if should_render:
            page_html.append(render_page(page, pn))
            rendered_count += 1

        if title:
            toc_entries.append((pn + 1, title))
        elif not toc_entries or toc_entries[-1][0] != pn + 1:
            toc_entries.append((pn + 1, f"Slide {pn + 1}"))

        body_parts.append('<section class="slide">')
        body_parts.extend(page_html)
        body_parts.append("</section>")

        print(f"  Page {pn + 1}/{n}\r", end="", flush=True)

    doc.close()
    print()

    if rendered_count > 0:
        print(f"  {rendered_count}/{n} pages rendered as images")

    # Build table of contents
    toc_html = '<nav><details open><summary>Table of Contents</summary><ol>'
    for pn, title in toc_entries:
        toc_html += f'<li><a href="#slide-{pn}">{esc(title)}</a></li>'
    toc_html += "</ol></details></nav>"

    return toc_html + "\n" + "\n".join(body_parts)


# ---------------------------------------------------------------------------
# PPTX extraction
# ---------------------------------------------------------------------------

def pptx_shape_sort_key(shape):
    """Sort shapes top-to-bottom, left-to-right."""
    return (shape.top or 0, shape.left or 0)


def convert_pptx(pptx_path):
    """Convert a PPTX file to HTML content."""
    if not HAS_PPTX:
        print("  ERROR: python-pptx not installed. Run: pip install python-pptx")
        return ""

    prs = Presentation(str(pptx_path))
    n = len(prs.slides)
    print(f"  {n} slides")

    toc_entries = []
    body_parts = []

    for sn, slide in enumerate(prs.slides, 1):
        parts = []
        slide_title = None

        shapes = sorted(slide.shapes, key=pptx_shape_sort_key)

        for shape in shapes:
            if shape.has_table:
                tbl = shape.table
                parts.append("<table>")
                for ri, row in enumerate(tbl.rows):
                    parts.append("<tr>")
                    for cell in row.cells:
                        tag = "th" if ri == 0 else "td"
                        parts.append(f"<{tag}>{esc(cell.text)}</{tag}>")
                    parts.append("</tr>")
                parts.append("</table>")
                continue

            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    blob = shape.image.blob
                    ext = shape.image.ext
                    b64 = base64.b64encode(blob).decode()
                    parts.append(
                        f'<img src="data:image/{ext};base64,{b64}" '
                        f'alt="Slide {sn} Image">'
                    )
                except Exception:
                    pass
                continue

            if not shape.has_text_frame or not shape.text.strip():
                continue

            is_title_shape = False
            try:
                if shape.placeholder_format is not None:
                    idx = shape.placeholder_format.idx
                    is_title_shape = idx in (0, 1)
            except Exception:
                pass

            tf = shape.text_frame
            in_list = False

            for para in tf.paragraphs:
                para_text = para.text.strip()
                if not para_text:
                    continue

                run_parts = []
                for run in para.runs:
                    t = esc(run.text)
                    if not t:
                        continue
                    try:
                        if run.font.bold:
                            t = f"<strong>{t}</strong>"
                        if run.font.italic:
                            t = f"<em>{t}</em>"
                    except Exception:
                        pass
                    run_parts.append(t)

                para_html = "".join(run_parts) if run_parts else esc(para_text)

                if is_title_shape and slide_title is None:
                    slide_title = para_text
                    parts.append(
                        f'<h2 class="slide-title" id="slide-{sn}">{para_html}</h2>'
                    )
                    is_title_shape = False
                elif para.level > 0:
                    if not in_list:
                        parts.append("<ul>")
                        in_list = True
                    parts.append(f"<li>{para_html}</li>")
                else:
                    if in_list:
                        parts.append("</ul>")
                        in_list = False
                    parts.append(f"<p>{para_html}</p>")

            if in_list:
                parts.append("</ul>")

        if slide_title:
            toc_entries.append((sn, slide_title))
        else:
            toc_entries.append((sn, f"Slide {sn}"))

        body_parts.append('<section class="slide">')
        body_parts.extend(parts)
        body_parts.append("</section>")

        print(f"  Slide {sn}/{n}\r", end="", flush=True)

    print()

    toc_html = '<nav><details open><summary>Table of Contents</summary><ol>'
    for sn, title in toc_entries:
        toc_html += f'<li><a href="#slide-{sn}">{esc(title)}</a></li>'
    toc_html += "</ol></details></nav>"

    return toc_html + "\n" + "\n".join(body_parts)


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def convert_file(path, render_mode=RENDER_AUTO):
    """Convert a single file and return (html_content, filename)."""
    p = Path(path)
    ext = p.suffix.lower()
    print(f"Converting: {p.name}")

    if ext == ".pdf":
        return convert_pdf(p, render_mode=render_mode), p.stem
    elif ext == ".pptx":
        return convert_pptx(p), p.stem
    else:
        print(f"  ERROR: Unsupported format '{ext}'. Use .pdf or .pptx")
        sys.exit(1)


def wrap_html(title, body):
    """Wrap content in full HTML document."""
    return (
        f"<!DOCTYPE html>\n<html lang=\"en\">\n<head>\n"
        f"<meta charset=\"UTF-8\">\n<title>{esc(title)}</title>\n"
        f"<style>\n{CSS}</style>\n</head>\n<body>\n"
        f"<h1>{esc(title)}</h1>\n{body}\n</body>\n</html>"
    )


def main():
    args = sys.argv[1:]

    if not args or args[0] in ("-h", "--help"):
        print("Usage: convert [options] <file1> [file2 ...] [-o output.html]")
        print()
        print("Options:")
        print("  --render      Render ALL pages as images (largest file)")
        print("  --no-render   Text extraction only, no page renders (smallest file)")
        print("  -o FILE       Output file (default: <input>.html)")
        print()
        print("Default: auto-renders pages with diagrams or equations (~40-60% of pages)")
        print()
        print("Examples:")
        print('  convert "SPCE5025 Class 3.pdf"')
        print('  convert lecture.pdf --render')
        print('  convert week1.pdf week2.pdf -o combined.html')
        sys.exit(0)

    # Parse arguments
    if "--render" in args:
        render_mode = RENDER_ALL
        args = [a for a in args if a != "--render"]
    elif "--no-render" in args:
        render_mode = RENDER_NONE
        args = [a for a in args if a != "--no-render"]
    else:
        render_mode = RENDER_AUTO

    output_path = None
    if "-o" in args:
        oi = args.index("-o")
        if oi + 1 < len(args):
            output_path = args[oi + 1]
            args = args[:oi] + args[oi + 2:]
        else:
            print("Error: -o requires a filename")
            sys.exit(1)

    input_files = [Path(a) for a in args]
    for f in input_files:
        if not f.exists():
            print(f"Error: File not found: {f}")
            sys.exit(1)

    if not input_files:
        print("Error: No input files specified")
        sys.exit(1)

    # Convert
    sections = []
    for f in input_files:
        content, name = convert_file(f, render_mode=render_mode)
        if len(input_files) > 1:
            content = f'<h1 id="file-{esc(name)}">{esc(name)}</h1>\n{content}'
        sections.append((content, name))

    # Determine output
    if output_path:
        out = Path(output_path)
    elif len(input_files) == 1:
        out = input_files[0].with_suffix(".html")
    else:
        out = Path("combined.html")

    # Build document
    if len(sections) == 1:
        title = sections[0][1]
        body = sections[0][0]
    else:
        title = "Combined: " + ", ".join(s[1] for s in sections)
        body = "\n<hr>\n".join(s[0] for s in sections)

    html = wrap_html(title, body)

    with open(out, "w", encoding="utf-8") as f:
        f.write(html)

    size_mb = out.stat().st_size / 1024 / 1024
    print(f"\nDone: {out} ({size_mb:.1f} MB)")


if __name__ == "__main__":
    main()
