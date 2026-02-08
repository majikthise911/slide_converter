  ---                                                                                     
  Built a tool to solve a problem that was driving me crazy.    
                                                                                          
  As an aerospace engineering grad student, my lecture slides are packed with orbital     
  mechanics equations, technical diagrams, and MATLAB code. I needed all of that in a
  single searchable document — but every existing converter either:

  - Couldn't install (dependency hell)
  - Lost all the diagrams
  - Turned equations into garbled text
  - Destroyed the document structure

  So I built slide_converter — a Python CLI that converts PDF/PPTX lecture slides to
  structured, single-file HTML with everything embedded.

  What it does:
  - Analyzes fonts to detect headings, bullets, equations, and code blocks
  - Auto-renders pages with diagrams or math as images (no more missing figures)
  - Embeds everything as base64 — one HTML file, no folders
  - Merges multiple files into one reference doc
  - Installs with one command: pip install
  git+https://github.com/majikthise911/slide_converter.git

  It's open source: https://github.com/majikthise911/slide_converter

  #Python #OpenSource #EdTech #AerospaceEngineering #GradSchool