# 📋 Pandoc Command Reference

Quick reference for the DOCX to PDF conversion commands used in your app.

## 🎯 Primary Conversion (wkhtmltopdf)
```bash
pandoc input.docx -o output.pdf \
  --pdf-engine=wkhtmltopdf \
  --pdf-engine-opt=--enable-local-file-access \
  --pdf-engine-opt=--page-size A4 \
  --pdf-engine-opt=--margin-top 1in \
  --pdf-engine-opt=--margin-bottom 1in \
  --pdf-engine-opt=--margin-left 1in \
  --pdf-engine-opt=--margin-right 1in
```

## 🔄 Fallback Conversion (LaTeX)
```bash
pandoc input.docx -o output.pdf --pdf-engine=pdflatex
```

## 🧪 Test Commands

### Install Check:
```bash
pandoc --version
wkhtmltopdf --version
pdflatex --version
```

### Manual Test:
```bash
# Create simple test
echo "Hello {{Name}}!" | pandoc -o test.pdf --pdf-engine=wkhtmltopdf

# Test DOCX conversion
pandoc sample.docx -o sample.pdf --pdf-engine=wkhtmltopdf
```

## ⚙️ Engine Options

### wkhtmltopdf (Primary)
- **Best for**: Rich formatting, images, web-style layouts
- **Pros**: Excellent DOCX support, fast conversion
- **Cons**: Requires display server (handled by aptfile)

### pdflatex (Fallback)  
- **Best for**: Text-heavy documents, academic papers
- **Pros**: High-quality typography, reliable
- **Cons**: Limited rich formatting support

## 📦 Dependencies (aptfile)
```
pandoc
wkhtmltopdf  
texlive-latex-base
texlive-fonts-recommended
```

Your Flask app automatically uses these commands via `subprocess.run()` in the `convert_docx_to_pdf_pandoc()` method!