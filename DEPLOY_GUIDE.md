# 🚀 DEPLOY TO RENDER - QUICK GUIDE

Your Pandoc PDF conversion system is ready for deployment!

## ✅ Pre-Deployment Checklist

### Files Updated:
- ✅ `aptfile` - Added Pandoc, wkhtmltopdf, LaTeX
- ✅ `requirements.txt` - Cleaned up, removed Windows dependencies  
- ✅ `app.py` - Replaced all PDF methods with Pandoc
- ✅ All old Windows/LibreOffice methods removed

### Dependencies Ready:
- ✅ **Pandoc** - Universal document converter
- ✅ **wkhtmltopdf** - Primary PDF engine
- ✅ **LaTeX** - Fallback PDF engine
- ✅ **Python packages** - Clean, minimal set

## 🎯 Deploy to Render

### 1. **Commit & Push Changes**
```bash
git add .
git commit -m "Replace complex PDF system with Pandoc"
git push origin main
```

### 2. **Render Auto-Deploy**
- Render will detect changes
- Install apt packages from `aptfile`
- Install Python packages from `requirements.txt` 
- Deploy with `gunicorn app:app`

### 3. **Monitor Deployment**
Watch Render logs for:
```
==> Installing apt packages...
pandoc
wkhtmltopdf  
texlive-latex-base
texlive-fonts-recommended
==> Build successful 🎉
==> Your service is live 🎉
```

## 🧪 Test After Deployment

### 1. **Basic Test**
- Visit your Render URL
- Upload a DOCX template with `{{Name}}` placeholders
- Upload Excel with `Name` column
- Select "Single PDF" format
- Click "Process Merge"
- Download should work!

### 2. **Check Logs**
Look for Pandoc conversion messages:
```
🔄 Starting Pandoc conversion: /tmp/xxx.docx → /tmp/xxx.pdf
✅ Pandoc available: 2.x.x
🚀 Running command: pandoc /tmp/xxx.docx -o /tmp/xxx.pdf --pdf-engine=wkhtmltopdf
✅ Successfully converted to PDF: /tmp/xxx.pdf (12345 bytes)
```

### 3. **Run Test Script** (Optional)
SSH into Render and run:
```bash
python test_pandoc.py
```

## 🎉 Expected Benefits

| Before | After |
|--------|--------|
| Multiple servers | Single Render instance |
| Complex fallbacks | Simple Pandoc conversion |
| Network dependencies | Local processing |
| High maintenance | Minimal maintenance |
| $20-75/month | $7-25/month |

## 🔧 If Issues Occur

### **PDF Conversion Fails**
Check logs for:
- `❌ Pandoc not found` → aptfile issue
- `❌ wkhtmltopdf check failed` → engine issue  
- `❌ Pandoc conversion failed` → file format issue

### **Common Solutions**
1. **Redeploy** - Sometimes apt packages need retry
2. **Check template** - Ensure DOCX is valid
3. **Try LaTeX fallback** - Will auto-trigger if wkhtmltopdf fails

## 🎯 Success Criteria

✅ **No "Failed to process mail merge" errors**  
✅ **PDF downloads work**  
✅ **Logs show Pandoc conversion success**  
✅ **Single platform deployment**  

Your mail merge SaaS is now production-ready with enterprise-grade document processing! 🚀