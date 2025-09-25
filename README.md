# Mail Merge SaaS - Complete Deployment Guide

## 🚀 Deploy to Render (Recommended)

### 1. Prepare Your Files
All files are ready! Your project structure should be:
```
site1/
├── app.py                 # Main Flask application
├── mailmerge.html         # Mail merge page
├── index.html            # Home page
├── style.css             # Styles
├── mailmerge.js          # Frontend JavaScript
├── requirements.txt      # Python dependencies
├── render.yaml           # Render config
├── gunicorn.conf.py      # Production server config
└── README.md             # This file
```

### 2. Deploy Steps

#### Option A: GitHub + Render (Recommended)
1. **Push to GitHub:**
   ```bash
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin https://github.com/yourusername/mail-merge-saas.git
   git push -u origin main
   ```

2. **Deploy on Render:**
   - Go to [render.com](https://render.com)
   - Click "New" → "Web Service"
   - Connect your GitHub repository
   - Configure:
     - **Name**: mail-merge-saas
     - **Environment**: Python 3
     - **Build Command**: `pip install -r requirements.txt`
     - **Start Command**: `gunicorn app:app`
   - Click "Deploy"

#### Option B: Direct Upload
1. **Zip your files**
2. **Upload to Render** using their dashboard
3. **Same configuration as above**

### 3. Configuration
Render will automatically:
- ✅ Install Python dependencies from `requirements.txt`
- ✅ Start your app with Gunicorn
- ✅ Provide HTTPS
- ✅ Give you a URL like: `https://your-app.onrender.com`

### 4. Test Your Deployment
1. Visit your Render URL
2. Test the mail merge functionality:
   - Upload a .docx template with fields like `{{name}}`
   - Upload an .xlsx file with matching columns
   - Select output format
   - Process and download

## 📝 How to Use Your SaaS

### For Your Users:
1. **Create Template:** Word document with `{{field_name}}` placeholders
2. **Prepare Data:** Excel file with column headers matching field names
3. **Upload & Process:** Use your website to merge and download

### Example Files:

**Template.docx content:**
```
Dear {{first_name}} {{last_name}},

Welcome to {{company}}! 
Your account email is: {{email}}

Best regards,
The Team
```

**Data.xlsx content:**
| first_name | last_name | company | email |
|------------|-----------|---------|-------|
| John | Doe | ABC Corp | john@abc.com |
| Jane | Smith | XYZ Ltd | jane@xyz.com |

## 🎯 Features
- ✅ Drag & drop file uploads
- ✅ Real-time data preview
- ✅ Multiple output formats (PDF/Word, single/multiple)
- ✅ Professional UI matching iLovePDF style
- ✅ Mobile responsive design
- ✅ Error handling and validation
- ✅ Automatic file cleanup

## 💰 Render Pricing
- **Free Tier**: 750 hours/month (perfect for testing)
- **Paid Tier**: $7/month for unlimited usage
- **Custom Domain**: Free on all tiers

## 🔒 Security & Privacy
- Files are temporarily stored during processing
- Automatic cleanup after download
- HTTPS encryption
- No permanent data storage

## 📊 Monitoring
- Built-in health check endpoint: `/health`
- Render provides logs and monitoring
- Automatic restarts if needed

## 🚀 Going Live
1. **Custom Domain:** Point your domain to Render
2. **Analytics:** Add Google Analytics to track usage
3. **Marketing:** Share your mail merge SaaS with users
4. **Monetization:** Add payment integration if desired

## 🆘 Troubleshooting

### Common Issues:
1. **PDF conversion fails**: LibreOffice will be installed automatically
2. **Large files timeout**: Files are limited to 50MB
3. **Build failures**: Check Python version compatibility

### Support:
- Check Render logs for detailed error messages
- Test locally first: `python app.py`
- Verify file formats and structure

Your Mail Merge SaaS is now ready for production! 🎉