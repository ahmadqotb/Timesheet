# Employee Timesheet Generator - Web App

## Deploy to Railway

### Step 1: Push to GitHub

1. Create a new repository on GitHub
2. Open terminal in this folder
3. Run these commands:

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPO.git
git push -u origin main
```

### Step 2: Deploy on Railway

1. Go to https://railway.app/new
2. Click **"GitHub Repository"**
3. Select your repository
4. Railway will auto-detect and deploy

### Step 3: Get Your URL

1. After deployment, click on your service
2. Go to **Settings** → **Networking**
3. Click **"Generate Domain"**
4. Your app will be live at: `https://your-app.railway.app`

---

## Local Development

```bash
npm install
npm start
```

Open: http://localhost:3000

---

## Files Overview

| File | Purpose |
|------|---------|
| server.js | Express web server |
| generator.js | PDF generation |
| index.html | Web interface |
| absenceReport.js | Absence reports |
| dataValidator.js | Data validation |
| foodAllowance.js | Food allowance |
| projectSummary.js | Project summary |
| nixpacks.toml | Railway Puppeteer config |

---

## Notes

- Uses Puppeteer for PDF generation (requires Chromium)
- `nixpacks.toml` configures Chromium for Railway
- Uploads/outputs are temporary (cleared on redeploy)
