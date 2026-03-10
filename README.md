# NFPA 241 Plan Generator — CAP Design Group

Web app that generates a formatted NFPA 241 Construction Safety Plan as a .docx file.

## Deploy to Render.com (free)

### Step 1 — Push to GitHub
1. Create a new GitHub repository (e.g. `nfpa241-generator`)
2. Upload all files in this folder to the repo root:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `templates/index.html`

### Step 2 — Deploy on Render
1. Go to https://render.com and sign up (free)
2. Click **New → Web Service**
3. Connect your GitHub account → select the `nfpa241-generator` repo
4. Settings:
   - **Name:** nfpa241-generator
   - **Runtime:** Python 3
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** Free
5. Click **Create Web Service**
6. Wait ~3 minutes for first deploy
7. Your URL: `https://nfpa241-generator.onrender.com`

### Step 3 — Share with team
Bookmark the URL and share with Mooni, Fil, and Pon.

## Local testing
```
pip install flask python-docx
python app.py
# Open http://localhost:5000
```

## Updating the app
Edit files → push to GitHub → Render auto-deploys in ~2 minutes.
