# How to deploy this app online (free)

**Goal:** Get a URL like `audico-peachtree.streamlit.app` that accounting can open
in any browser to upload an EC xlsx and download Peachtree-ready CSVs.

**Time required:** About 15 minutes the first time. After that, updates take 30 seconds.

**Cost:** $0 forever (Streamlit Community Cloud free tier).

---

## Prerequisites (one-time)

You'll need two free accounts. If you already have them, skip ahead.

1. **GitHub account** — https://github.com/signup (Streamlit reads your code from GitHub)
2. **Streamlit account** — https://streamlit.io/cloud (sign in with the same GitHub account)

---

## Step 1 — Put the code on GitHub

### Option A: Using the GitHub website (easiest, no tools needed)

1. Go to https://github.com/new
2. **Repository name:** `audico-peachtree-converter` (or anything you like)
3. **Visibility:** Private (only you can see the code) — that's fine for Streamlit Cloud
4. Check ☑ **"Add a README file"**
5. Click **Create repository**
6. Now you need to upload the project files. On your new repo page:
   - Click **"Add file" → "Upload files"**
   - Drag the entire contents of the `streamlit_app` folder into the upload area:
     - `app.py`
     - `convert.py`
     - `requirements.txt`
     - `config/` folder (including `hotel_mapping.json` inside it)
   - Scroll down and click **"Commit changes"**

Your repo should now contain:
```
audico-peachtree-converter/
  app.py
  convert.py
  requirements.txt
  config/
    hotel_mapping.json
  README.md         (auto-created)
```

### Option B: Using git from your computer (if you know git)

```bash
cd streamlit_app
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/YOURUSERNAME/audico-peachtree-converter.git
git push -u origin main
```

---

## Step 2 — Deploy to Streamlit Community Cloud

1. Go to https://share.streamlit.io
2. Sign in with your GitHub account
3. Click **"New app"** (or "Create app" on some versions)
4. Fill in the form:
   - **Repository:** `YOURUSERNAME/audico-peachtree-converter`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL (custom subdomain):** `audico-peachtree` (or whatever — this becomes the public URL)
5. Click **Deploy**

Streamlit will now install the dependencies from `requirements.txt` and boot the app.
**First deploy takes 2–4 minutes.** After that, any updates redeploy in ~30 seconds.

When it's done, you'll get a URL like:
```
https://audico-peachtree.streamlit.app
```

That's the link to send to accounting. They bookmark it and use it every month.

---

## Step 3 — Test it

1. Open the URL in your browser
2. Upload the `EC_03_-_Marzo_2026.xlsx` file as a sanity test
3. Click **Convertir**
4. You should see **87 cotizaciones creadas**, download the ZIP, and spot-check
   one CSV before declaring victory

---

## How to update the app later

Any change you make to the code in GitHub → Streamlit automatically redeploys.

**Common updates:**

### Adding a new hotel mapping, fixing a customer ID, adding a skip row
1. On GitHub, open `config/hotel_mapping.json`
2. Click the pencil ✏️ icon (edit)
3. Make your change, scroll down, click **Commit changes**
4. Wait 30 seconds — the live app updates itself

### Fixing a bug in the converter
Same idea, but edit `convert.py`. No redeploy needed — GitHub push triggers it automatically.

---

## Things to know about Streamlit Community Cloud

### Pros
- Free forever for public and private apps
- Auto-redeploys on every GitHub commit
- Handles HTTPS, hosting, scaling, uptime for you
- Built-in error logs if something breaks

### Quirks to be aware of
- **Apps "sleep" after 7 days of zero use.** First visit after that takes ~30 seconds to wake up. After that it's instant. For a once-a-month tool, this means accounting may see a 30-second spinner the first time each month — then it's snappy for the rest of the session.
- **File uploads are in-memory only.** The xlsx you upload never gets saved on the Streamlit server — it lives in RAM during processing, then the temporary folder is deleted when the run finishes. This is actually good for privacy.
- **1 GB RAM limit.** Fine for EC xlsx files (a few MB).
- **Repository must be public OR you link your GitHub account.** Private repos work fine on the free tier as long as you authenticate Streamlit with GitHub.

### When to move off Streamlit Cloud
If one day accounting needs user logins, audit trails, or the app stops being free — migrations to paid hosts (Railway, Render, Fly.io) take ~15 minutes. The same code works.

---

## Troubleshooting

**"ModuleNotFoundError: No module named 'openpyxl'"** when deploying
→ `requirements.txt` wasn't uploaded or was committed to the wrong folder. It must sit next to `app.py` at the repo root.

**"File not found: config/hotel_mapping.json"** in the app
→ The `config/` folder wasn't uploaded to GitHub. Re-upload making sure the folder structure is preserved.

**App shows but conversion fails**
→ Click "Manage app" in the bottom-right of the Streamlit app, then "Logs". Copy the error and send it to me.

**Accounting complains about the 30-second wake-up delay**
→ Upgrade to Streamlit Cloud's paid tier ($20/mo) for always-on apps, OR
→ Migrate to Railway/Render's free tier which doesn't sleep (slightly more setup),
OR
→ Visit the URL yourself once a day to keep it warm (silly but works).

---

## A realistic note on security

This app lives on Streamlit's servers (operated by Snowflake). When accounting uploads an EC file:
- It goes over HTTPS (encrypted in transit)
- It sits in RAM on Streamlit's servers for ~5 seconds during processing
- It's deleted when the session ends
- Streamlit doesn't look at it, but they could if subpoenaed

For the kind of data in an EC file (client event names, amounts, hotel RUC numbers) this risk profile is generally fine for small businesses. If anyone at Audico is uncomfortable with this, the alternative is to self-host — put the same app on a PC at the Audico office that's always on. Same code, same UX, your network. Let me know if that's the better fit.
