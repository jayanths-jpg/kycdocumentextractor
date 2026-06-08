# Deploying to Streamlit Community Cloud

## Step 1 — Push to GitHub
1. Create a new GitHub repo (public or private)
2. Push this entire `kyc_annotator` folder as the root of the repo

## Step 2 — Deploy on Streamlit Cloud
1. Go to https://share.streamlit.io
2. Click **New app**
3. Connect your GitHub account and select the repo
4. Set **Main file path** to `app.py`
5. Click **Deploy**

## Step 3 — Add Secrets
1. In your deployed app, click **⋮ → Settings → Secrets**
2. Paste the contents of `secrets_template.toml` (filled in with real values):

```toml
OPENAI_API_KEY = "sk-your-real-key"

[users.alice]
password     = "alice123"
display_name = "Alice Smith"
role         = "admin"

[users.bob]
password     = "bobpass"
display_name = "Bob Raj"
role         = "user"
```

3. Click **Save** — the app restarts automatically.

## Adding / Removing Users
Just edit the Secrets in Streamlit Cloud settings and save.
No code changes or redeployment needed.

## Password Hashing (optional, more secure)
Generate a SHA-256 hash of a password in Python:
```python
import hashlib
print(hashlib.sha256("mypassword".encode()).hexdigest())
```
Paste the hash as the `password` value. The app accepts both plain text and hashes.

## App URL
After deployment your app is available at:
`https://<your-app-name>.streamlit.app`

Share this URL with your users along with their username and password.
