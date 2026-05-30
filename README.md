# SELLCO Email Bot

Send one email campaign to many recipients in a single run. Each person gets a personalized message using placeholders like `$name` and `$email`.

**Two ways to use it:**
- **Web app** (recommended for your team) — simple browser UI, upload Excel, click Send
- **Command line** — for scripts and automation

---

## Web app for your team (free deploy)

### Run locally first

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Copy secrets for local testing:
   ```bash
   copy .streamlit\secrets.toml.example .streamlit\secrets.toml
   ```
   Edit `.streamlit\secrets.toml` with your SMTP details and a team password.

3. Start the web app:
   ```bash
   streamlit run streamlit_app.py
   ```
   Or double-click `run_web.bat`.

4. Open **http://localhost:8501** — upload Excel, write subject/body, send.

### Deploy free on Streamlit Cloud

Streamlit Community Cloud is **free** and gives your team a link like `https://sellco-email-bot.streamlit.app`.

1. **Push this folder to GitHub** (create a new public repo — code is public, passwords are not).

2. Go to [share.streamlit.io](https://share.streamlit.io) and sign in with GitHub.

3. Click **New app** → pick your repo → set **Main file path** to `streamlit_app.py`.

4. Open **App settings → Secrets** and paste:

   ```toml
   APP_PASSWORD = "your-team-password"

   [smtp]
   host = "smtp.gmail.com"
   port = 587
   use_tls = true
   username = "your-email@gmail.com"
   password = "your-app-password"
   from_email = "your-email@gmail.com"
   from_name = "SELLCO"

   SEND_DELAY_SECONDS = 1
   ```

5. Click **Deploy**. Share the URL + `APP_PASSWORD` with your company.

**How it works for your team:**
- Admin configures SMTP once (in Secrets — never visible to users)
- Everyone uses the same link + password
- Each person uploads their Excel list and sends their own campaign
- No Python or command line needed

**Free tier limits:** Public GitHub repo required; app sleeps after inactivity (wakes on next visit); fine for small/medium company use.

---

## Command line (optional)
1. **Install Python 3.10+** if you don't have it.

2. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

3. **Configure SMTP** — copy the example env file and fill in your details:

   ```bash
   copy .env.example .env
   ```

   | Provider | SMTP_HOST | Notes |
   |----------|-----------|-------|
   | Gmail | `smtp.gmail.com` | Use an [App Password](https://myaccount.google.com/apppasswords), not your regular password |
   | Outlook / Microsoft 365 | `smtp.office365.com` | Port 587, TLS enabled |
   | Custom | Your provider's SMTP host | Check their docs for port and TLS settings |

4. **Edit your recipient list in Excel** — open `recipients.xlsx` and fill in your users:

   | email | name | company |
   |-------|------|---------|
   | alice@example.com | Alice | Acme Corp |
   | bob@example.com | Bob | Beta Inc |

   Required column: **email**. Optional: **name**, plus any extra columns for personalization.

5. **Edit the email content** in `email_body.txt` (and optionally set the subject on the command line).

6. **Preview without sending:**

   ```bash
   python email_bot.py --dry-run --subject "Hello from SELLCO"
   ```

7. **Send to everyone:**

   ```bash
   python email_bot.py --subject "Hello from SELLCO"
   ```

   To use your own Excel file:

   ```bash
   python email_bot.py -r "C:\path\to\users.xlsx" --subject "Hello from SELLCO"
   ```

## Excel format

- Save as `.xlsx` (standard Excel format)
- First row must be column headers
- Must include an `email` column (case doesn't matter — `Email` works too)
- Add any extra columns you want to use in the email body, e.g. `company`, `product`
- Empty rows are skipped automatically

If your workbook has multiple sheets, pick one with `--sheet`:

```bash
python email_bot.py --sheet "Contacts" --subject "Hello"
python email_bot.py --sheet 1 --subject "Hello"
```

## Personalization

Use `$placeholder` syntax in the subject and body. Built-in placeholders:

- `$email` — recipient email
- `$name` — name from Excel (falls back to the part before `@`)

Any extra Excel columns become placeholders too. With a `company` column, use `$company` in your template.

## Command-line options

```bash
python email_bot.py --help
```

| Option | Description |
|--------|-------------|
| `-r`, `--recipients` | Excel, CSV, or JSON file (default: `recipients.xlsx`) |
| `--sheet` | Excel sheet name or index (default: first sheet) |
| `-s`, `--subject` | Email subject line |
| `--subject-file` | Read subject from a file |
| `-b`, `--body` | Plain-text body inline |
| `--body-file` | Body file (default: `email_body.txt`) |
| `--html-file` | Optional HTML version of the email |
| `--delay` | Seconds between sends (default: 1) |
| `--dry-run` | Preview only, no emails sent |

## CSV or JSON (optional)

You can still use CSV or JSON instead of Excel:

```bash
python email_bot.py -r recipients.csv --subject "Hello"
python email_bot.py -r recipients.json --subject "Hello"
```

## Tips

- Start with `--dry-run` to verify placeholders and recipient list.
- Keep `--delay` at 1 second or higher for large lists to reduce spam-filter risk.
- Never commit `.env` — it contains your SMTP password.
