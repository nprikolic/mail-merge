# Mail Merge

A mail merge tool for sending bulk personalized emails from an Excel spreadsheet. Available as a **web app** (recommended) or a **CLI script**.

## Web App (recommended)

### Setup

```bash
pip install streamlit pandas openpyxl
```

### Run

```bash
streamlit run app.py
```

Opens in the browser at `http://localhost:8501`.

### Usage

1. Fill in the configuration panel (SMTP host, sender email, password, subject, etc.)
2. Upload an Excel file (`.xlsx`)
3. Select the sheet containing your data
4. Click **Send emails**

#### Excel file format

The selected sheet must have these two columns:

| Column | Description |
|--------|-------------|
| `mail` | Recipient email address |
| `msg`  | Message body. Use `\n` (literal backslash-n) for line breaks. |

Rows with a missing or invalid `mail` value are skipped automatically.

### Notes

- The password is entered in the UI and never written to disk.
- The SMTP server (`mail.grf.bg.ac.rs`) only accepts connections from within the university network. Run the app on a machine connected to the university network or VPN.
- Emails are sent in batches (default: 40) with a sleep between batches (default: 30 s) to avoid rate limiting.

---

## CLI script (legacy)

### Setup

1. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

2. Create `password.txt` in this folder with your email password as the only content.

3. Open `mail_merge.py` and update `SUBJECT` and `SHEET` at the top of the file.

### Run

```bash
python mail_merge.py
```

> `password.txt` is git-ignored and must be created manually before running.
