import io
import ssl
import smtplib
import time

import pandas as pd
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.header import Header
from email.utils import formataddr


def make_message(from_addr, from_name, subject, to_addr, body_text):
    msg = MIMEMultipart("alternative")
    msg["From"] = formataddr((Header(from_name, "utf-8").encode(), from_addr))
    msg["To"] = to_addr
    msg["Cc"] = ""
    msg["Subject"] = Header(subject, "utf-8")
    msg.attach(MIMEText(body_text.replace(r"\n", "\n"), "plain", "utf-8"))
    return msg


def smtp_connect(host, port, from_addr, password):
    ctx = ssl.create_default_context()
    ctx.set_ciphers("DEFAULT:@SECLEVEL=0")
    s = smtplib.SMTP(host, int(port))
    s.starttls(context=ctx)
    s.login(from_addr, password)
    return s


# ── Page setup ────────────────────────────────────────────────────────────────

st.set_page_config(page_title="Mail Merge", page_icon="✉️", layout="wide")
st.title("✉️ Mail Merge")

# ── Session state defaults ────────────────────────────────────────────────────

for key in ("file_bytes", "sheet_names", "df"):
    if key not in st.session_state:
        st.session_state[key] = None

# ── Layout ────────────────────────────────────────────────────────────────────

col_cfg, col_send = st.columns([1, 1], gap="large")

# ── Left column: configuration form ──────────────────────────────────────────

with col_cfg:
    st.subheader("Configuration")
    with st.form("config"):
        smtp_host  = st.text_input("SMTP Host",  value="mail.grf.bg.ac.rs")
        smtp_port  = st.number_input("SMTP Port", value=587, step=1, min_value=1)
        from_addr  = st.text_input("Sender Email", value="nmilovanovic@grf.bg.ac.rs")
        from_name  = st.text_input("Sender Name",  value="Никола Миловановић")
        password   = st.text_input("Password", type="password")
        subject    = st.text_input("Subject", value="Путна инфраструктура, поставка задатка")
        bcc_raw    = st.text_input("BCC (comma-separated)", value="904_22@student.grf.bg.ac.rs")
        batch_size = st.number_input("Batch size", value=40, min_value=1)
        sleep_sec  = st.number_input("Sleep between batches (sec)", value=30, min_value=0)
        st.form_submit_button("Save config")

# ── Right column: file + preview + send ──────────────────────────────────────

with col_send:
    st.subheader("Data & Send")

    uploaded = st.file_uploader("Upload Excel file", type=["xlsx", "xls"])
    st.caption(
        "Required columns in the selected sheet: "
        "**`mail`** (recipient address) · **`msg`** (message body, use `\\n` for line breaks). "
        "The sheet is selected below — defaults to `input`."
    )

    if uploaded is not None:
        file_bytes = uploaded.read()
        if file_bytes != st.session_state.file_bytes:
            # New file uploaded — reset derived state
            st.session_state.file_bytes = file_bytes
            st.session_state.sheet_names = pd.ExcelFile(io.BytesIO(file_bytes)).sheet_names
            st.session_state.df = None

    if st.session_state.sheet_names:
        default_sheet = "input"
        default_idx = (
            st.session_state.sheet_names.index(default_sheet)
            if default_sheet in st.session_state.sheet_names
            else 0
        )
        sheet = st.selectbox("Sheet", st.session_state.sheet_names, index=default_idx)

        # Preview
        df_raw = pd.read_excel(io.BytesIO(st.session_state.file_bytes), sheet_name=sheet)

        if "mail" not in df_raw.columns or "msg" not in df_raw.columns:
            st.error(
                f"Sheet **{sheet}** is missing required columns. "
                f"Found: `{list(df_raw.columns)}`"
            )
            st.session_state.df = None
        else:
            df_valid = df_raw[
                df_raw["mail"].str.contains("@", na=False) & df_raw["msg"].notna()
            ].reset_index(drop=True)

            n_batches = max(1, -(-len(df_valid) // int(batch_size)))  # ceiling div
            st.success(
                f"**{len(df_valid)}** valid addresses — "
                f"**{n_batches}** batch(es) of ≤{int(batch_size)}"
            )

            preview = df_valid[["mail", "msg"]].head(5).copy()
            preview["msg"] = preview["msg"].str[:80] + "…"
            st.dataframe(preview, use_container_width=True)

            st.session_state.df = df_valid

    # ── Send ──────────────────────────────────────────────────────────────────

    st.divider()

    can_send = (
        st.session_state.df is not None
        and len(st.session_state.df) > 0
        and bool(password)
    )

    if not password:
        st.caption("Enter a password in the configuration form to enable sending.")

    if st.button("Send emails", type="primary", disabled=not can_send):
        df_valid = st.session_state.df
        bs = int(batch_size)
        batches = [df_valid.iloc[i : i + bs] for i in range(0, len(df_valid), bs)]
        total = len(df_valid)
        bcc_list = [b.strip() for b in bcc_raw.split(",") if b.strip()]

        progress_bar = st.progress(0.0)
        status_text  = st.empty()
        log_area     = st.container()

        sent_count = 0
        errors = []

        for batch_num, batch in enumerate(batches, 1):
            status_text.info(f"Batch {batch_num}/{len(batches)} — connecting to SMTP…")
            try:
                s = smtp_connect(smtp_host, smtp_port, from_addr, password)
            except Exception as exc:
                st.error(f"SMTP connection failed on batch {batch_num}: {exc}")
                break

            for _, row in batch.iterrows():
                to_addr = row["mail"].strip()
                try:
                    msg = make_message(from_addr, from_name, subject, to_addr, str(row["msg"]))
                    rcpt_list = list(set([to_addr] + bcc_list))
                    s.sendmail(from_addr, rcpt_list, msg.as_string())
                    sent_count += 1
                    log_area.write(f"  sent → {to_addr}")
                except Exception as exc:
                    errors.append((to_addr, str(exc)))
                    log_area.warning(f"  FAILED {to_addr}: {exc}")

                progress_bar.progress(
                    min((sent_count + len(errors)) / total, 1.0)
                )

            s.quit()

            if batch_num < len(batches):
                ss = int(sleep_sec)
                for remaining in range(ss, 0, -1):
                    status_text.info(
                        f"Sleeping {remaining}s before batch {batch_num + 1}/{len(batches)}…"
                    )
                    time.sleep(1)

        status_text.empty()
        progress_bar.progress(1.0)

        if errors:
            st.error(
                f"Finished with **{len(errors)}** error(s). "
                f"Sent: {sent_count}, Failed: {', '.join(e[0] for e in errors)}"
            )
        else:
            st.success(f"All **{sent_count}** emails sent successfully.")
