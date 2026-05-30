"""SELLCO Email Bot — simple web app for your team."""

from __future__ import annotations

import io
import os
from pathlib import Path

import streamlit as st
from openpyxl import Workbook

from email_bot import (
    Recipient,
    SmtpConfig,
    load_recipients_upload,
    parse_sheet_arg,
    render_template,
    send_bulk,
)

st.set_page_config(
    page_title="SELLCO Email Bot",
    page_icon="📧",
    layout="centered",
)

DEFAULT_BODY = """Hi $name,

We wanted to reach out from SELLCO regarding $company.

If you have any questions, just reply to this email.

Best regards,
The SELLCO Team
"""


def get_app_password() -> str:
    try:
        return str(st.secrets.get("APP_PASSWORD", "")).strip()
    except Exception:
        return os.getenv("APP_PASSWORD", "").strip()


def get_smtp_config() -> SmtpConfig:
    try:
        smtp = st.secrets.get("smtp", {})
        if smtp:
            return SmtpConfig.from_mapping(dict(smtp))
    except Exception:
        pass
    return SmtpConfig.from_env()


def get_send_delay() -> float:
    try:
        return float(st.secrets.get("SEND_DELAY_SECONDS", os.getenv("SEND_DELAY_SECONDS", "1")))
    except (TypeError, ValueError):
        return 1.0


def require_login() -> bool:
    app_password = get_app_password()
    if not app_password:
        return True

    if st.session_state.get("authenticated"):
        return True

    st.title("SELLCO Email Bot")
    st.caption("Sign in with your company password to continue.")
    password = st.text_input("Password", type="password")
    if st.button("Sign in", type="primary", use_container_width=True):
        if password == app_password:
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def build_template_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipients"
    ws.append(["email", "name", "company"])
    ws.append(["alice@example.com", "Alice", "Acme Corp"])
    ws.append(["bob@example.com", "Bob", "Beta Inc"])
    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def recipients_to_table(recipients: list[Recipient]) -> list[dict[str, str]]:
    rows: list[dict[str, str]] = []
    for recipient in recipients:
        row = {"email": recipient.email, "name": recipient.name}
        if recipient.extra:
            row.update(recipient.extra)
        rows.append(row)
    return rows


def main() -> None:
    if not require_login():
        return

    st.title("SELLCO Email Bot")
    st.write("Upload an Excel file, write your message, and send to everyone at once.")

    with st.sidebar:
        st.header("Settings")
        sheet_input = st.text_input("Excel sheet (optional)", placeholder="First sheet")
        delay = st.number_input("Delay between emails (seconds)", min_value=0.0, value=get_send_delay(), step=0.5)
        st.download_button(
            label="Download Excel template",
            data=build_template_xlsx(),
            file_name="recipients_template.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.divider()
        st.caption("SMTP is configured by your admin. Team members only upload contacts and write the email.")

        if get_app_password():
            if st.button("Sign out", use_container_width=True):
                st.session_state.authenticated = False
                st.rerun()

    uploaded = st.file_uploader(
        "Upload recipient list",
        type=["xlsx", "xlsm", "csv", "json"],
        help="Excel file must have an `email` column in the first row.",
    )

    subject = st.text_input("Subject", placeholder="Hello from SELLCO")
    body = st.text_area("Email message", value=DEFAULT_BODY, height=220)

    with st.expander("Personalization help"):
        st.markdown(
            """
            Use placeholders in the subject and message:

            - `$name` — recipient name
            - `$email` — recipient email
            - Any extra Excel column, e.g. `$company`
            """
        )

    recipients: list[Recipient] = []
    if uploaded is not None:
        try:
            recipients = load_recipients_upload(
                uploaded.getvalue(),
                uploaded.name,
                sheet=parse_sheet_arg(sheet_input or None),
            )
            st.success(f"Loaded **{len(recipients)}** recipient(s) from `{uploaded.name}`")
            st.dataframe(recipients_to_table(recipients), use_container_width=True, hide_index=True)
        except Exception as exc:
            st.error(str(exc))

    col_preview, col_send = st.columns(2)

    with col_preview:
        preview_clicked = st.button("Preview", use_container_width=True, disabled=not recipients or not subject.strip())

    with col_send:
        send_clicked = st.button("Send emails", type="primary", use_container_width=True, disabled=not recipients or not subject.strip() or not body.strip())

    if preview_clicked and recipients:
        st.subheader("Preview")
        sample = recipients[0]
        st.markdown(f"**To:** {sample.email}")
        st.markdown(f"**Subject:** {render_template(subject.strip(), sample, {})}")
        st.text(render_template(body.strip(), sample, {}))

    if send_clicked and recipients:
        st.subheader("Sending...")
        progress = st.progress(0.0)
        status = st.empty()
        log_lines: list[str] = []

        def on_progress(current: int, total: int, email: str, ok: bool, error: str | None) -> None:
            progress.progress(current / total)
            if ok:
                log_lines.append(f"✅ {email}")
            else:
                log_lines.append(f"❌ {email}: {error}")
            status.markdown("\n".join(f"- {line}" for line in log_lines[-8:]))

        try:
            config = get_smtp_config()
        except ValueError as exc:
            st.error(f"Email server is not configured: {exc}")
            st.info("Ask your admin to add SMTP settings in Streamlit secrets or `.env`.")
            return

        success, failed = send_bulk(
            recipients,
            subject.strip(),
            body.strip(),
            config=config,
            delay_seconds=delay,
            on_progress=on_progress,
        )

        progress.progress(1.0)
        if failed == 0:
            st.success(f"Done! Sent {success} email(s).")
        else:
            st.warning(f"Finished. Sent: {success}, Failed: {failed}")


if __name__ == "__main__":
    main()
