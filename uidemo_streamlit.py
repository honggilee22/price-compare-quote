# -*- coding: utf-8 -*-
from __future__ import annotations

import base64
from datetime import date, datetime
from email.message import EmailMessage
from io import BytesIO
import hashlib
import json
import os
from pathlib import Path
import platform
import re
import shutil
import smtplib
import subprocess
import tempfile
import time

import openpyxl
import streamlit as st
import streamlit.components.v1 as components

HTML_PATH = Path(__file__).with_name("uidemo.html")
TEMPLATE_PATH = Path(__file__).with_name("SS.xlsx")
COMPONENT_DIR = Path(__file__).with_name("components") / "price_compare"
CONFIG_PATH = Path(__file__).with_name("output") / "app_settings.json"
DEFAULT_SAVE_DIR = Path(__file__).with_name("output") / "quotes"

HARDCODE_GMAIL_USER = "jumsune2@gmail.com"
HARDCODE_GMAIL_APP_PASSWORD = ""

MAX_TEMPLATE_ROWS = 9
COMPONENT_HEIGHT = 1200
CONTENT_WIDTH = 2000
CONTENT_HEIGHT = 1200

BRIDGE_SCRIPT = """<script id=\"price-compare-bridge\">
(function() {
  function buildPayload(action) {
    const recipient = (document.getElementById("quote-recipient")?.value || "").trim();
    const ext = (document.getElementById("quote-ext")?.value || "").trim();
    const quoteDate = document.getElementById("quote-date")?.value || "";
    const email = (document.getElementById("quote-email")?.value || "").trim();

    const plan1El = document.querySelector('[data-plan="1"]');
    const plan2El = document.querySelector('[data-plan="2"]');

    const plan1 = collectPlanData(plan1El, "1안");
    const plan2 = collectPlanData(plan2El, "2안");

    const discount1 = getDiscount(plan1El) || 0;
    const discount2 = getDiscount(plan2El) || 0;

    return {
      request_id: String(Date.now()),
      action: action,
      data: {
        recipient: recipient,
        ext: ext,
        quote_date: quoteDate,
        email: email,
        plan1: plan1,
        plan2: plan2,
        discount1: discount1,
        discount2: discount2
      }
    };
  }

  function post(action) {
    const payload = buildPayload(action);
    window.parent.postMessage({ source: "price-compare", payload: payload }, "*");
  }

  function ensurePdfButton() {
    const summary = document.getElementById("summary");
    const csvBtn = document.getElementById("export-csv");
    if (!summary || !csvBtn) return;
    if (document.getElementById("export-pdf")) return;

    const pdfBtn = csvBtn.cloneNode(true);
    pdfBtn.id = "export-pdf";
    pdfBtn.textContent = "PDF 저장";
    pdfBtn.style.right = "90px";
    summary.appendChild(pdfBtn);

    pdfBtn.addEventListener("click", () => {
      updateAll();
      post("download_pdf");
    });
  }

  function enableSendButton() {
    const buttons = document.querySelectorAll(".menu button");
    const sendBtn = Array.from(buttons).find(
      (btn) => btn.textContent.replace(/\\s+/g, "") === "견적서발송"
    );
    if (!sendBtn) return;

    sendBtn.disabled = false;
    sendBtn.addEventListener("click", () => {
      updateAll();
      post("send_email");
    });
  }

  function init() {
    ensurePdfButton();
    enableSendButton();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else {
    init();
  }
})();
</script>
"""


def load_settings():
    if not CONFIG_PATH.exists():
        return {}
    try:
        return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def persist_save_dir(raw_value):
    save_path = Path(raw_value).expanduser()
    if not save_path.is_absolute():
        save_path = Path(__file__).parent / save_path
    save_path.mkdir(parents=True, exist_ok=True)
    CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
    payload = {"save_dir": str(save_path)}
    CONFIG_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    return payload["save_dir"]


def inject_bridge(html_text):
    if "price-compare-bridge" in html_text:
        return html_text
    if "</body>" in html_text:
        return html_text.replace("</body>", f"{BRIDGE_SCRIPT}\n</body>")
    return f"{html_text}\n{BRIDGE_SCRIPT}"


def sanitize_filename(text):
    cleaned = re.sub(r"[\\/:*?\"<>|]", "_", text or "")
    cleaned = re.sub(r"\s+", "", cleaned)
    return cleaned or "견적서"


def format_date_label(quote_date):
    if not quote_date:
        return ""
    return quote_date.strftime("%Y.%m.%d")


def format_file_date(quote_date):
    if not quote_date:
        return date.today().strftime("%m월%d일")
    return quote_date.strftime("%m월%d일")


def parse_date(value):
    if not value:
        return date.today()
    try:
        return datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return date.today()


def parse_number(value):
    if value is None:
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if not text:
        return 0
    text = re.sub(r"[^\d]", "", text)
    return int(text or 0)


def parse_float(value):
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def normalize_rows(rows):
    normalized = []
    for row in rows:
        model = str(row.get("model", "") or "").strip()
        price = parse_number(row.get("price"))
        qty = parse_number(row.get("qty"))
        if model or price or qty:
            normalized.append({"model": model, "price": price, "qty": qty})
    if not normalized:
        normalized = [{"model": "", "price": 0, "qty": 0}]
    while len(normalized) < MAX_TEMPLATE_ROWS:
        normalized.append({"model": "", "price": 0, "qty": 0})
    return normalized[:MAX_TEMPLATE_ROWS]


def compute_totals(rows):
    total_sum = 0
    for row in rows:
        price = parse_number(row.get("price"))
        qty = parse_number(row.get("qty"))
        total_sum += price * qty
    return total_sum


def format_won(value):
    try:
        amount = int(round(float(value)))
    except (TypeError, ValueError):
        amount = 0
    return f"{amount:,}원"


def build_summary_text(total1, total2):
    if total2 < total1:
        diff = total1 - total2
        return f"연 {format_won(diff * 12)} 할인 / 월 {format_won(diff)} 할인"
    if total2 > total1:
        diff = total2 - total1
        return f"월 {format_won(diff)} 추가"
    return "차이 0원"


def load_template():
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"템플릿을 찾을 수 없습니다: {TEMPLATE_PATH}")
    return openpyxl.load_workbook(TEMPLATE_PATH)


def fill_template(
    recipient,
    ext,
    quote_date,
    email,
    plan1_rows,
    plan2_rows,
    plan1_total,
    plan2_total,
    plan1_prepay,
    plan2_prepay,
    summary_text="",
):
    wb = load_template()
    ws = wb.active

    ws["B4"] = recipient or ""
    ws["B13"] = summary_text or ""
    ws["D5"] = format_date_label(quote_date)
    ws["D6"] = ext or ""
    ws["D7"] = email or ""

    ws["E11"] = plan1_total
    ws["N11"] = plan2_total
    ws["E12"] = plan1_prepay
    ws["N12"] = plan2_prepay

    for index, row in enumerate(plan1_rows):
        excel_row = 15 + index
        model = row.get("model", "")
        price = parse_number(row.get("price"))
        qty = parse_number(row.get("qty"))
        total = price * qty
        ws[f"B{excel_row}"] = model
        ws[f"E{excel_row}"] = price if price else ""
        ws[f"G{excel_row}"] = qty if qty else ""
        ws[f"H{excel_row}"] = total if total else ""

    for index, row in enumerate(plan2_rows):
        excel_row = 15 + index
        model = row.get("model", "")
        price = parse_number(row.get("price"))
        qty = parse_number(row.get("qty"))
        total = price * qty
        ws[f"K{excel_row}"] = model
        ws[f"N{excel_row}"] = price if price else ""
        ws[f"P{excel_row}"] = qty if qty else ""
        ws[f"Q{excel_row}"] = total if total else ""

    return wb


def build_excel_bytes(**kwargs):
    wb = fill_template(**kwargs)
    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def convert_excel_to_pdf(xlsx_bytes):
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_dir = Path(tmp_dir)
        xlsx_path = tmp_dir / "quote.xlsx"
        pdf_path = tmp_dir / "quote.pdf"
        xlsx_path.write_bytes(xlsx_bytes)

        errors = []
        if platform.system().lower().startswith("win"):
            try:
                import win32com.client  # type: ignore

                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                excel.DisplayAlerts = False
                workbook = excel.Workbooks.Open(str(xlsx_path))
                workbook.ExportAsFixedFormat(0, str(pdf_path))
                workbook.Close(False)
                excel.Quit()
                return pdf_path.read_bytes()
            except Exception as exc:
                errors.append(f"Excel 변환 실패: {exc}")

        soffice = shutil.which("soffice")
        if soffice:
            cmd = [
                soffice,
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                str(tmp_dir),
                str(xlsx_path),
            ]
            result = subprocess.run(cmd, capture_output=True, text=True, check=False)
            if result.returncode == 0 and pdf_path.exists():
                return pdf_path.read_bytes()
            errors.append(result.stderr or result.stdout or "LibreOffice 변환 실패")

        if not errors:
            errors.append("PDF 변환 도구를 찾을 수 없습니다.")
        raise RuntimeError(" / ".join(errors))


def resolve_gmail_credentials():
    if HARDCODE_GMAIL_USER and HARDCODE_GMAIL_APP_PASSWORD:
        return HARDCODE_GMAIL_USER, HARDCODE_GMAIL_APP_PASSWORD
    secrets = getattr(st, "secrets", {})
    user = secrets.get("gmail_user") or os.getenv("GMAIL_USER")
    password = secrets.get("gmail_app_password") or os.getenv("GMAIL_APP_PASSWORD")
    return user, password


def send_email(to_email, subject, body, pdf_bytes, pdf_name):
    smtp_user, smtp_password = resolve_gmail_credentials()
    if not smtp_user or not smtp_password:
        raise RuntimeError(
            "Gmail 계정 정보가 없습니다. st.secrets 또는 환경변수에 설정하세요."
        )

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = smtp_user
    msg["To"] = to_email
    msg.set_content(body)
    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg)


def build_filename(recipient, quote_date):
    recipient_label = recipient.strip() if recipient else ""
    base_label = f"{recipient_label}귀하" if recipient_label else "견적서"
    date_label = format_file_date(quote_date)
    return sanitize_filename(f"{base_label}{date_label}")


def save_artifacts(save_dir, file_stem, xlsx_bytes, pdf_bytes=None):
    save_path = Path(save_dir)
    save_path.mkdir(parents=True, exist_ok=True)
    xlsx_path = save_path / f"{file_stem}.xlsx"
    xlsx_path.write_bytes(xlsx_bytes)
    if pdf_bytes:
        pdf_path = save_path / f"{file_stem}.pdf"
        pdf_path.write_bytes(pdf_bytes)


def handle_event(event, save_dir):
    action = event.get("action")
    request_id = event.get("request_id") or str(int(time.time() * 1000))
    data = event.get("data") or {}

    recipient_raw = str(data.get("recipient") or "").strip()
    recipient = f"{recipient_raw} 귀하" if recipient_raw else ""
    ext = str(data.get("ext") or "").strip()
    email = str(data.get("email") or "").strip()
    quote_date = parse_date(data.get("quote_date"))

    plan1_rows = (data.get("plan1") or {}).get("rows") or []
    plan2_rows = (data.get("plan2") or {}).get("rows") or []
    plan1_rows_norm = normalize_rows(plan1_rows)
    plan2_rows_norm = normalize_rows(plan2_rows)

    plan1_total = compute_totals(plan1_rows_norm)
    plan2_total = compute_totals(plan2_rows_norm)

    discount1 = max(parse_float(data.get("discount1")), 0.0)
    discount2 = max(parse_float(data.get("discount2")), 0.0)
    plan1_prepay = int(round(plan1_total * (1 - discount1)))
    plan2_prepay = int(round(plan2_total * (1 - discount2)))
    summary_text = build_summary_text(plan1_prepay, plan2_prepay)

    xlsx_bytes = build_excel_bytes(
        recipient=recipient,
        ext=ext,
        quote_date=quote_date,
        email=email,
        plan1_rows=plan1_rows_norm,
        plan2_rows=plan2_rows_norm,
        plan1_total=plan1_total,
        plan2_total=plan2_total,
        plan1_prepay=plan1_prepay,
        plan2_prepay=plan2_prepay,
        summary_text=summary_text,
    )

    file_stem = build_filename(recipient_raw, quote_date)

    try:
        pdf_bytes = convert_excel_to_pdf(xlsx_bytes)
    except Exception as exc:
        return {
            "id": request_id,
            "type": "message",
            "message": f"PDF 생성 실패: {exc}",
        }

    if action == "send_email":
        if not email:
            return {
                "id": request_id,
                "type": "message",
                "message": "이메일 주소를 입력하세요.",
            }
        try:
            subject = f"비교 견적서 ({recipient_raw or '수신자'})"
            body = "첨부된 PDF 견적서를 확인해 주세요."
            send_email(
                to_email=email,
                subject=subject,
                body=body,
                pdf_bytes=pdf_bytes,
                pdf_name=f"{file_stem}.pdf",
            )
            save_artifacts(save_dir, file_stem, xlsx_bytes=xlsx_bytes, pdf_bytes=pdf_bytes)
            return {
                "id": request_id,
                "type": "pdf",
                "filename": f"{file_stem}.pdf",
                "content": base64.b64encode(pdf_bytes).decode("ascii"),
                "message": "이메일 전송 완료",
            }
        except Exception as exc:
            return {
                "id": request_id,
                "type": "message",
                "message": f"전송 실패: {exc}",
            }

    if action == "download_pdf":
        return {
            "id": request_id,
            "type": "pdf",
            "filename": f"{file_stem}.pdf",
            "content": base64.b64encode(pdf_bytes).decode("ascii"),
        }

    return {
        "id": request_id,
        "type": "message",
        "message": "알 수 없는 요청입니다.",
    }


st.set_page_config(
    page_title="가격 비교",
    layout="wide",
    initial_sidebar_state="collapsed",
)

DEFAULT_SAVE_DIR.mkdir(parents=True, exist_ok=True)
settings = load_settings()
if "save_dir" not in st.session_state:
    st.session_state["save_dir"] = settings.get("save_dir", str(DEFAULT_SAVE_DIR))

with st.sidebar:
    st.markdown("저장 설정")
    save_dir_input = st.text_input(
        "저장 폴더",
        value=st.session_state["save_dir"],
    )
    if st.button("저장 위치 저장"):
        try:
            saved_dir = persist_save_dir(save_dir_input)
            st.session_state["save_dir"] = saved_dir
            st.success("저장 완료")
        except Exception as exc:
            st.error(f"저장 실패: {exc}")
    st.caption(f"현재 위치: {st.session_state['save_dir']}")
    st.divider()
    st.caption("Gmail 앱 비밀번호는 st.secrets 또는 환경변수로 설정하세요.")

st.markdown(
    """
    <style>
      #MainMenu {visibility: hidden;}
      footer {visibility: hidden;}
      header {visibility: hidden;}
      .block-container {padding: 0;}
    </style>
    """,
    unsafe_allow_html=True,
)

if not HTML_PATH.exists():
    st.error(f"HTML 파일을 찾을 수 없습니다: {HTML_PATH}")
    st.stop()

if not COMPONENT_DIR.exists():
    st.error(f"컴포넌트 폴더를 찾을 수 없습니다: {COMPONENT_DIR}")
    st.stop()

html = HTML_PATH.read_text(encoding="utf-8")
html = inject_bridge(html)
html_hash = hashlib.sha256(html.encode("utf-8")).hexdigest()

price_compare_component = components.declare_component(
    "price_compare",
    path=str(COMPONENT_DIR),
)

response = st.session_state.get("bridge_response")
last_request_id = st.session_state.get("last_request_id")

payload = price_compare_component(
    html=html,
    html_hash=html_hash,
    response=response,
    content_width=CONTENT_WIDTH,
    content_height=CONTENT_HEIGHT,
    height=COMPONENT_HEIGHT,
)

if payload and isinstance(payload, dict):
    request_id = payload.get("request_id")
    if request_id and request_id != last_request_id:
        st.session_state["last_request_id"] = request_id
        st.session_state["bridge_response"] = handle_event(
            payload,
            st.session_state["save_dir"],
        )
        st.experimental_rerun()
