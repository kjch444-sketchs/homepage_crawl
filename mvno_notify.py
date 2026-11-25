"""가이드위반 시트를 메일 본문으로 보내고 위반 요금제를 텔레그램으로 알림."""

import argparse
import html
import json
import math
import os
import re
import smtplib
import urllib.parse
import urllib.request
from collections import Counter
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from numbers import Real
from pathlib import Path

from openpyxl import load_workbook


def guide_sections(report):
    wb = load_workbook(report, read_only=True, data_only=True)
    try:
        ws = wb["가이드위반"]
        sections = []
        current = None
        for cells in ws.iter_rows(min_col=2, max_col=11, values_only=True):
            number, carrier, plan, monthly, period, after = cells[:6]
            rule = cells[9]
            if isinstance(number, str) and "가이드 위반" in number:
                current = {"title": number, "rule": rule or "", "rows": []}
                sections.append(current)
            elif current is not None and isinstance(number, int) and carrier:
                current["rows"].append((number, carrier, plan, monthly, period, after))
        if len(sections) != 2:
            raise RuntimeError("가이드위반 시트에서 RS/RM 구역을 찾지 못했습니다")
        return sections
    finally:
        wb.close()


def _mail_table(title, headers, rows):
    def cell(value):
        if value is None:
            return "", "left"
        if isinstance(value, Real) and not isinstance(value, bool):
            if not math.isfinite(value):
                return "", "left"
            number = int(value) if float(value).is_integer() else value
            return html.escape(f"{number:,}"), "right"
        return html.escape(str(value)), "left"

    parts = [
        f'<h3 style="margin:24px 0 8px;color:#1F4E78;">{html.escape(title)}</h3>',
        '<table style="border-collapse:collapse;width:100%;font-family:Arial,Malgun Gothic,sans-serif;font-size:13px;">',
        '<tr>',
    ]
    for header in headers:
        parts.append(f'<th style="border:1px solid #cfd8e3;background:#D9EAF7;padding:7px;text-align:center;">{html.escape(header)}</th>')
    parts.append("</tr>")
    if not rows:
        parts.append(f'<tr><td colspan="{len(headers)}" style="border:1px solid #cfd8e3;padding:8px;text-align:center;color:#666;">해당 없음</td></tr>')
    for row in rows:
        parts.append("<tr>")
        for value in row:
            escaped, align = cell(value)
            parts.append(f'<td style="border:1px solid #cfd8e3;padding:7px;text-align:{align};vertical-align:top;">{escaped}</td>')
        parts.append("</tr>")
    return "".join(parts) + "</table>"


def _render_daily_report(source, rs_rows, rm_rows, lowest_rows, attachment_name, lowest_note=""):
    guide_headers = ("사업자", "요금제 명", "월 요금", "할인 기간", "기간 이후 요금")
    intro = "자동 크롤링 및 가이드 점검 결과입니다."
    safe_name = html.escape(attachment_name)
    return "".join([
        '<!doctype html><html><body style="font-family:Arial,Malgun Gothic,sans-serif;color:#222;line-height:1.5;">',
        '<div style="max-width:1400px;margin:0 auto;">',
        f'<h2 style="color:#17365D;margin-bottom:4px;">{html.escape(source)} 요금제 Daily Report</h2>',
        f'<p style="color:#666;margin-top:0;">{intro}</p>',
        _mail_table("RS 요금제 가이드 위반", guide_headers, rs_rows),
        _mail_table("RM 요금제 가이드 위반", guide_headers, rm_rows),
        _mail_table("요금제군별 최저가", ("요금제군", "사업자", "망정보", "월 요금(원)"), lowest_rows),
        f'<p style="color:#666;font-size:12px;">{html.escape(lowest_note)}</p>' if lowest_note else "",
        f'<p style="margin-top:28px;color:#666;">상세 내용은 첨부된 <b>{safe_name}</b> 파일을 확인해 주세요.</p>',
        '<p style="color:#888;font-size:12px;">이 메일은 자동으로 생성되었습니다.</p>',
        '</div></body></html>',
    ])


HIGH_DATA_CAPACITIES = (95, 125, 150, 180)
_HIGH_DATA_RE = re.compile(
    r"^\s*(?:월\s*)?(95|125|150|180)\s*(?:GB|G|기가바이트|기가)"
    r"(?![a-z0-9]|\s*(?:/\s*일|매일))",
    re.IGNORECASE,
)


def _high_data_capacity(value):
    """Match the base monthly allowance, not a daily/add-on amount or plan name."""
    match = _HIGH_DATA_RE.match(str(value))
    return int(match.group(1)) if match else None


def _high_data_lowest_rows(rows):
    """Append four capacity groups, retaining all tied carriers and networks."""
    best = {}
    for data, carrier, network, price in rows:
        capacity = _high_data_capacity(data)
        if (capacity is None or not isinstance(price, Real) or isinstance(price, bool)
                or not math.isfinite(price) or price < 0):
            continue
        identity = (str(carrier or "").strip(), str(network or "").strip())
        if not all(identity):
            continue
        previous_price, winners = best.get(capacity, (math.inf, set()))
        if price < previous_price:
            best[capacity] = price, {identity}
        elif price == previous_price:
            winners.add(identity)
    result = []
    for capacity in HIGH_DATA_CAPACITIES:
        label = f"{capacity}G"
        if capacity not in best:
            result.append([label, "해당 없음", "", None])
            continue
        price, winners = best[capacity]
        result.extend([label, carrier, network, price] for carrier, network in sorted(winners))
    return result


def lowest_price_rows(report):
    # The crawler's 참조 keys are the actual voice/text/data tuples in C:E.
    # Never read the Excel formula cache: openpyxl does not recalculate it.
    wb = load_workbook(report, read_only=True, data_only=True)
    try:
        categories = {}
        for _, _, voice, text, data, label, category in wb["참조"].iter_rows(
            min_row=3, max_col=7, values_only=True
        ):
            if data not in (None, ""):
                categories[(voice, text, data)] = (label, category)
        groups = {
            key: label.strip() if isinstance(label, str) and label.strip()
                 else f"{key[2]} / 통화 {key[0] or '-'} / 문자 {key[1] or '-'}"
            for key, (label, category) in categories.items()
            if str(category or "").strip() == "RS" and _high_data_capacity(key[2]) is None
        }
        # When a label is shared, keep the exact allowance groups distinguishable.
        counts = Counter(groups.values())
        for key, label in groups.items():
            if counts[label] > 1:
                groups[key] = f"{label} ({key[2]} / 통화 {key[0] or '-'} / 문자 {key[1] or '-'})"

        minimums = {}
        high_data_rows = []
        for carrier, _, voice, text, data, network, _, monthly in wb["홈페이지"].iter_rows(
            min_row=2, max_col=8, values_only=True
        ):
            key = (voice, text, data)
            if _high_data_capacity(data) is not None:
                high_data_rows.append((data, carrier, network, monthly))
            if (key not in groups or not carrier or not network
                    or not isinstance(monthly, Real) or isinstance(monthly, bool)
                    or not math.isfinite(monthly) or monthly < 0):
                continue
            price, winners = minimums.get(key, (math.inf, set()))
            identity = (str(carrier).strip(), str(network).strip())
            if monthly < price:
                minimums[key] = monthly, {identity}
            elif monthly == price:
                winners.add(identity)
        rows = []
        for key, name in groups.items():
            if key not in minimums:
                rows.append((name, "해당 없음", "", None))
                continue
            amount, winners = minimums[key]
            rows.extend((name, carrier, network, amount) for carrier, network in sorted(winners))
        rows.extend(_high_data_lowest_rows(high_data_rows))
        partial = "수집 오류" in wb.sheetnames and any(
            row[0] for row in wb["수집 오류"].iter_rows(min_row=2, max_col=1, values_only=True)
        )
        return rows, partial
    finally:
        wb.close()


def email_body(sections, report, attachment_name):
    guide_rows = {
        label: [row[1:] for section in sections if f" {label} " in section["title"]
                for row in section["rows"]]
        for label in ("RS", "RM")
    }
    lowest_rows, partial = lowest_price_rows(report)
    note = ("참조 시트에서 RS로 지정한 요금제군의 월 요금 기준입니다. "
            "95G·125G·150G·180G는 RS/RM과 관계없이 월 요금 기준입니다.")
    if partial:
        note += " 일부 사업자는 수집에 실패하여 성공한 데이터만 비교했습니다."
    return _render_daily_report(
        "홈페이지", guide_rows["RS"], guide_rows["RM"], lowest_rows, attachment_name, note
    )


def telegram_messages(sections, report):
    lines = ["홈페이지 가이드 위반"]
    for label in ("RS", "RM"):
        section = next((s for s in sections if f" {label} " in s["title"]), None)
        if section is None:
            raise RuntimeError(f"가이드위반 시트에서 {label} 구역을 찾지 못했습니다")
        lines.append(f"\n[{label}]")
        if not section["rows"]:
            lines.append("없음.")
        for _, carrier, plan, monthly, period, after in section["rows"]:
            lines.append(f"• {carrier} | {plan} | {monthly}원 / {period if period is not None else '-'}개월 / 이후 {after if after is not None else '-'}원")
    # 한 항목이 아주 길더라도 텔레그램의 메시지 제한을 넘지 않도록 분할.
    messages, chunk = [], ""
    for line in lines:
        for part in (line[i:i + 3000] for i in range(0, len(line), 3000)):
            if len(chunk) + len(part) + 1 > 3500:
                messages.append(chunk)
                chunk = ""
            chunk += ("\n" if chunk else "") + part
    if chunk:
        messages.append(chunk)
    return messages


def required(*names):
    missing = [name for name in names if not os.environ.get(name)]
    if missing:
        raise RuntimeError("GitHub Actions Secret 누락: " + ", ".join(missing))


def mail_subject():
    run_stamp = os.environ.get("RUN_STAMP", "")
    try:
        started_kst = datetime.strptime(run_stamp, "%Y%m%d_%H%M%S")
    except ValueError:
        started_kst = datetime.now(timezone(timedelta(hours=9)))
    return f"{started_kst:%Y%m%d}_{started_kst:%H}시{started_kst:%M}분_홈페이지_크롤링데이터"


def send_mail(report, sections):
    required("GMAIL_ADDRESS", "GMAIL_APP_PASSWORD", "MVNO_MAIL_TO")
    sender = os.environ["GMAIL_ADDRESS"]
    message = EmailMessage()
    message["From"] = sender
    message["To"] = os.environ["MVNO_MAIL_TO"]
    message["Subject"] = mail_subject()
    attachment_name = f"{message['Subject']}.xlsx"
    message.set_content("가이드 위반 및 요금제군별 최저가는 HTML 메일 본문과 첨부 엑셀을 확인하세요.")
    message.add_alternative(email_body(sections, report, attachment_name), subtype="html")
    message.add_attachment(report.read_bytes(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=attachment_name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=60) as smtp:
        smtp.login(sender, os.environ["GMAIL_APP_PASSWORD"].replace(" ", ""))
        smtp.send_message(message)
    print("메일 발송 완료")


def send_telegram(report, sections):
    required("MVNO_TELEGRAM_BOT_TOKEN", "MVNO_TELEGRAM_CHAT_ID")
    token = os.environ["MVNO_TELEGRAM_BOT_TOKEN"]
    for message in telegram_messages(sections, report):
        body = urllib.parse.urlencode({"chat_id": os.environ["MVNO_TELEGRAM_CHAT_ID"],
                                      "text": message}).encode("utf-8")
        request = urllib.request.Request(f"https://api.telegram.org/bot{token}/sendMessage",
                                         data=body, method="POST")
        try:
            with urllib.request.urlopen(request, timeout=60) as response:
                if not json.load(response).get("ok"):
                    raise RuntimeError("텔레그램 API가 요청을 거부했습니다")
        except Exception as exc:
            # HTTPError 등에 토큰이 포함된 URL이 찍히지 않도록 상세 오류는 숨긴다.
            raise RuntimeError("텔레그램 발송 실패: Bot token, Chat ID, 봇 대화 시작 여부를 확인하세요") from None
    print("텔레그램 발송 완료")


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("report", type=Path)
    args = parser.parse_args()
    sections = guide_sections(args.report)
    send_mail(args.report, sections)
    send_telegram(args.report, sections)
