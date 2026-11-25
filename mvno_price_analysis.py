"""홈페이지 보고서: 직전 회차 최저가 비교와 RS/RM 경쟁력 대시보드."""

import json
import math
import re
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from fractions import Fraction
from pathlib import Path
from tempfile import NamedTemporaryFile
from zipfile import BadZipFile

from openpyxl import load_workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.utils.exceptions import InvalidFileException

DASHBOARD = "경쟁력 대시보드"
MAIL_BLOCK = "직전 회차 대비 최저가 변화"
SCHEMA = "price-analytics/1"
KG_PROVIDER = "KG모바일"
KG_ERROR_STAGES = {"KG 목록/API 전체", "상세 API 조회", "요금/할인 데이터 검증"}
NAVY, BLUE, TEAL, LIGHT, TEXT = "13263E", "235B99", "007D7D", "E8F0FA", "20344D"


def label(value):
    return str(value if value is not None else "").strip()


def normalized(value):
    # Normalize width, case and whitespace, not underlying allowance units/specs.
    return "".join(unicodedata.normalize("NFKC", label(value)).casefold().split())


def valid_price(value):
    return (type(value) in (int, float)
            and math.isfinite(value) and value >= 0 and int(value) == value)


def diagnostic_failures(stamp, run_id, directory):
    """Read only the KG diagnostic from the same workflow run and attempt."""
    if not re.fullmatch(r"\d+_\d+", run_id):
        return set()
    path = directory / f"mvno_kg_diagnostics_{stamp}_{run_id}.jsonl"
    if not path.is_file():
        return set()
    try:
        with path.open(encoding="utf-8-sig") as records:
            for line in records:
                if not line.strip():
                    continue
                item = json.loads(line)
                if isinstance(item, dict) and item.get("stage") in KG_ERROR_STAGES and item.get("error_type") and item.get("error"):
                    return {KG_PROVIDER}
    except (OSError, UnicodeError, ValueError, TypeError):
        # A broken same-run diagnostic is not evidence of complete collection.
        return {KG_PROVIDER}
    return set()


def recorded_diagnostic_status(book):
    if DASHBOARD not in book.sheetnames or book[DASHBOARD]["X1"].value != SCHEMA:
        return None
    value = book[DASHBOARD]["Y1"].value
    if not isinstance(value, str):
        return None
    try:
        statuses = json.loads(value)
    except ValueError:
        return None
    if not isinstance(statuses, list) or any(not isinstance(item, str) for item in statuses):
        return None
    return set(statuses)


def failed_providers(book):
    found = recorded_diagnostic_status(book) or set()
    if "수집 오류" in book.sheetnames:
        found.update(
            label(provider) for (provider,) in book["수집 오류"].iter_rows(
                min_row=2, max_col=1, values_only=True
            ) if label(provider)
        )
    return found


@dataclass
class Segment:
    group_count: int
    rows: list
    winning_groups: list


def rankings(book, extra_failures=()):
    failed = failed_providers(book) | set(extra_failures)
    failed_keys = {normalized(carrier) for carrier in failed}
    by_segment = {"RS": defaultdict(dict), "RM": defaultdict(dict)}
    displays, rejected = {}, Counter()
    # Match the existing crawler/Excel classification (참조 C&D&E → G, else RM).
    reference = {
        "".join(str(v or "") for v in (voice, sms, data)): label(kind)
        for voice, sms, data, _, kind in book["참조"].iter_rows(
            min_row=3, min_col=3, max_col=7, values_only=True
        ) if any(v is not None for v in (voice, sms, data))
    }
    group_labels = {}
    for carrier, plan, voice, sms, data, network, generation, monthly in book["홈페이지"].iter_rows(
        min_row=2, max_col=8, values_only=True
    ):
        if not any(v is not None and v != "" for v in (carrier, plan, voice, sms, data, network, generation, monthly)):
            continue
        company = normalized(carrier)
        if company in failed_keys:
            rejected["수집 실패 사업자 행"] += 1
            continue
        if not (company and label(plan) and label(network) and label(generation)):
            rejected["사업자·요금제·망·세대 누락"] += 1
            continue
        allowance = tuple(normalized(v) for v in (voice, sms, data))
        if not all(allowance):
            rejected["통화·문자·데이터 누락"] += 1
            continue
        if not valid_price(monthly):
            rejected["월 요금 누락·오류"] += 1
            continue
        kind = "RS" if reference.get("".join(str(v or "") for v in (voice, sms, data))) == "RS" else "RM"
        displays.setdefault(company, label(carrier))
        group_labels.setdefault((kind, allowance), f"{label(data)} / 통화 {label(voice)} / 문자 {label(sms)}")
        prices = by_segment[kind][allowance]
        prices[company] = min(prices.get(company, math.inf), int(monthly))

    result = {}
    for kind, groups in by_segment.items():
        portfolios, wins = Counter(), defaultdict(list)
        for group, prices in groups.items():
            best = min(prices.values())
            for company, price in prices.items():
                portfolios[company] += 1
                if price == best:
                    wins[company].append((group_labels[(kind, group)], price))
        company_order = sorted(portfolios, key=lambda c: (-len(wins[c]), -Fraction(len(wins[c]), portfolios[c]), displays[c]))
        rows, all_wins, previous_score, rank = [], [], None, 0
        for position, company in enumerate(company_order, 1):
            win, sold = len(wins[company]), portfolios[company]
            score = (win, Fraction(win, sold))
            if score != previous_score:
                rank, previous_score = position, score
            strengths = sorted(wins[company], key=lambda item: (item[1], item[0]))
            all_wins.extend((displays[company], group, price) for group, price in strengths)
            sample = " · ".join(f"{group} {price:,}원" for group, price in strengths[:3])
            if win > 3:
                sample += f" 외 {win - 3}군"
            rows.append((rank, displays[company], win, sold, win / sold, sold / len(groups), sample or "—"))
        result[kind] = Segment(len(groups), rows, all_wins)
    return result, rejected, failed


def _group_prices(rows):
    lookup = {}
    for group, carrier, network, price in rows:
        if not valid_price(price):
            lookup.setdefault(group, (None, set()))
            continue
        if group not in lookup or lookup[group][0] is None:
            lookup[group] = (int(price), set())
        stored_price, winners = lookup[group]
        if price != stored_price:
            raise ValueError(f"중복 최저가 그룹 표시명: {group}")
        winners.add((label(carrier), label(network)))
    return lookup


def _winners(items):
    return ", ".join(f"{company} ({network})" for company, network in sorted(items))


def _date_label(stamp):
    try:
        return datetime.strptime(stamp, "%Y%m%d_%H%M%S").strftime("%m/%d %H:%M")
    except (ValueError, TypeError):
        return None


def build_mail_change(current, previous, lowest_fn, stamp=None, previous_stamp=None,
                      current_diagnostics=()):
    """Compare one actual prior run; never compare partial/ambiguous minima."""
    dates = _date_label(previous_stamp), _date_label(stamp)
    timeframe = f"비교 기준: {dates[0]} → {dates[1]} (한국시간). " if all(dates) else ""
    if previous is None or not previous.is_file():
        return {"note": "첫 회차이거나 직전 비교 파일이 없어 비교하지 않았습니다.", "rows": []}
    try:
        current_rows, current_failed, current_invalid, current_providers, _ = lowest_fn(current, quality=True)
        old_rows, old_failed, old_invalid, old_providers, prior_diagnostics = lowest_fn(previous, quality=True)
        current_failed |= set(current_diagnostics)
        before, now = _group_prices(old_rows), _group_prices(current_rows)
    except (OSError, KeyError, BadZipFile, InvalidFileException, ValueError):
        return {"note": timeframe + "직전 회차·이번 회차의 비교 자료를 확인할 수 없어 보류했습니다.", "rows": []}
    if current_failed or old_failed:
        fail_info = []
        if old_failed:
            fail_info.append("직전: " + ", ".join(sorted(old_failed)))
        if current_failed:
            fail_info.append("이번: " + ", ".join(sorted(current_failed)))
        return {"note": timeframe + "수집 오류로 최저가 변화를 확정하지 않았습니다. (" + " / ".join(fail_info) + ")", "rows": []}
    if prior_diagnostics is None and KG_PROVIDER in old_providers:
        return {"note": timeframe + "직전 회차의 KG모바일 상세 수집 상태를 검증할 수 없어 비교를 보류했습니다.", "rows": []}
    if not current_providers or current_providers != old_providers:
        return {"note": timeframe + "두 회차 수집 사업자 범위가 달라 최저가 비교를 보류했습니다.", "rows": []}
    if set(before) != set(now):
        return {"note": timeframe + "두 회차의 요금제군 정의가 달라 최저가 비교를 보류했습니다.", "rows": []}

    changes, withheld = [], []
    after_available = 0
    for group, (new_price, new_winners) in now.items():
        old_price, old_winners = before[group]
        if group in current_invalid or group in old_invalid or (old_price is None) != (new_price is None):
            withheld.append(group)
            continue
        if new_price is None:
            continue
        after_available += 1
        if new_price == old_price and new_winners == old_winners:
            continue
        if new_price == old_price:
            change = "최저가 유지 · 동률 사업자/망 변경"
        else:
            amount = new_price - old_price
            change = f"{amount:+,}원" + (" · 최저가 사업자/망 변경" if new_winners != old_winners else "")
        changes.append((group, old_price, _winners(old_winners), new_price, _winners(new_winners), change))
    note = timeframe
    if not changes:
        note += "비교 가능한 요금제군에서 최저가 및 동률 사업자·망 변경이 없습니다." if after_available else "비교 가능한 최저가 자료가 없습니다."
    else:
        note += f"비교 가능한 {after_available}개 군 중 {len(changes)}개 군에 변화가 있습니다."
    if withheld:
        note += f" 한쪽에서만 확인됐거나 가격·망 정보가 불완전한 {len(withheld)}개 군은 비교 보류: " + ", ".join(withheld) + "."
    return {"note": note, "rows": changes}


def _value(ws, row, col, value):
    cell = ws.cell(row, col, value)
    if isinstance(value, str):
        # Data sourced from third-party plan text must be displayed, never executed as formulas.
        cell.data_type = "s"
    return cell


def _band(ws, row, value, color, size=12):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    cell = _value(ws, row, 1, value)
    cell.fill = PatternFill("solid", fgColor=color)
    cell.font = Font(name="맑은 고딕", size=size, bold=True, color="FFFFFF")
    cell.alignment = Alignment(vertical="center", indent=1)
    ws.row_dimensions[row].height = 28 if size < 16 else 35


def _note(ws, row, value, height=21):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=7)
    cell = _value(ws, row, 1, value)
    cell.font = Font(name="맑은 고딕", size=10, color=TEXT)
    cell.alignment = Alignment(vertical="center", wrap_text=True, indent=1)
    ws.row_dimensions[row].height = height


def write_dashboard(book, stamp, segments, rejected, failed, change, diagnostics):
    if DASHBOARD in book.sheetnames:
        del book[DASHBOARD]
    ws = book.create_sheet(DASHBOARD, 0)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 85
    ws.freeze_panes = "C7"
    ws.sheet_properties.tabColor = TEAL
    _band(ws, 1, "사업자별 월 요금 경쟁력", NAVY, 18)
    _note(ws, 2, f"수집 {_date_label(stamp) or stamp} (한국시간) | 홈페이지 · 월 납부액 기준 | 가입점유율이나 총혜택 포함 가격이 아닙니다.")
    _note(ws, 3, "순위: 최저가 군 수 → 최저가율. 공동 최저가 포함, 같은 사업자는 한 군에 1회. 군은 통화·문자·데이터 제공량이 동일한 조합입니다.")
    _note(ws, 4, "최저가율 = 최저가 군 ÷ 판매 군 | 커버리지 = 판매 군 ÷ 해당 RS/RM 유효 전체 군 | 망/세대는 동일 제공량군 내에서 경쟁합니다.")
    failed_part = ("수집 실패 사업자 제외: " + ", ".join(sorted(failed)) + ". 순위는 수집된 범위 기준. ") if failed else "기록된 수집 실패 없음. "
    rejected_part = (" | 제외 행: " + " · ".join(f"{k} {v:,}건" for k, v in sorted(rejected.items()))) if rejected else " | 필수값 제외 없음"
    _note(ws, 5, failed_part + rejected_part, 31)
    _note(ws, 6, "대표 강점 3개 외에 요금제군별 전체 최저가 결과는 순위표 아래에 있습니다.")
    row = 7
    cols = ("순위", "사업자", "최저가 군", "판매 군", "최저가율", "군 커버리지", "강점: 최저가 요금제군 예시 (최대 3)")
    for kind, color in (("RS", BLUE), ("RM", TEAL)):
        segment = segments[kind]
        _band(ws, row, f"{kind}  |  유효 요금제군 {segment.group_count:,}개 · 사업자 {len(segment.rows):,}곳", color)
        row += 1
        for col, title in enumerate(cols, 1):
            heading = _value(ws, row, col, title)
            heading.font = Font(name="맑은 고딕", size=10, bold=True, color=TEXT)
            heading.fill = PatternFill("solid", fgColor=LIGHT)
            heading.alignment = Alignment(horizontal="center" if col != 7 else "left", vertical="center")
        ws.row_dimensions[row].height = 22
        row += 1
        start = row
        if not segment.rows:
            _note(ws, row, "유효한 요금제가 없습니다.")
            row += 1
        for item in segment.rows:
            for col, value in enumerate(item, 1):
                cell = _value(ws, row, col, value)
                cell.alignment = Alignment(horizontal="left" if col in (2, 7) else "center", vertical="center", wrap_text=col == 7)
                cell.font = Font(name="맑은 고딕", size=10, bold=col in (2, 3), color=TEXT)
                if row % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F3F7FB")
                if col in (5, 6):
                    cell.number_format = "0%"
            ws.row_dimensions[row].height = 34 if len(item[-1]) > 56 else 22
            row += 1
        if segment.rows:
            ws.conditional_formatting.add(f"C{start}:C{row - 1}", DataBarRule(start_type="num", start_value=0, end_type="max", color="73ABD7"))
            ws.conditional_formatting.add(f"F{start}:F{row - 1}", DataBarRule(start_type="num", start_value=0, end_type="num", end_value=1, color="9ACBC6"))
        row += 2
    _band(ws, row, "사업자별 전체 최저가 요금제군", NAVY)
    row += 1
    for col, title in ((1, "구분"), (2, "사업자"), (3, "월 요금(원)"), (7, "요금제군")):
        heading = _value(ws, row, col, title)
        heading.font = Font(name="맑은 고딕", size=10, bold=True, color=TEXT)
        heading.fill = PatternFill("solid", fgColor=LIGHT)
    row += 1
    for kind in ("RS", "RM"):
        for carrier, name, price in segments[kind].winning_groups:
            for col, value in ((1, kind), (2, carrier), (3, price), (7, name)):
                cell = _value(ws, row, col, value)
                cell.font = Font(name="맑은 고딕", size=10, color=TEXT)
                if col == 3:
                    cell.number_format = "#,##0"
            row += 1
    for col, width in enumerate((9, 24, 13, 12, 14, 16, 65), 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    # Freeze the exact report-time comparison so mail never reads a replaced history/latest.xlsx.
    ws.column_dimensions["X"].hidden = True
    ws.column_dimensions["Y"].hidden = True
    _value(ws, 1, 25, json.dumps(sorted(diagnostics), ensure_ascii=False))
    _value(ws, 1, 24, SCHEMA)
    _value(ws, 2, 24, change["note"])
    for idx, item in enumerate(change["rows"], 3):
        _value(ws, idx, 24, json.dumps(item, ensure_ascii=False))
    return ws


def prepare_report(report, previous, stamp, previous_stamp, lowest_fn, baseline_issue=None,
                   diagnostics=()):
    with_report = load_workbook(report)
    tmp_path = None
    try:
        segments, rejected, failed = rankings(with_report, diagnostics)
        change = ({"note": baseline_issue, "rows": []} if baseline_issue
                  else build_mail_change(report, previous, lowest_fn, stamp, previous_stamp, diagnostics))
        write_dashboard(with_report, stamp, segments, rejected, failed, change, diagnostics)
        # Same-directory replacement preserves the existing result if generation fails.
        with NamedTemporaryFile(dir=report.parent, suffix=".xlsx", delete=False) as temp:
            tmp_path = Path(temp.name)
        with_report.save(tmp_path)
        tmp_path.replace(report)
    finally:
        with_report.close()
        if tmp_path is not None:
            tmp_path.unlink(missing_ok=True)


def read_mail_change(report):
    fallback = ("직전 회차 비교 데이터가 이 파일에 없어 비교하지 않았습니다.", [])
    book = load_workbook(report, read_only=True, data_only=True)
    try:
        if DASHBOARD not in book.sheetnames:
            return fallback
        values = (v for (v,) in book[DASHBOARD].iter_rows(min_row=1, min_col=24, max_col=24, values_only=True))
        if next(values, None) != SCHEMA:
            return fallback
        note = next(values, None)
        return note or fallback[0], [json.loads(v) for v in values if v]
    finally:
        book.close()
