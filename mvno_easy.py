"""이지모바일 KT/LG의 동적 목록 표 추출."""
import re
from bs4 import BeautifulSoup


def clean(value):
    return re.sub(r'\s+', ' ', value or '').strip()


def unit(value):
    return re.sub(r'(?<=\d)\s+(?=분|건|GB|MB|Mbps)', '', clean(value))


def parse_easy_html(html, network):
    soup = BeautifulSoup(html, 'html.parser')
    table = soup.select_one('#bill-list')
    if table is None:
        raise RuntimeError(f'이지모바일 {network}: #bill-list 없음')
    result = []
    for tr in table.select('tr'):
        if not tr.select_one('td.name'):
            continue  # 사은품 안내용 행 제외
        def text(selector):
            node = tr.select_one(selector)
            return clean(node.get_text(' ', strip=True)) if node else ''
        name, data, voice, sms = [text(x) for x in ('td.name','td.data','td.phone','td.txt')]
        if re.search(r'선불|prepaid', name, re.I):
            continue
        # 정가(.fixed_price)가 아닌 실제 할인 월요금만 사용.
        monthly = re.search(r'([\d,]+)', text('td.fee .price strong') or text('td.fee .price'))
        fee = text('td.fee')
        discount = re.search(r'(\d+)\s*개월\s*(차부터|이후|후)\s*(?:평생\s*)?([\d,]+)\s*원', fee)
        if not all((name,voice,sms,monthly)) or (not data and '데이터 쉐어링' not in name):
            raise RuntimeError(f'이지모바일 {network}: 필수 항목 누락: {name}')
        if '개월' in fee and not discount:
            raise RuntimeError(f'이지모바일 {network}: 해석할 수 없는 할인 문구: {name} / {fee}')
        result.append({'사업자':'이지모바일', '요금제':name,
            '통화 제공량':unit(voice), '문자 제공량':unit(sms),
            '데이터 제공량':unit('월 '+re.sub(r'^월\s*', '', data)) if data else '',
            '망구분':network, 'LTE/5G 구분':'5G' if re.search(r'5G(?!B)', name, re.I) else 'LTE',
            '월 요금':monthly.group(1).replace(',', ''),
            '할인 기간':int(discount.group(1))-(discount.group(2)=='차부터') if discount else '',
            '기간 이후 요금':discount.group(3).replace(',', '') if discount else '',
            '요금구분':'후불', '홈페이지':'홈페이지'})
    if not result:
        raise RuntimeError(f'이지모바일 {network}: 요금제 행 없음')
    return result


def crawl_easy(page):
    rows = []
    for te, network in (('kt','KT'),('lg','LG U+')):
        url = f'https://www.egmobile.co.kr/charge/list?te={te}'
        for attempt in range(1,4):
            try:
                response = page.goto(url, wait_until='domcontentloaded', timeout=60000)
                if not response or response.status != 200:
                    raise RuntimeError(f'HTTP {response.status if response else "응답 없음"}')
                # 목록은 billSearchList()가 비동기로 채운다.
                page.wait_for_selector('#bill-list tr td.name', timeout=45000)
                selected = page.locator('.agency_tab .te.on').first.get_attribute('data-telecom')
                expected = 'KT' if te == 'kt' else 'LGU+'
                if selected != expected:
                    raise RuntimeError(f'망 선택 불일치: {selected} / {expected}')
                break
            except Exception as exc:
                print(f'이지모바일 {network} 접속/로딩 실패 {attempt}/3: {str(exc).splitlines()[0]}')
                if attempt == 3:
                    raise RuntimeError(f'이지모바일 {network} 접속/로딩 3회 실패: {str(exc).splitlines()[0]}') from exc
        visited = set()
        for page_no in range(1,201):
            # 초기 비동기 렌더가 완료될 때까지 행 내용의 안정성을 확인.
            last, stable = None, 0
            for _ in range(20):
                current = page.locator('#bill-list').inner_text()
                stable = stable+1 if current == last else 0
                if stable >= 2:
                    break
                last = current
                page.wait_for_timeout(500)
            else:
                raise RuntimeError(f'이지모바일 {network}: 목록이 안정되지 않음')
            if current in visited:
                raise RuntimeError(f'이지모바일 {network}: 페이지 목록 반복')
            visited.add(current)
            page_rows = parse_easy_html(page.content(),network)
            rows.extend(page_rows)
            print(f'이지모바일 {network} {page_no}페이지: {len(page_rows)}건')
            # 첨부 KT/LG는 페이지 영역이 비어 있으며 전체가 한 번에 표시됨.
            # 페이지 버튼이 추가될 경우 다음 숫자를 확인하여 순차 이동.
            pager = page.locator('#rate_pageing')
            if not pager.count() or not pager.inner_text().strip():
                break
            next_button = pager.get_by_text(str(page_no+1),exact=True)
            if not next_button.count():
                next_button = pager.locator('a[rel="next"], button[rel="next"], a.next, button.next').filter(visible=True)
            if not next_button.count():
                numbers = re.findall(r'\b\d+\b',pager.inner_text())
                if numbers and max(map(int,numbers)) <= page_no:
                    break
                raise RuntimeError(f'이지모바일 {network}: 다음 페이지 판별 불가')
            next_button.first.click()
            page.wait_for_function('(old) => document.querySelector("#bill-list")?.innerText !== old',arg=current,timeout=30000)
            page.wait_for_selector('#bill-list tr td.name',timeout=30000)
        else:
            raise RuntimeError(f'이지모바일 {network}: 페이지 상한 초과')
    print(f'이지모바일 합계: {len(rows)}건')
    return rows
