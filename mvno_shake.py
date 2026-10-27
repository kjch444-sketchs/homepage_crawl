"""쉐이크모바일 공개 요금제 API: 사이트 탭 목록을 동적으로 순회."""
import re
from decimal import Decimal, InvalidOperation
from bs4 import BeautifulSoup

BASE = 'https://shakemobile.co.kr'


def text(value):
    return re.sub(r'\s+', ' ', BeautifulSoup(str(value or ''), 'html.parser').get_text(' ', strip=True)).strip()


def amount(value, field):
    if value is None or str(value).strip() == '':
        raise RuntimeError(f'쉐이크모바일 {field} 값 없음')
    try:
        return Decimal(str(value).replace(',', '').strip())
    except InvalidOperation:
        raise RuntimeError(f'쉐이크모바일 {field} 숫자 오류: {value}') from None


def parse_shake_plan(plan):
    name = text(plan.get('gdnm'))
    if re.search(r'선불|prepaid', name, re.I):
        return None
    voice, sms, data = [text(plan.get(k)) for k in ('voiceamountv','letteramountv','seedata')]
    if not all((name,voice,sms,data)):
        raise RuntimeError(f'쉐이크모바일 필수 항목 누락: {plan.get("gdcd")} / {name}')
    monthly = amount(plan.get('tt_AMT'),'월 요금')
    period = amount(plan.get('promo_MM'),'할인 기간')
    if monthly < 0 or period < 0 or period != int(period) or monthly != int(monthly):
        raise RuntimeError(f'쉐이크모바일 요금/기간 값 오류: {name}')
    later = ''
    if period:
        # 사이트 prnTG01Box()와 동일: amt - parseInt(life_TIME_DISCOUNT * 1.1)
        later = int(amount(plan.get('amt'),'정가')) - int(amount(plan.get('life_TIME_DISCOUNT'),'평생 할인액')*Decimal('1.1'))
        if later < 0:
            raise RuntimeError(f'쉐이크모바일 종료 요금 음수: {name}')
    def unit(s):
        return re.sub(r'(?<=\d)\s+(?=분|건|GB|MB|Mbps)', '', s)
    return {'사업자':'쉐이크모바일', '요금제':name, '통화 제공량':unit(voice),
        '문자 제공량':unit(sms), '데이터 제공량':unit('월 '+re.sub(r'^월\s*','',data)),
        '망구분':'KT',  # 현재 사이트의 요금제 카드가 고정 표시하는 망
        'LTE/5G 구분':'5G' if str(plan.get('netdiv','')).upper() == '5G' or re.search(r'5G(?!B)',name,re.I) else 'LTE',
        '월 요금':int(monthly), '할인 기간':int(period) if period else '',
        '기간 이후 요금':later, '요금구분':'후불','홈페이지':'홈페이지'}


def crawl_shake(page):
    response = page.goto(BASE+'/M2Mobile/Set3G',wait_until='domcontentloaded',timeout=60000)
    if not response or response.status != 200:
        raise RuntimeError('쉐이크모바일 초기 페이지 접속 실패')
    def query(path, payload=None):
        for attempt in range(1,4):
            try:
                kwargs = {'timeout':30000}
                if payload is not None:
                    kwargs['data'] = payload
                result = page.request.post(BASE+path, **kwargs)
                if not result.ok:
                    raise RuntimeError(f'HTTP {result.status}')
                data = result.json()
                if not isinstance(data,list) or any(not isinstance(item,dict) for item in data):
                    raise RuntimeError('API 응답이 객체 목록이 아님')
                return data
            except Exception:
                if attempt == 3:
                    raise
                page.wait_for_timeout(1000*attempt)
    tabs = query('/M2Mobile/getTG01NETDIV')
    if not tabs:
        raise RuntimeError('쉐이크모바일 탭 목록 없음')
    seen_tabs, products = set(), {}
    for tab in tabs:
        code = tab.get('ltediv')
        if not isinstance(code,str) or not code:
            raise RuntimeError('쉐이크모바일 탭 코드 누락')
        if code in seen_tabs:
            continue
        seen_tabs.add(code)
        plans = query('/M2Mobile/getTG01ByLtediv',{'sLteDiv':code,'promoPrtnChk':[],'sort':'RECOMMEND'})
        for plan in plans:
            gdcd = str(plan.get('gdcd') or '').strip()
            if not gdcd:
                raise RuntimeError(f'쉐이크모바일 {code}: 상품 코드 누락')
            row = parse_shake_plan(plan)
            if row is None:
                continue
            if gdcd in products and products[gdcd] != row:
                raise RuntimeError(f'쉐이크모바일 같은 상품 코드의 요금 정보 불일치: {gdcd}')
            products[gdcd] = row
        print(f'쉐이크모바일 {tab.get("ltediv_NM",code)}: {len(plans)}건')
    if not products:
        raise RuntimeError('쉐이크모바일 수집 가능한 후불 요금제 없음')
    print(f'쉐이크모바일 합계: {len(products)}건 (상품 코드 기준)')
    return list(products.values())
