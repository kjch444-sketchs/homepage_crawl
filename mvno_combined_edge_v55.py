"""현재 저장소의 v55에 큰사람 수집 실패 처리를 추가하고 v56으로 복사."""
from pathlib import Path

root = Path(__file__).resolve().parent
source = root / 'mvno_combined_edge_v55.py'
target = root / 'mvno_combined_edge_v56.py'
workflow = root / '.github' / 'workflows' / 'mvno.yml'

if not source.is_file() or not workflow.is_file():
    raise SystemExit('이 파일을 mvno_combined_edge_v55.py와 같은 저장소 최상위 폴더에 놓으세요.')
if target.exists():
    raise SystemExit(f'{target.name}이 이미 있습니다. 덮어쓰지 않았습니다.')
code = source.read_text(encoding='utf-8')
old = '            rows.extend(crawl_eyagi(page))'
new = '''            try:
                rows.extend(crawl_eyagi(page))
            except (PlaywrightError, RuntimeError) as exc:
                COLLECTION_FAILURES["큰사람"] = str(exc)
                print(f"⚠️ 큰사람 수집 실패. 다음 사업자로 진행합니다: {exc}")'''
if code.count(old) != 1 or 'COLLECTION_FAILURES = {}' not in code:
    raise SystemExit('v55의 예상 코드와 다릅니다. 파일을 올려 주시면 정확히 수정하겠습니다.')
yaml = workflow.read_text(encoding='utf-8')
if yaml.count('mvno_combined_edge_v55.py') != 1:
    raise SystemExit('워크플로가 v55를 정확히 한 번 참조해야 합니다. 파일을 수정하지 않았습니다.')
compiled = code.replace(old, new)
compile(compiled, str(target), 'exec')
target.write_text(compiled, encoding='utf-8')
workflow.write_text(yaml.replace('mvno_combined_edge_v55.py', 'mvno_combined_edge_v56.py'), encoding='utf-8')
print('완료: mvno_combined_edge_v56.py 및 .github/workflows/mvno.yml')
