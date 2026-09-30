# MVNO 요금제 GitHub Actions 실행

현재 v36의 수집 로직을 가져온 v38입니다. `참조` 시트는 통화·문자·데이터의 조합으로 RS를 판단합니다. 실행이 성공하면 다음 순서로 처리합니다.

1. `history/latest.xlsx`가 있으면 직전 정상 수집과 비교해 엑셀 생성
2. `history/YYYY/MM/mvno_combined_plans_날짜_시각_실행ID.xlsx`에 원본 보관
3. `history/runs.csv`에 실행 시각, 요금제·위반·변동 건수, 파일 경로와 해시 기록
4. `history/latest.xlsx` 갱신 및 저장소 커밋, Actions artifact 30일 백업
5. 메일 본문에 `가이드위반` 시트의 RS/RM 표와 룰, 첨부에 결과 엑셀; 텔레그램에는 위반 목록 발송

첫 실행에는 비교할 파일이 없으므로 `전일 대비 변동` 시트가 비어 있습니다. 그다음부터는 **전일이 아니라 직전 정상 실행**과 비교합니다. 날짜가 넘어가도 같은 기준입니다. 정상 수집에 실패하면 이력과 알림은 갱신되지 않습니다.

## 1. 저장소 준비

GitHub에서 **비공개 저장소**를 하나 만들고 이 폴더의 파일 및 `.github/workflows/mvno.yml`을 같은 경로 구조로 기본 브랜치에 올립니다. `history` 폴더는 첫 실행 때 생성됩니다. 수정한 `참조` 시트가 들어 있는 `mvno_combined_plans_reference.xlsx`를 함께 올려야 합니다. 보고서 엑셀과 실행 이력에는 사업자 요금제 정보가 들어가므로 저장소 접근 권한을 확인하세요.

GitHub 저장소 **Settings → Actions → General → Workflow permissions**에서 해당 워크플로에 `contents: write` 사용이 허용되어야 합니다. 조직 정책이 쓰기를 막으면 이력 커밋 단계가 실패합니다. 예약 실행은 기본 브랜치에 워크플로 파일이 있어야 동작합니다.

## 2. Secrets 등록

저장소 **Settings → Secrets and variables → Actions → New repository secret**에 아래 값을 넣습니다. 값은 절대 파일에 적거나 채팅에 붙여 넣지 마세요.

| 이름 | 내용 |
| --- | --- |
| `GMAIL_ADDRESS` | 발송에 사용할 전체 Gmail 주소 |
| `GMAIL_APP_PASSWORD` | Google 계정에서 발급한 앱 비밀번호 (일반 로그인 비밀번호 아님) |
| `MVNO_MAIL_TO` | 메일 받을 주소 |
| `MVNO_TELEGRAM_BOT_TOKEN` | BotFather가 발급한 봇 토큰 |
| `MVNO_TELEGRAM_CHAT_ID` | 알림 받을 텔레그램 채팅 ID |

Gmail 전송은 코드 안에서 `smtp.gmail.com:465`(SSL)를 사용하므로 SMTP 서버 설정 Secret은 필요 없습니다. 앱 비밀번호를 발급하려면 해당 Google 계정에 **2단계 인증**이 켜져 있어야 합니다. 계정 종류·보안 설정에 따라 앱 비밀번호 메뉴가 보이지 않을 수 있습니다. [Google 앱 비밀번호 안내](https://support.google.com/accounts/answer/185833)를 참조하세요.

텔레그램에서는 봇을 만든 뒤 받을 계정에서 **봇과 먼저 대화를 시작**합니다. 해당 봇의 `getUpdates` API로 채팅 ID를 확인하거나 본인이 관리하는 그룹의 채팅 ID를 확인해서 등록합니다. 봇 토큰은 URL이나 로그 화면에 남길 수 있으니 공개된 곳에 붙여 넣지 마세요.

## 3. 처음 실행

저장소 **Actions → MVNO 요금제 수집 및 알림 → Run workflow**를 기본 브랜치에서 수동 실행합니다. 성공 후 다음을 확인하세요.

- 프리티·모빙·KG·KCT 등 사업자별 수집 로그와 전체 건수
- `history/runs.csv`, `history/latest.xlsx`, 실행 시각별 원본 파일 생성
- 받은 메일의 HTML 본문에 RS/RM 위반 내용, 엑셀 첨부
- 받은 텔레그램의 RS/RM 위반 건수와 상세 목록

첫 실행에서 사용자 PC와 다른 GitHub 러너의 IP·SSL·사이트 차단으로 수집이 실패할 수 있습니다. 그 경우 `Actions` 로그의 **실패한 사업자**를 확인하고 사이트 접근 문제부터 수정해야 합니다. 크롤러가 실패하면 메일과 텔레그램은 발송되지 않습니다.

## 4. 예약 시간 변경

`.github/workflows/mvno.yml`의 `cron: '17 9,13,18 * * *'`은 **한국 시간 매일 09:17, 13:17, 18:17**입니다. 쉼표로 시각을 추가하거나 `schedule` 항목을 추가할 수 있습니다. GitHub 예약 작업은 시작 시간이 조금 늦어질 수 있습니다. 너무 촘촘하게 설정하면 실행이 겹치거나 대기 중 실행이 취소될 수 있으니 수집 소요 시간보다 간격을 넉넉히 두세요.

## 이력 보관과 용량

이 버전은 실행 원본과 `latest.xlsx`를 **Git 저장소의 `history/`에 계속 누적**합니다. 이메일은 별도의 사본이며, artifact는 30일간의 추가 백업입니다. 실행을 오래 지속하면 Git 저장소 크기가 커집니다. 몇 달간 운영 후에는 S3 등의 파일 저장소로 원본 보관처를 옮기고 `runs.csv`만 남기는 편이 관리하기 좋습니다. 민감한 가입 정보가 들어오는 경우에는 저장 전에 별도 검토가 필요합니다.

알림 전송만 실패하면 이미 생성한 이력은 남습니다. Actions에서 동일 실행을 다시 돌리면 **새 실행 ID의 수집과 알림**으로 처리되므로, 실패 채널을 확인한 뒤 재시도하세요.
