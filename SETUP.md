# ROBUST V20: 자동 신호 + 텔레그램 + 모바일 앱 설정 (한 번만)

매일 평일 한국시간 06:30경 GitHub가 신호를 계산해 `docs/signal.json`을 갱신하고, 텔레그램으로 알려 줍니다.
폰의 앱(`docs/index.html`)은 열 때마다 최신 `signal.json`을 읽어 화면을 그립니다.

## 1. 파일 올리기
이 폴더의 모든 파일을 저장소 루트에 올립니다. `.github` 폴더와 `docs` 폴더도 포함입니다.

## 2. GitHub 권한
Settings → Actions → General → Workflow permissions → **Read and write permissions** 저장.

## 3. 텔레그램 연결
1. 텔레그램 @BotFather 에서 봇을 만들거나 기존 봇의 토큰을 확인합니다.
2. 만든 봇에게 아무 메시지나 한 번 보냅니다.
3. 브라우저에서 `https://api.telegram.org/bot<토큰>/getUpdates` 를 열고 `"chat":{"id":숫자}` 의 숫자가 chat id 입니다.
4. 저장소 Settings → Secrets and variables → Actions → **Secrets** 에 추가합니다.
   - `TELEGRAM_BOT_TOKEN` = 봇 토큰
   - `TELEGRAM_CHAT_ID` = chat id
   - (선택) `ACCOUNT_JSON` = `{"cash":10000,"QQQ":0,"QLD":0,"TQQQ":100}` 를 넣으면 텔레그램 메시지에 내 계좌 기준 주문안이 함께 옵니다. 잔고가 바뀌면 이 값도 고쳐야 합니다.
5. (선택) **Variables** 탭에 `APP_URL` = 앱 주소를 넣으면 메시지 끝에 링크가 붙습니다.

토큰은 채팅창이나 코드 파일에 붙여 넣지 마세요. 이미 노출됐다면 BotFather에서 `/revoke` 로 새로 받으세요.

## 4. 첫 실행
Actions 탭 → `daily-signal` → **Run workflow**. 성공하면 `docs/signal.json` 이 생기고 텔레그램 메시지가 도착합니다.

## 5. 모바일 앱 켜기
1. Settings → Pages → Source: **Deploy from a branch**, Branch: `main`, Folder: `/docs` 저장.
2. 표시되는 `https://<계정>.github.io/<저장소>/` 를 폰 크롬에서 엽니다.
3. 메뉴(⋮) → **홈 화면에 추가 / 앱 설치**. 홈 화면에 ROBUST V20 아이콘이 생깁니다.
4. 앱의 **계좌설정**에서 보유수량·평단·현금을 입력합니다. 입력값은 그 폰 안에만 저장됩니다.

## 주의
- 무료 계정은 비공개 저장소에서 Pages를 쓸 수 없습니다. 저장소를 공개로 바꾸면 전략 코드와 `signal.json` 이 누구에게나 보입니다. 내 계좌 정보는 `signal.json` 에 들어가지 않습니다(앱은 폰에, `ACCOUNT_JSON` 은 비밀 저장소에 있음).
- 앱은 푸시 알림을 보내지 않습니다. 알림은 텔레그램이 담당하고, 앱은 열었을 때 최신 상태를 보여 줍니다.
- 이 설정은 실제 GitHub·텔레그램·폰에서 아직 실행해 보지 않았습니다. 첫 실행 결과를 꼭 확인하세요.
- yfinance 가 일시적으로 실패할 수 있어 3회 재시도합니다. 그래도 실패하면 Actions 탭에 빨간 표시가 남습니다.
- 메시지를 미리 보려면 PC에서 `python send_telegram.py --dry-run` (signal.json 필요).
- 주문은 직접 넣습니다. 이 도구는 신호와 수량만 계산하며, 실제 시가·수수료·매수가능액을 다시 확인하세요.
