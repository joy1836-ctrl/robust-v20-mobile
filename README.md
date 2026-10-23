# ROBUST V20 Mobile

매일 평일 한국시간 06:30경 GitHub Actions가 동결된 V20 엔진으로 신호를 계산해 `docs/signal.json`을 갱신하고
텔레그램으로 알립니다. `docs/index.html`은 폰에 설치하는 읽기 전용 앱입니다.

- 계좌 정보는 이 저장소에 저장되지 않습니다 (앱은 폰 안에, 선택 항목 `ACCOUNT_JSON`은 GitHub Secrets에).
- 설정 순서는 `SETUP.md`를 보세요.
- 이 도구는 신호와 주문 수량만 계산하며 자동 주문을 하지 않습니다.
