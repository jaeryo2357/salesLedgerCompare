# 엑셀 매입자료 비교 웹사이트

두 파일을 선택하면 **브라우저 안에서만** 비교하고, B 전자세금계산서 파일의 셀 위치와 값을 유지한 `.xlsx` 결과를 다운로드합니다.

## Render 배포

1. Render에서 **New → Blueprint**를 선택하고 이 저장소를 연결합니다.
2. `render.yaml`의 Static Site 설정을 확인하고 배포합니다.
4. 생성된 `onrender.com` 주소를 사용합니다.

Render는 정적 HTML·CSS·JavaScript 파일만 제공하며 엑셀 파일을 받지 않습니다. 엑셀 비교와 결과 생성은 사용자의 브라우저 메모리에서 이루어집니다.
