# 엑셀 매입자료 비교 웹사이트

두 파일을 업로드하면 B 전자세금계산서 파일의 셀 위치와 값을 유지한 `.xlsx` 결과를 다운로드합니다.

## 로컬 실행

Node.js 22 LTS를 사용합니다. 이 프로젝트는 `npm` 기준으로 설치합니다. `pnpm`으로 설치된 `node_modules`가 이미 있다면 먼저 삭제한 뒤 진행하세요.

```bash
nvm use 22
rm -rf node_modules
npm install
npm start
```

브라우저에서 `http://localhost:3000`을 엽니다.

## Render 배포

1. 이 폴더를 GitHub 저장소에 올립니다.
2. Render에서 **New → Blueprint**를 선택하고 저장소를 연결합니다.
3. `render.yaml`을 확인하고 배포합니다.
4. 생성된 `onrender.com` 주소를 사용합니다.

업로드 파일은 서버 메모리에서만 처리하며, 애플리케이션은 파일을 저장하거나 데이터베이스에 기록하지 않습니다. 파일당 업로드 한도는 10MB입니다.
