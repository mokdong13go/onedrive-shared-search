# OneDrive 공유 폴더 검색 앱

Microsoft 계정으로 로그인한 뒤, OneDrive/SharePoint **공유 링크**가 가리키는
폴더 안의 파일을 **파일명·경로 기준으로 검색**하는 React 웹 앱입니다.

## 동작 방식

1. 사용자가 "Microsoft 계정으로 로그인" (OAuth 2.0 / MSAL)
2. 공유 링크 입력 → Graph API `share ID`로 변환
3. `GET /shares/{id}/driveItem` 로 폴더 정보 획득
4. `children` 을 재귀 순회해 전체 파일 목록 수집
5. 브라우저에서 파일명으로 실시간 필터링

> 참고: "링크 있는 모두" 익명 공유라도 Graph API 호출에는 **로그인 토큰이 필요**합니다.
> 로그인한 계정이 해당 공유 링크에 접근 권한만 있으면 됩니다.

## 사전 준비: Azure 앱 등록 (무료, 1회)

1. https://portal.azure.com 접속 → **Microsoft Entra ID** → **앱 등록** → **새 등록**
2. 이름 입력 (예: `onedrive-search`)
3. **지원되는 계정 유형**:
   - 개인 + 회사 계정 모두 → "모든 조직 디렉터리의 계정 및 개인 Microsoft 계정"
4. **리디렉션 URI**: 플랫폼 `SPA(단일 페이지 애플리케이션)` 선택 후
   - 개발용: `http://localhost:3000`
   - 배포용: 실제 도메인 (예: `https://myapp.example.com`)
5. 등록 완료 후 **개요** 화면의 **애플리케이션(클라이언트) ID** 복사
6. **API 권한** → **권한 추가** → Microsoft Graph → 위임된 권한:
   - `User.Read`
   - `Files.Read.All`
   - (원한다면 "관리자 동의" 부여)

## 설정

프로젝트 루트에 `.env` 파일 생성:

```
VITE_MS_CLIENT_ID=여기에_복사한_클라이언트_ID
```

(또는 `src/authConfig.js` 의 `YOUR_AZURE_CLIENT_ID_HERE` 를 직접 교체)

## 실행

```bash
npm install
npm run dev
```

브라우저에서 http://localhost:3000 접속.

## 배포

```bash
npm run build   # dist/ 생성
npm run preview # 로컬 미리보기
```

`dist/` 를 GitHub Pages, Azure Static Web Apps, Netlify 등 정적 호스팅에 올리면 됩니다.
배포 도메인을 Azure 앱 등록의 리디렉션 URI에 추가하는 것을 잊지 마세요.

## 알려진 제약

- **파일 내용 검색은 미지원** (파일명·경로·메타데이터만). 내용 검색이 필요하면
  Graph 검색 API 또는 문서 파싱 로직을 추가해야 합니다.
- 폴더가 매우 크면 재귀 순회에 시간이 걸립니다. (페이지네이션은 처리됨)
- SharePoint 공유 링크도 동일한 방식으로 동작합니다.
