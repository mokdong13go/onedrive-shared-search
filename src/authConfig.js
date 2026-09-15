/**
 * MSAL(Microsoft 인증 라이브러리) 설정.
 *
 * clientId 는 Azure Portal에서 "앱 등록" 후 발급받은 Application (client) ID 로 교체하세요.
 * 자세한 등록 방법은 README.md 참고.
 */

// 개발 중에는 .env 의 VITE_MS_CLIENT_ID 를 사용하고, 없으면 아래 자리표시자를 씁니다.
const PLACEHOLDER = "YOUR_AZURE_CLIENT_ID_HERE";
const clientId = import.meta.env.VITE_MS_CLIENT_ID || PLACEHOLDER;

// client ID 가 아직 설정되지 않았는지 여부. UI 에서 안내 문구를 띄우는 데 사용.
export const isClientIdConfigured = clientId !== PLACEHOLDER && clientId.length > 0;

// 리디렉션 URI. window.location.origin 은 경로(/onedrive-shared-search/)를 빼버리므로
// Vite 의 BASE_URL(배포 시 "/onedrive-shared-search/", 로컬은 "/")을 붙여
// Azure 에 등록한 값과 정확히 일치시킨다.
const redirectUri = window.location.origin + import.meta.env.BASE_URL;

export const msalConfig = {
  auth: {
    clientId,
    // "common" = 개인/회사 계정 모두 허용. 특정 조직만 허용하려면 테넌트 ID로 교체.
    authority: "https://login.microsoftonline.com/common",
    redirectUri,
  },
  cache: {
    cacheLocation: "sessionStorage",
    storeAuthStateInCookie: false,
  },
};

// 로그인 시 요청할 권한(scope). 파일 읽기 전용.
export const loginRequest = {
  scopes: ["User.Read", "Files.Read.All"],
};

export const GRAPH_BASE = "https://graph.microsoft.com/v1.0";
