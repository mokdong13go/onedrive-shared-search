import { useState, useMemo, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import { loginRequest, isClientIdConfigured } from "./authConfig";
import { loadSharedFolder } from "./graph";

// 기본 예시 링크 (사용자가 제공한 공유 폴더). 필요 시 지우고 다른 링크 입력.
const DEFAULT_SHARE_URL =
  "https://1drv.ms/f/c/745a99f9c58cd890/IgC2Lb3kS8OiQrXalMTnPguvAVTtdb1-fc07NUJot79Rz6s?e=pvh4nO";

function formatSize(bytes) {
  if (bytes == null) return "";
  const units = ["B", "KB", "MB", "GB"];
  let n = bytes;
  let i = 0;
  while (n >= 1024 && i < units.length - 1) {
    n /= 1024;
    i++;
  }
  return `${n.toFixed(n < 10 && i > 0 ? 1 : 0)} ${units[i]}`;
}

export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [shareUrl, setShareUrl] = useState(DEFAULT_SHARE_URL);
  const [items, setItems] = useState([]);
  const [rootName, setRootName] = useState("");
  const [query, setQuery] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const login = useCallback(() => {
    instance.loginPopup(loginRequest).catch((e) => setError(e.message));
  }, [instance]);

  const logout = useCallback(() => {
    instance.logoutPopup().catch((e) => setError(e.message));
  }, [instance]);

  // 액세스 토큰을 조용히(silent) 획득, 실패 시 팝업으로 재시도.
  const getToken = useCallback(async () => {
    const request = { ...loginRequest, account: accounts[0] };
    try {
      const res = await instance.acquireTokenSilent(request);
      return res.accessToken;
    } catch {
      const res = await instance.acquireTokenPopup(request);
      return res.accessToken;
    }
  }, [instance, accounts]);

  const handleLoad = useCallback(async () => {
    setError("");
    setLoading(true);
    setItems([]);
    try {
      const token = await getToken();
      const { root, items } = await loadSharedFolder(shareUrl, token);
      setRootName(root.name || "공유 폴더");
      setItems(items);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [getToken, shareUrl]);

  // 파일명·경로 기준 필터링 (메타데이터 검색)
  // 한글 파일명은 저장소에 따라 자모 분리형(NFD)으로 올 수 있어, 완성형(NFC)으로
  // 입력한 검색어와 그대로는 매칭되지 않는다. 양쪽을 NFC 로 정규화해 비교한다.
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase().normalize("NFC");
    const files = items.filter((it) => !it.isFolder);
    if (!q) return files;
    return files.filter(
      (it) =>
        it.name.toLowerCase().normalize("NFC").includes(q) ||
        it.path.toLowerCase().normalize("NFC").includes(q)
    );
  }, [items, query]);

  return (
    <div className="app">
      <header>
        <h1>📁 OneDrive 공유 폴더 검색</h1>
        {!isClientIdConfigured ? null : isAuthenticated ? (
          <div className="account">
            <span>{accounts[0]?.username}</span>
            <button onClick={logout}>로그아웃</button>
          </div>
        ) : (
          <button className="primary" onClick={login}>
            Microsoft 계정으로 로그인
          </button>
        )}
      </header>

      {!isClientIdConfigured && (
        <div className="error">
          ⚙️ Azure 클라이언트 ID가 아직 설정되지 않았습니다.
          {"\n"}프로젝트 루트에 <code>.env</code> 파일을 만들고 다음을 넣으세요:
          {"\n\n"}VITE_MS_CLIENT_ID=발급받은_클라이언트_ID
          {"\n\n"}발급 방법은 README.md 의 "Azure 앱 등록" 절을 참고하세요.
        </div>
      )}

      {isClientIdConfigured && !isAuthenticated && (
        <p className="hint">
          공유 폴더를 검색하려면 먼저 Microsoft 계정으로 로그인하세요. 로그인한
          계정이 해당 공유 링크에 접근 권한이 있어야 합니다.
        </p>
      )}

      {isAuthenticated && (
        <>
          <div className="load-row">
            <input
              type="text"
              value={shareUrl}
              onChange={(e) => setShareUrl(e.target.value)}
              placeholder="OneDrive 공유 링크 (https://1drv.ms/... 또는 sharepoint.com/...)"
            />
            <button className="primary" onClick={handleLoad} disabled={loading}>
              {loading ? "불러오는 중…" : "폴더 불러오기"}
            </button>
          </div>

          {rootName && (
            <div className="search-row">
              <input
                type="text"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                placeholder={`"${rootName}" 안에서 파일명 검색…`}
              />
              <span className="count">
                {filtered.length}개 / 전체 {items.filter((i) => !i.isFolder).length}개 파일
              </span>
            </div>
          )}
        </>
      )}

      {error && <div className="error">⚠️ {error}</div>}

      {filtered.length > 0 && (
        <table>
          <thead>
            <tr>
              <th>이름</th>
              <th>경로</th>
              <th>크기</th>
              <th>수정일</th>
            </tr>
          </thead>
          <tbody>
            {filtered.map((it) => (
              <tr key={it.id}>
                <td>
                  <a href={it.webUrl} target="_blank" rel="noreferrer">
                    {it.name}
                  </a>
                </td>
                <td className="path">{it.path}</td>
                <td>{formatSize(it.size)}</td>
                <td>
                  {it.lastModified
                    ? new Date(it.lastModified).toLocaleDateString("ko-KR")
                    : ""}
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      )}

      {rootName && filtered.length === 0 && !loading && (
        <p className="hint">검색 결과가 없습니다.</p>
      )}
    </div>
  );
}
