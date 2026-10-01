import { useState, useMemo, useCallback } from "react";
import { useMsal, useIsAuthenticated } from "@azure/msal-react";
import * as XLSX from "xlsx";
import { loginRequest, isClientIdConfigured } from "./authConfig";
import { loadSharedFolder } from "./graph";

// 기본 예시 링크 (사용자가 제공한 공유 폴더). 필요 시 지우고 다른 링크 입력.
const DEFAULT_SHARE_URL =
  "https://1drv.ms/f/c/745a99f9c58cd890/IgB5LBCn_n8rQIvOqbkWtIkOAeHYdoZod_Wo1B-9JtzPIF8?e=iXMNuj";

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

// 경로에서 파일명을 떼고 상위 폴더 경로까지만 반환한다.
// 예) "A/B/강임호.pdf" → "A/B", 최상위 파일 "강임호.pdf" → "" (루트)
function parentPathOf(path) {
  if (!path) return "";
  const idx = path.lastIndexOf("/");
  return idx === -1 ? "" : path.slice(0, idx);
}

export default function App() {
  const { instance, accounts } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const [shareUrl, setShareUrl] = useState(DEFAULT_SHARE_URL);
  const [query, setQuery] = useState("");
  const [items, setItems] = useState([]);
  const [rootName, setRootName] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [loadMs, setLoadMs] = useState(null); // 전체 불러오기 소요시간(ms)
  const [loaded, setLoaded] = useState(false); // 전체 목록을 한 번이라도 불러왔는지
  // 정렬 상태. 기본은 "정렬 안 함"(불러온 순서 그대로) — 헤더를 클릭해야 정렬 시작.
  const [sortKey, setSortKey] = useState(null); // null | "name" | "path" | "lastModified"
  const [sortDir, setSortDir] = useState("asc"); // "asc" | "desc"

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

  // 공유 폴더 전체를 한 번 불러온다(=full scan). 이후 검색은 아래 filtered 로 즉시 처리.
  const handleLoad = useCallback(async () => {
    setError("");
    setLoading(true);
    setItems([]);
    setLoadMs(null);
    setLoaded(false);
    try {
      const token = await getToken();
      const started = performance.now();
      const { root, items } = await loadSharedFolder(shareUrl, token);
      const elapsed = performance.now() - started;
      setLoadMs(elapsed);
      setRootName(root.name || "공유 폴더");
      setItems(items.filter((it) => !it.isFolder));
      setLoaded(true);
    } catch (e) {
      setError(e.message);
    } finally {
      setLoading(false);
    }
  }, [getToken, shareUrl]);

  // 파일명·경로 기준 클라이언트 필터링.
  // 한글 파일명은 저장소에 따라 자모 분리형(NFD)으로 올 수 있어, 완성형(NFC)으로
  // 입력한 검색어와 그대로는 매칭되지 않는다. 양쪽을 NFC 로 정규화해 비교한다.
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase().normalize("NFC");
    if (!q) return items;
    return items.filter(
      (it) =>
        it.name.toLowerCase().normalize("NFC").includes(q) ||
        it.path.toLowerCase().normalize("NFC").includes(q)
    );
  }, [items, query]);

  // 정렬. 헤더를 클릭해 sortKey 가 정해졌을 때만 정렬하고, 그 전에는 불러온 순서를
  // 그대로 쓴다(폴더 불러오기 직후 추가 정렬 비용 없음).
  const sorted = useMemo(() => {
    if (!sortKey) return filtered; // 기본: 불러온 순서 유지
    const dir = sortDir === "asc" ? 1 : -1;
    // 문자열(이름/경로)은 NFC 정규화 + 로케일 비교, 수정일은 시간값 비교.
    const cmpStr = (a, b) =>
      (a || "").normalize("NFC").localeCompare((b || "").normalize("NFC"), "ko");
    const cmpDate = (a, b) =>
      (a ? new Date(a).getTime() : 0) - (b ? new Date(b).getTime() : 0);

    const primary = (a, b) => {
      if (sortKey === "lastModified") return cmpDate(a.lastModified, b.lastModified);
      if (sortKey === "name") return cmpStr(a.name, b.name);
      return cmpStr(a.path, b.path); // "path"
    };
    // 2순위 보조: 이름 정렬이면 경로를, 경로 정렬이면 이름을 보조로. (방향 동일)
    const secondary = (a, b) => {
      if (sortKey === "name") return cmpStr(a.path, b.path);
      if (sortKey === "path") return cmpStr(a.name, b.name);
      return 0; // 수정일은 보조정렬 없음
    };

    return [...filtered].sort((a, b) => {
      const p = primary(a, b);
      if (p !== 0) return p * dir;
      return secondary(a, b) * dir;
    });
  }, [filtered, sortKey, sortDir]);

  // 헤더 클릭: 같은 키면 방향 토글, 다른 키면 그 키로 바꾸고 오름차순부터.
  const toggleSort = useCallback(
    (key) => {
      if (key === sortKey) {
        setSortDir((d) => (d === "asc" ? "desc" : "asc"));
      } else {
        setSortKey(key);
        setSortDir("asc");
      }
    },
    [sortKey]
  );

  // 헤더에 붙일 정렬 방향 표시(▲/▼). 활성 컬럼에만 표시.
  const sortArrow = (key) =>
    key === sortKey ? (sortDir === "asc" ? " ▲" : " ▼") : "";

  // 현재 화면에 표시된 결과(sorted)를 .xlsx 로 내려받는다.
  // 마지막 "파일바로열기" 열에는 webUrl 을 클릭 가능한 하이퍼링크로 넣는다.
  const handleExport = useCallback(() => {
    const header = ["이름", "경로", "크기", "수정일", "파일바로열기"];
    const rows = sorted.map((it) => [
      it.name,
      parentPathOf(it.path),
      formatSize(it.size),
      it.lastModified
        ? new Date(it.lastModified).toLocaleDateString("ko-KR")
        : "",
      "파일 열기", // 셀 표시 텍스트. 아래에서 하이퍼링크(l)를 붙인다.
    ]);
    const ws = XLSX.utils.aoa_to_sheet([header, ...rows]);

    // 마지막 열(E, 0-based 4) 각 데이터 행에 하이퍼링크 부여.
    sorted.forEach((it, i) => {
      if (!it.webUrl) return;
      const addr = XLSX.utils.encode_cell({ r: i + 1, c: 4 }); // +1: 헤더 다음 행
      const cell = ws[addr];
      if (cell) cell.l = { Target: it.webUrl, Tooltip: it.name };
    });

    // 보기 좋게 열 너비 지정.
    ws["!cols"] = [
      { wch: 40 }, // 이름
      { wch: 50 }, // 경로
      { wch: 10 }, // 크기
      { wch: 12 }, // 수정일
      { wch: 14 }, // 파일바로열기
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "검색결과");
    const base = (rootName || "공유폴더").replace(/[\\/:*?"<>|]/g, "_");
    XLSX.writeFile(wb, `${base}_검색결과.xlsx`);
  }, [sorted, rootName]);

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
          {/* 검색창(상단). 전체 목록을 불러온 뒤 입력하면 즉시 필터링된다. */}
          <div className="search-row">
            <input
              type="text"
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              placeholder={
                loaded
                  ? `"${rootName}" 안에서 파일명 검색…`
                  : "먼저 아래에서 폴더를 불러오세요 (약 10초 예상)"
              }
              disabled={!loaded}
              autoFocus
            />
            {loaded && (
              <span className="count">
                {filtered.length}개 / 전체 {items.length}개 파일
                {loadMs != null && ` · ${(loadMs / 1000).toFixed(2)}초`}
              </span>
            )}
            {loaded && (
              <button
                className="excel"
                onClick={handleExport}
                disabled={filtered.length === 0}
                title="현재 표시된 결과를 엑셀(.xlsx)로 저장"
              >
                결과 엑셀 다운로드
              </button>
            )}
          </div>

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
        </>
      )}

      {error && <div className="error">⚠️ {error}</div>}

      {filtered.length > 0 && (
        <table>
          <thead>
            <tr>
              <th className="sortable" onClick={() => toggleSort("name")}>
                이름{sortArrow("name")}
              </th>
              <th className="sortable" onClick={() => toggleSort("path")}>
                경로{sortArrow("path")}
              </th>
              <th>크기</th>
              <th
                className="sortable"
                onClick={() => toggleSort("lastModified")}
              >
                수정일{sortArrow("lastModified")}
              </th>
            </tr>
          </thead>
          <tbody>
            {sorted.map((it) => (
              <tr key={it.id}>
                <td>
                  <a href={it.webUrl} target="_blank" rel="noreferrer">
                    {it.name}
                  </a>
                </td>
                <td className="path">{parentPathOf(it.path)}</td>
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

      {loaded && filtered.length === 0 && !loading && (
        <p className="hint">검색 결과가 없습니다.</p>
      )}
    </div>
  );
}
