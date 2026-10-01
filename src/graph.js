import { GRAPH_BASE } from "./authConfig";

/**
 * OneDrive 공유 링크를 Graph API 의 share ID 로 변환합니다.
 * 규칙: URL을 base64url 인코딩 → 패딩(=) 제거 → 앞에 "u!" 를 붙임.
 * 참고: https://learn.microsoft.com/graph/api/shares-get
 */
export function encodeShareUrl(shareUrl) {
  const b64 = btoa(unescape(encodeURIComponent(shareUrl.trim())));
  const b64url = b64.replace(/=+$/, "").replace(/\//g, "_").replace(/\+/g, "-");
  return "u!" + b64url;
}

async function graphGet(url, accessToken) {
  const res = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
  });
  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Graph 요청 실패 (${res.status}): ${body}`);
  }
  return res.json();
}

/**
 * 공유 링크가 가리키는 항목(폴더 또는 파일)의 메타데이터를 가져옵니다.
 */
export async function getSharedItem(shareUrl, accessToken) {
  const shareId = encodeShareUrl(shareUrl);
  return graphGet(`${GRAPH_BASE}/shares/${shareId}/driveItem`, accessToken);
}

/**
 * 비동기 작업 배열을 동시성 상한(limit) 안에서 병렬 실행합니다.
 * 모든 폴더를 무제한으로 동시에 치면 Graph 가 429(스로틀링)를 돌려주므로,
 * 적당한 상한을 두고 작업이 끝나는 대로 다음 작업을 밀어 넣습니다.
 */
async function mapWithConcurrency(tasks, limit, worker) {
  const results = new Array(tasks.length);
  let next = 0;
  async function run() {
    while (next < tasks.length) {
      const i = next++;
      results[i] = await worker(tasks[i], i);
    }
  }
  const runners = Array.from({ length: Math.min(limit, tasks.length) }, run);
  await Promise.all(runners);
  return results;
}

// children 호출 시 꼭 필요한 필드만 받아 페이로드·파싱 비용을 줄인다.
const CHILDREN_SELECT =
  "id,name,size,lastModifiedDateTime,webUrl,folder,file";

// 결과에서 숨길 OS 생성 시스템 파일(대소문자 무시). 썸네일 캐시 등.
const HIDDEN_FILES = new Set(["thumbs.db", ".ds_store", "desktop.ini"]);

/**
 * 공유 폴더 내부의 모든 파일/하위폴더를 재귀적으로 수집합니다.
 * driveId + itemId 기준으로 children 을 순회합니다.
 * 같은 레벨의 하위 폴더들은 병렬로(동시성 상한 내) 순회해 속도를 높입니다.
 * @returns {Promise<Array>} 평탄화된 항목 배열 (경로 정보 포함)
 */
export async function listAllItems(driveId, itemId, accessToken, parentPath = "") {
  const results = [];
  const subFolders = []; // 이 폴더 아래에서 추가로 순회할 하위 폴더들
  let url = `${GRAPH_BASE}/drives/${driveId}/items/${itemId}/children?$top=200&$select=${CHILDREN_SELECT}`;

  // 1) 현재 폴더의 모든 페이지를 받아 항목을 기록하고, 하위 폴더 목록을 모은다.
  while (url) {
    const page = await graphGet(url, accessToken);
    for (const item of page.value) {
      const itemPath = parentPath ? `${parentPath}/${item.name}` : item.name;
      // Thumbs.db 등 OS 가 만든 시스템 파일은 결과에서 숨긴다(폴더는 영향 없음).
      if (!item.folder && HIDDEN_FILES.has(item.name.toLowerCase())) {
        continue;
      }
      results.push({
        id: item.id,
        name: item.name,
        path: itemPath,
        isFolder: !!item.folder,
        size: item.size,
        lastModified: item.lastModifiedDateTime,
        webUrl: item.webUrl,
        childCount: item.folder?.childCount ?? 0,
      });
      // 주의: childCount 를 신뢰하지 않는다. OneDrive 가 동기화 지연 등으로
      // 하위 항목이 있는데도 childCount 를 0 으로 보고하는 경우가 있어,
      // 그 폴더 안의 파일이 통째로 누락되던 버그가 있었다. folder 여부만 보고 재귀한다.
      if (item.folder) {
        subFolders.push({ id: item.id, path: itemPath });
      }
    }
    url = page["@odata.nextLink"] || null;
  }

  // 2) 하위 폴더들을 병렬로 순회한다(동시성 상한 8). 순차 await 대비 큰 폴더에서 수 배 빨라진다.
  const subResults = await mapWithConcurrency(subFolders, 8, (f) =>
    listAllItems(driveId, f.id, accessToken, f.path)
  );
  for (const sub of subResults) results.push(...sub);

  return results;
}

/**
 * 공유 링크로부터 폴더 전체 파일 목록을 가져옵니다.
 * @returns {Promise<{root: object, items: Array}>}
 */
export async function loadSharedFolder(shareUrl, accessToken) {
  const root = await getSharedItem(shareUrl, accessToken);
  const driveId = root.parentReference?.driveId;
  if (!root.folder) {
    // 공유 링크가 단일 파일을 가리키는 경우
    return {
      root,
      items: [
        {
          id: root.id,
          name: root.name,
          path: root.name,
          isFolder: false,
          size: root.size,
          lastModified: root.lastModifiedDateTime,
          webUrl: root.webUrl,
        },
      ],
    };
  }
  const items = await listAllItems(driveId, root.id, accessToken);
  return { root, items };
}
