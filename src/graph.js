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
 * 공유 폴더 내부의 모든 파일/하위폴더를 재귀적으로 수집합니다.
 * driveId + itemId 기준으로 children 을 순회합니다.
 * @returns {Promise<Array>} 평탄화된 항목 배열 (경로 정보 포함)
 */
export async function listAllItems(driveId, itemId, accessToken, parentPath = "") {
  const results = [];
  let url = `${GRAPH_BASE}/drives/${driveId}/items/${itemId}/children?$top=200`;

  while (url) {
    const page = await graphGet(url, accessToken);
    for (const item of page.value) {
      const itemPath = parentPath ? `${parentPath}/${item.name}` : item.name;
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
      // 하위 폴더 재귀 탐색
      if (item.folder && item.folder.childCount > 0) {
        const sub = await listAllItems(driveId, item.id, accessToken, itemPath);
        results.push(...sub);
      }
    }
    url = page["@odata.nextLink"] || null;
  }
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
