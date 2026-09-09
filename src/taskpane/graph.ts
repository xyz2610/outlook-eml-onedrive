import { GRAPH_BASE, GRAPH_SCOPES } from "./config";
import { getAccessToken } from "./auth";

export type DriveFolder = {
  id: string;
  name: string;
  path: string;
  webUrl?: string;
};

type DriveItem = {
  id: string;
  name: string;
  webUrl?: string;
  folder?: { childCount?: number };
};

async function graphFetch(path: string, init: RequestInit = {}): Promise<Response> {
  const token = await getAccessToken(GRAPH_SCOPES);
  const headers = new Headers(init.headers || {});
  headers.set("Authorization", `Bearer ${token}`);

  const url = path.startsWith("https://") ? path : `${GRAPH_BASE}${path}`;
  const response = await fetch(url, { ...init, headers });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(`Graph ${response.status}: ${detail || response.statusText}`);
  }
  return response;
}

export async function getOneDriveRoot(): Promise<DriveFolder> {
  const response = await graphFetch("/me/drive/root?$select=id,name,webUrl");
  const item = (await response.json()) as DriveItem;
  return { id: item.id, name: "OneDrive", path: "/", webUrl: item.webUrl };
}

export async function listFolders(parent: DriveFolder): Promise<DriveFolder[]> {
  const query = encodeURIComponent("id,name,folder,webUrl");
  let next: string | undefined = `/me/drive/items/${encodeURIComponent(parent.id)}/children?$select=${query}`;
  const items: DriveItem[] = [];

  while (next) {
    const response = await graphFetch(next);
    const data = (await response.json()) as { value: DriveItem[]; "@odata.nextLink"?: string };
    items.push(...data.value);
    next = data["@odata.nextLink"];
  }

  return items
    .filter((item) => Boolean(item.folder))
    .map((item) => ({
      id: item.id,
      name: item.name,
      path: parent.path === "/" ? `/${item.name}` : `${parent.path}/${item.name}`,
      webUrl: item.webUrl,
    }))
    .sort((a, b) => a.name.localeCompare(b.name, undefined, { sensitivity: "base" }));
}

export async function createFolder(parent: DriveFolder, name: string): Promise<DriveFolder> {
  const response = await graphFetch(`/me/drive/items/${encodeURIComponent(parent.id)}/children`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    }),
  });
  const item = (await response.json()) as DriveItem;
  return {
    id: item.id,
    name: item.name,
    path: parent.path === "/" ? `/${item.name}` : `${parent.path}/${item.name}`,
    webUrl: item.webUrl,
  };
}

export async function getMessageMime(messageId: string): Promise<ArrayBuffer> {
  const response = await graphFetch(`/me/messages/${encodeURIComponent(messageId)}/$value`, {
    headers: { Accept: "message/rfc822, application/octet-stream" },
  });
  return response.arrayBuffer();
}

export async function uploadEml(parent: DriveFolder, filename: string, bytes: ArrayBuffer): Promise<DriveItem> {
  const safeSegment = encodeURIComponent(filename);
  const response = await graphFetch(
    `/me/drive/items/${encodeURIComponent(parent.id)}:/${safeSegment}:/content?@microsoft.graph.conflictBehavior=rename`,
    {
      method: "PUT",
      headers: { "Content-Type": "message/rfc822" },
      body: bytes,
    }
  );
  return (await response.json()) as DriveItem;
}
