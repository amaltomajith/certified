/**
 * api.ts — typed fetch wrappers for the FastAPI backend.
 * All calls go to /api/* which Vite proxies to http://localhost:8000/*
 */

const API_BASE = import.meta.env.VITE_API_BASE_URL || '';
export const BASE = API_BASE ? API_BASE.replace(/\/$/, '') : '/api';

export type StudentDetail = {
  name: string;
  school: string;
  event_name: string;
  position?: string;
};

export type PocGroup = {
  poc_email: string;
  poc_name?: string;
  student_count: number;
  students: string[];
  student_details?: StudentDetail[];
  pdf_files?: string[];
};

export type ExcelUploadResult = {
  excel_columns: string[];
  groups: PocGroup[];
  excel_path: string;
};

export type ColumnMapping = Record<string, string>;

export type ManifestSummary = {
  groups: PocGroup[];
  total: number;
  first_pdf: string | null;
};

export type SendResult = {
  poc_name: string;
  sent_to: string;
  attachments: number;
  status: 'sent' | 'failed';
};

// ── Helpers ────────────────────────────────────────────────────────────────

async function handleJSON<T>(res: Response): Promise<T> {
  if (!res.ok) {
    let msg = res.statusText;
    try { msg = (await res.json()).detail ?? msg; } catch { /* ignore */ }
    throw new Error(msg);
  }
  return res.json() as Promise<T>;
}

/** Stream a POST request with SSE-formatted body, calling onMessage for each line. */
export async function streamPost(
  path: string,
  body: unknown,
  onMessage: (msg: string) => void,
): Promise<void> {
  const res = await fetch(`${BASE}${path}`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body),
  });

  if (!res.ok || !res.body) {
    let msg = res.statusText;
    try { msg = (await res.json()).detail ?? msg; } catch { /* ignore */ }
    throw new Error(msg);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = '';

  while (true) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    // Each SSE event is delimited by \n\n
    const parts = buffer.split('\n\n');
    buffer = parts.pop() ?? '';

    for (const part of parts) {
      for (const line of part.split('\n')) {
        if (line.startsWith('data: ')) {
          onMessage(line.slice(6));
        }
      }
    }
  }

  // Parse any leftover message in the buffer after stream closes
  if (buffer.length > 0) {
    const parts = buffer.split('\n\n');
    for (const part of parts) {
      for (const line of part.split('\n')) {
        if (line.startsWith('data: ')) {
          onMessage(line.slice(6));
        }
      }
    }
  }
}

// ── API calls ──────────────────────────────────────────────────────────────

export async function uploadExcel(file: File): Promise<ExcelUploadResult> {
  const fd = new FormData();
  fd.append('file', file);
  const res = await fetch(`${BASE}/upload/excel`, { method: 'POST', body: fd });
  return handleJSON<ExcelUploadResult>(res);
}

export async function clearExcel(): Promise<void> {
  await fetch(`${BASE}/upload/excel`, { method: 'DELETE' });
}

export async function getActiveType(): Promise<{ active_type: string }> {
  const res = await fetch(`${BASE}/config/active-type`);
  return handleJSON(res);
}

export async function setActiveType(active_type: string): Promise<{ active_type: string }> {
  const res = await fetch(`${BASE}/config/active-type`, {
    method: 'PUT',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ active_type }),
  });
  return handleJSON(res);
}

export async function uploadTemplate(
  file: File,
  winnerPosition?: '1st' | '2nd' | '3rd'
): Promise<{ template_mode: string; path: string }> {
  const fd = new FormData();
  fd.append('file', file);
  let url = `${BASE}/upload/template`;
  if (winnerPosition) {
    url += `?winner_position=${winnerPosition}`;
  }
  const res = await fetch(url, { method: 'POST', body: fd });
  return handleJSON(res);
}

export async function getColumns(): Promise<{
  columns: ColumnMapping;
  excel_path: string;
  excel_columns?: string[];
  groups?: PocGroup[];
}> {
  const res = await fetch(`${BASE}/config/columns`);
  return handleJSON(res);
}

export async function saveColumns(
  columns: ColumnMapping,
): Promise<{ groups: PocGroup[]; excel_columns: string[] }> {
  const res = await fetch(`${BASE}/config/columns`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ columns }),
  });
  return handleJSON(res);
}

export async function getCoordinatesPreview(winnerPosition?: '1st' | '2nd' | '3rd'): Promise<{ image: string; width: number; height: number }> {
  let url = `${BASE}/coordinates/preview`;
  if (winnerPosition) url += `?winner_position=${winnerPosition}`;
  const res = await fetch(url, { method: 'POST' });
  return handleJSON(res);
}

export async function getTextFields(winnerPosition?: '1st' | '2nd' | '3rd'): Promise<{ image_text_fields: Record<string, unknown> }> {
  let url = `${BASE}/config/text-fields`;
  if (winnerPosition) url += `?winner_position=${winnerPosition}`;
  const res = await fetch(url);
  return handleJSON(res);
}

export async function saveTextFields(fields: Record<string, unknown>, winnerPosition?: '1st' | '2nd' | '3rd'): Promise<void> {
  let url = `${BASE}/config/text-fields`;
  if (winnerPosition) url += `?winner_position=${winnerPosition}`;
  await fetch(url, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(fields),
  });
}

export async function getManifest(): Promise<ManifestSummary> {
  const res = await fetch(`${BASE}/manifest`);
  return handleJSON(res);
}

export async function setEmailAuth(sender_email: string, app_password: string): Promise<void> {
  const res = await fetch(`${BASE}/config/email-auth`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ sender_email, app_password }),
  });
  if (!res.ok) {
    const msg = (await res.json().catch(() => ({}))).detail ?? res.statusText;
    throw new Error(msg);
  }
}

export async function signOutEmailAuth(): Promise<void> {
  await fetch(`${BASE}/config/email-auth/signout`, { method: 'POST' });
}

export async function getEmailAuthStatus(): Promise<{ sender_email: string; password_set: boolean }> {
  const res = await fetch(`${BASE}/config/email-auth`);
  return handleJSON(res);
}

export async function getDriveAuthStatus(): Promise<{
  drive_available: boolean;
  has_credentials: boolean;
  credentials_path: string | null;
  service_account_email: string | null;
  oauth_client_available: boolean;
  oauth_token_available: boolean;
  oauth_drive_available: boolean;
  oauth_user_email: string | null;
  root_folder_id: string;
}> {
  const res = await fetch(`${BASE}/config/drive-auth`);
  return handleJSON(res);
}

export async function saveDriveRootFolder(folder_url_or_id: string): Promise<void> {
  const res = await fetch(`${BASE}/config/drive-auth/root-folder`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ folder_url_or_id }),
  });
  if (!res.ok) {
    const msg = (await res.json().catch(() => ({}))).detail ?? res.statusText;
    throw new Error(msg);
  }
}

export async function uploadDriveCredentials(file: File): Promise<void> {
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch(`${BASE}/config/drive-auth/upload`, {
    method: 'POST',
    body: formData,
  });
  if (!res.ok) {
    const msg = (await res.json().catch(() => ({}))).detail ?? res.statusText;
    throw new Error(msg);
  }
}

export async function deleteDriveCredentials(): Promise<void> {
  await fetch(`${BASE}/config/drive-auth/delete`, { method: 'POST' });
}

export async function uploadOAuthClientSecrets(file: File): Promise<void> {
  const formData = new FormData();
  formData.append('file', file);
  const res = await fetch(`${BASE}/config/drive-auth/oauth-upload`, {
    method: 'POST',
    body: formData,
  });
  if (!res.ok) {
    const msg = (await res.json().catch(() => ({}))).detail ?? res.statusText;
    throw new Error(msg);
  }
}

export async function getOAuthUrl(redirectUri: string): Promise<{ url: string; state: string }> {
  const res = await fetch(`${BASE}/config/drive-auth/oauth-url?redirect_uri=${encodeURIComponent(redirectUri)}`);
  return handleJSON(res);
}


export async function getOAuthFlowStatus(): Promise<{
  status: string;
  error: string | null;
  oauth_drive_available: boolean;
  oauth_user_email: string | null;
}> {
  const res = await fetch(`${BASE}/config/drive-auth/oauth-status`);
  return handleJSON(res);
}

export async function revokeOAuthToken(): Promise<void> {
  await fetch(`${BASE}/config/drive-auth/oauth-revoke`, { method: 'POST' });
}



export function pdfUrl(filename: string, download = false): string {
  return `${BASE}/pdf/${encodeURIComponent(filename)}${download ? '?download=true' : ''}`;
}

export function pdfPreviewUrl(filename: string): string {
  return `${BASE}/pdf-preview/${encodeURIComponent(filename)}`;
}

/**
 * Fetch a PDF and trigger a browser download with the correct filename.
 * This is more reliable than <a download> across browsers and proxies.
 */
export async function downloadPdf(filename: string): Promise<void> {
  const url = pdfUrl(filename, true);
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch PDF: ${res.statusText}`);
  const blob = await res.blob();
  const blobUrl = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = blobUrl;
  a.download = filename;          // exact filename the user will see
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  setTimeout(() => URL.revokeObjectURL(blobUrl), 5000);
}
