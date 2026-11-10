export type OutlookConnectionStatus =
  | 'CONNECTED'
  | 'DISCONNECTED'
  | 'REAUTHENTICATION_REQUIRED'
  | 'ERROR'
  | 'CHECKING'
  | 'VERIFICATION_UNAVAILABLE';

export interface OutlookStatusResponse {
  connected: boolean;
  status: OutlookConnectionStatus;
  email?: string;
  displayName?: string;
  accountType?: 'work_school' | 'personal' | 'unknown';
  scopes?: string[];
  lastVerifiedAt?: string;
  connectedAt?: string;
  errorMessage?: string;
  isConfigured: boolean;
  isSandbox?: boolean;
}

export async function getOutlookConfig(): Promise<{ isConfigured: boolean; redirectUri: string }> {
  const res = await fetch('/api/outlook/config');
  if (!res.ok) throw new Error('Failed to fetch Outlook configuration');
  return res.json();
}

export async function initiateOutlookConnect(teacherId: string): Promise<{
  authUrl: string;
  state: string;
  isConfigured: boolean;
  redirectUri: string;
}> {
  const res = await fetch('/api/outlook/connect', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    const errorData = await res.json().catch(() => ({}));
    throw new Error(errorData.error || 'Failed to initiate Outlook connection');
  }
  return res.json();
}

export async function getOutlookStatus(teacherId: string): Promise<OutlookStatusResponse> {
  const res = await fetch(`/api/outlook/status?teacherId=${encodeURIComponent(teacherId)}`);
  if (!res.ok) {
    throw new Error('Failed to fetch Outlook connection status');
  }
  return res.json();
}

export async function verifyOutlookConnection(teacherId: string): Promise<OutlookStatusResponse> {
  const res = await fetch('/api/outlook/verify', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    throw new Error('Live verification request failed');
  }
  return res.json();
}

export async function disconnectOutlook(teacherId: string): Promise<void> {
  const res = await fetch('/api/outlook/disconnect', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ teacherId }),
  });
  if (!res.ok) {
    throw new Error('Failed to disconnect Outlook');
  }
}

export async function logoutDisconnectOutlook(teacherId: string): Promise<void> {
  try {
    await fetch('/api/outlook/logout-disconnect', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ teacherId }),
    });
  } catch (e) {
    // Silent on background logout
  }
}
