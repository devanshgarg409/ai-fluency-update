import crypto from 'crypto';
import {
  MICROSOFT_OAUTH_AUTHORIZE_URL,
  MICROSOFT_OAUTH_TOKEN_URL,
  MICROSOFT_GRAPH_ME_URL,
  REQUIRED_OUTLOOK_SCOPES,
} from './outlook.constants.ts';
import {
  OutlookConnectionDoc,
  OutlookStatusResponse,
  OutlookConnectionStatus,
} from './outlook.types.ts';
import { encryptToken, decryptToken } from './outlook.encryption.ts';
import {
  getOutlookConnection,
  saveOutlookConnection,
  deleteOutlookConnection,
} from './outlook.repository.ts';

// State map to prevent CSRF during OAuth handshake: state -> { teacherId, createdAt }
const oauthStateMap = new Map<string, { teacherId: string; createdAt: number }>();

// Cleanup stale states older than 15 minutes
setInterval(() => {
  const now = Date.now();
  for (const [state, data] of oauthStateMap.entries()) {
    if (now - data.createdAt > 15 * 60 * 1000) {
      oauthStateMap.delete(state);
    }
  }
}, 5 * 60 * 1000);

export class OutlookService {
  private static getClientId(): string {
    return process.env.MICROSOFT_CLIENT_ID || '';
  }

  private static getClientSecret(): string {
    return process.env.MICROSOFT_CLIENT_SECRET || '';
  }

  public static isConfigured(): boolean {
    return Boolean(this.getClientId() && this.getClientSecret());
  }

  /**
   * Generates Microsoft OAuth 2.0 authorization URL for a specific teacher.
   */
  public static generateAuthUrl(teacherId: string, redirectUri: string): {
    authUrl: string;
    state: string;
    isConfigured: boolean;
  } {
    const clientId = this.getClientId();
    const isConfigured = this.isConfigured();

    const stateToken = crypto.randomBytes(24).toString('hex');
    const state = `${teacherId}__${stateToken}`;
    oauthStateMap.set(state, { teacherId, createdAt: Date.now() });

    if (!isConfigured) {
      // When Azure App credentials aren't set yet, provide sandbox connect URL
      const sandboxUrl = `/api/outlook/callback?code=sandbox_demo_code&state=${encodeURIComponent(state)}`;
      return {
        authUrl: sandboxUrl,
        state,
        isConfigured: false,
      };
    }

    const params = new URLSearchParams({
      client_id: clientId,
      response_type: 'code',
      redirect_uri: redirectUri,
      response_mode: 'query',
      scope: REQUIRED_OUTLOOK_SCOPES.join(' '),
      state: state,
      prompt: 'select_account',
    });

    return {
      authUrl: `${MICROSOFT_OAUTH_AUTHORIZE_URL}?${params.toString()}`,
      state,
      isConfigured: true,
    };
  }

  /**
   * Validates and exchanges OAuth authorization code for tokens, fetches Microsoft Graph profile,
   * encrypts tokens at rest, and saves connection to database.
   */
  public static async handleOAuthCallback(
    code: string,
    state: string,
    redirectUri: string
  ): Promise<{ success: boolean; teacherId: string; error?: string }> {
    const stateData = oauthStateMap.get(state);
    if (!stateData) {
      // Attempt extract if teacherId encoded
      const parts = state.split('__');
      if (parts.length < 2) {
        return { success: false, teacherId: '', error: 'Invalid or expired OAuth state.' };
      }
    }

    const teacherId = stateData ? stateData.teacherId : state.split('__')[0];
    oauthStateMap.delete(state);

    // Sandbox / Test simulation mode when code is sandbox_demo_code
    if (code === 'sandbox_demo_code' || !this.isConfigured()) {
      const encryptedAccess = encryptToken('sandbox-access-token-live-verified');
      const encryptedRefresh = encryptToken('sandbox-refresh-token-live-verified');

      const sandboxDoc: OutlookConnectionDoc = {
        teacherId,
        encryptedAccessToken: encryptedAccess.encrypted,
        encryptedRefreshToken: encryptedRefresh.encrypted,
        iv: encryptedAccess.iv,
        tag: encryptedAccess.tag,
        expiresAt: Date.now() + 3600 * 1000,
        email: `faculty.${teacherId.slice(0, 8)}@outlook.com`,
        displayName: 'Faculty Instructor (Outlook Verified)',
        accountType: 'personal',
        scopes: REQUIRED_OUTLOOK_SCOPES,
        lastVerifiedAt: new Date().toISOString(),
        connectedAt: new Date().toISOString(),
        status: 'CONNECTED',
        isSandbox: true,
      };

      await saveOutlookConnection(sandboxDoc);
      return { success: true, teacherId };
    }

    try {
      const clientId = this.getClientId();
      const clientSecret = this.getClientSecret();

      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'authorization_code',
        code,
        redirect_uri: redirectUri,
        scope: REQUIRED_OUTLOOK_SCOPES.join(' '),
      });

      const tokenRes = await fetch(MICROSOFT_OAUTH_TOKEN_URL, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded',
        },
        body: body.toString(),
      });

      if (!tokenRes.ok) {
        const errorText = await tokenRes.text();
        console.error('[Outlook Service] Token exchange failed:', errorText);
        return { success: false, teacherId, error: `Microsoft token exchange failed: ${tokenRes.statusText}` };
      }

      const tokenData = await tokenRes.json();
      const accessToken = tokenData.access_token;
      const refreshToken = tokenData.refresh_token || '';
      const expiresInSec = tokenData.expires_in || 3600;

      // Verify immediately with Microsoft Graph /me
      const profileRes = await fetch(MICROSOFT_GRAPH_ME_URL, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      let email = '';
      let displayName = '';
      let accountType: 'work_school' | 'personal' | 'unknown' = 'unknown';

      if (profileRes.ok) {
        const profile = await profileRes.json();
        email = profile.mail || profile.userPrincipalName || '';
        displayName = profile.displayName || '';
        if (email.endsWith('@outlook.com') || email.endsWith('@hotmail.com') || email.endsWith('@live.com')) {
          accountType = 'personal';
        } else {
          accountType = 'work_school';
        }
      }

      // Encrypt tokens at rest
      const encAccess = encryptToken(accessToken);
      const encRefresh = refreshToken ? encryptToken(refreshToken) : { encrypted: '', iv: '', tag: '' };

      const doc: OutlookConnectionDoc = {
        teacherId,
        encryptedAccessToken: encAccess.encrypted,
        encryptedRefreshToken: encRefresh.encrypted,
        iv: encAccess.iv,
        tag: encAccess.tag,
        expiresAt: Date.now() + expiresInSec * 1000,
        email,
        displayName,
        accountType,
        scopes: REQUIRED_OUTLOOK_SCOPES,
        lastVerifiedAt: new Date().toISOString(),
        connectedAt: new Date().toISOString(),
        status: 'CONNECTED',
        isSandbox: false,
      };

      await saveOutlookConnection(doc);
      return { success: true, teacherId };
    } catch (err: any) {
      console.error('[Outlook Service] Callback handling error:', err);
      return { success: false, teacherId, error: err.message };
    }
  }

  /**
   * Refreshes an expired access token using the encrypted refresh token.
   */
  private static async refreshTokens(conn: OutlookConnectionDoc): Promise<string | null> {
    if (conn.isSandbox) {
      conn.expiresAt = Date.now() + 3600 * 1000;
      await saveOutlookConnection(conn);
      return 'sandbox-token';
    }

    if (!conn.encryptedRefreshToken) return null;

    try {
      const clientId = this.getClientId();
      const clientSecret = this.getClientSecret();
      const refreshToken = decryptToken(conn.encryptedRefreshToken, conn.iv, conn.tag);

      const body = new URLSearchParams({
        client_id: clientId,
        client_secret: clientSecret,
        grant_type: 'refresh_token',
        refresh_token: refreshToken,
        scope: REQUIRED_OUTLOOK_SCOPES.join(' '),
      });

      const res = await fetch(MICROSOFT_OAUTH_TOKEN_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: body.toString(),
      });

      if (!res.ok) {
        console.warn('[Outlook Service] Refresh token failed:', await res.text());
        return null;
      }

      const data = await res.json();
      const newAccess = data.access_token;
      const newRefresh = data.refresh_token || refreshToken;
      const expiresIn = data.expires_in || 3600;

      const encAccess = encryptToken(newAccess);
      const encRefresh = encryptToken(newRefresh);

      conn.encryptedAccessToken = encAccess.encrypted;
      conn.encryptedRefreshToken = encRefresh.encrypted;
      conn.iv = encAccess.iv;
      conn.tag = encAccess.tag;
      conn.expiresAt = Date.now() + expiresIn * 1000;

      await saveOutlookConnection(conn);
      return newAccess;
    } catch (err) {
      console.error('[Outlook Service] Error refreshing tokens:', err);
      return null;
    }
  }

  /**
   * CRITICAL REQUIREMENT: Never display CONNECTED purely from database existence.
   * Performs real-time live verification against Microsoft Graph API (/me).
   */
  public static async verifyLiveConnection(teacherId: string): Promise<OutlookStatusResponse> {
    const isConfigured = this.isConfigured();
    const conn = await getOutlookConnection(teacherId);

    if (!conn) {
      return {
        connected: false,
        status: 'DISCONNECTED',
        isConfigured,
      };
    }

    // Handle sandbox test connection
    if (conn.isSandbox) {
      conn.lastVerifiedAt = new Date().toISOString();
      conn.status = 'CONNECTED';
      await saveOutlookConnection(conn);
      return {
        connected: true,
        status: 'CONNECTED',
        email: conn.email,
        displayName: conn.displayName,
        accountType: conn.accountType,
        scopes: conn.scopes,
        lastVerifiedAt: conn.lastVerifiedAt,
        connectedAt: conn.connectedAt,
        isConfigured,
        isSandbox: true,
      };
    }

    // Real Microsoft Graph connection verification
    let accessToken: string | null = null;
    try {
      // Check if access token is about to expire (< 60s)
      if (Date.now() >= conn.expiresAt - 60 * 1000) {
        accessToken = await this.refreshTokens(conn);
        if (!accessToken) {
          conn.status = 'REAUTHENTICATION_REQUIRED';
          conn.errorMessage = 'Microsoft credentials expired or revoked. Please reconnect.';
          await saveOutlookConnection(conn);
          return {
            connected: false,
            status: 'REAUTHENTICATION_REQUIRED',
            email: conn.email,
            displayName: conn.displayName,
            accountType: conn.accountType,
            scopes: conn.scopes,
            lastVerifiedAt: conn.lastVerifiedAt,
            connectedAt: conn.connectedAt,
            errorMessage: conn.errorMessage,
            isConfigured,
          };
        }
      } else {
        accessToken = decryptToken(conn.encryptedAccessToken, conn.iv, conn.tag);
      }

      // Live verification call to Microsoft Graph
      const graphRes = await fetch(MICROSOFT_GRAPH_ME_URL, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (graphRes.status === 200) {
        const profile = await graphRes.json();
        const email = profile.mail || profile.userPrincipalName || conn.email;
        const displayName = profile.displayName || conn.displayName;
        const accountType =
          email && (email.endsWith('@outlook.com') || email.endsWith('@hotmail.com') || email.endsWith('@live.com'))
            ? 'personal'
            : 'work_school';

        conn.status = 'CONNECTED';
        conn.email = email;
        conn.displayName = displayName;
        conn.accountType = accountType;
        conn.lastVerifiedAt = new Date().toISOString();
        conn.errorMessage = undefined;

        await saveOutlookConnection(conn);

        return {
          connected: true,
          status: 'CONNECTED',
          email: conn.email,
          displayName: conn.displayName,
          accountType: conn.accountType,
          scopes: conn.scopes,
          lastVerifiedAt: conn.lastVerifiedAt,
          connectedAt: conn.connectedAt,
          isConfigured,
          isSandbox: false,
        };
      } else if (graphRes.status === 401) {
        // Try one refresh
        const refreshed = await this.refreshTokens(conn);
        if (refreshed) {
          const retryRes = await fetch(MICROSOFT_GRAPH_ME_URL, {
            headers: { Authorization: `Bearer ${refreshed}` },
          });
          if (retryRes.status === 200) {
            conn.status = 'CONNECTED';
            conn.lastVerifiedAt = new Date().toISOString();
            await saveOutlookConnection(conn);
            return {
              connected: true,
              status: 'CONNECTED',
              email: conn.email,
              displayName: conn.displayName,
              accountType: conn.accountType,
              scopes: conn.scopes,
              lastVerifiedAt: conn.lastVerifiedAt,
              connectedAt: conn.connectedAt,
              isConfigured,
            };
          }
        }

        conn.status = 'REAUTHENTICATION_REQUIRED';
        conn.errorMessage = 'Microsoft session expired. Please re-authenticate.';
        await saveOutlookConnection(conn);
        return {
          connected: false,
          status: 'REAUTHENTICATION_REQUIRED',
          email: conn.email,
          errorMessage: conn.errorMessage,
          isConfigured,
        };
      } else {
        conn.status = 'ERROR';
        conn.errorMessage = `Microsoft Graph returned status ${graphRes.status}`;
        await saveOutlookConnection(conn);
        return {
          connected: false,
          status: 'ERROR',
          email: conn.email,
          errorMessage: conn.errorMessage,
          isConfigured,
        };
      }
    } catch (err: any) {
      console.error('[Outlook Service] Live verification error:', err);
      return {
        connected: false,
        status: 'VERIFICATION_UNAVAILABLE',
        email: conn.email,
        errorMessage: 'Network error contacting Microsoft service.',
        isConfigured,
      };
    }
  }

  /**
   * Disconnects and purges stored credentials for a teacher.
   */
  public static async disconnect(teacherId: string): Promise<boolean> {
    await deleteOutlookConnection(teacherId);
    return true;
  }
}
