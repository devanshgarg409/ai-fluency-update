import { Router, Request, Response } from 'express';
import { OutlookService } from './outlook.service.ts';

const router = Router();

function getRedirectUri(req: Request): string {
  if (process.env.MICROSOFT_REDIRECT_URI) {
    return process.env.MICROSOFT_REDIRECT_URI.trim();
  }
  const protocol = req.headers['x-forwarded-proto'] || (req.secure ? 'https' : 'http');
  const host = req.headers['x-forwarded-host'] || req.headers.host;
  if (host && !host.includes('localhost') && !host.includes('127.0.0.1')) {
    return `${protocol}://${host}/api/outlook/callback`;
  }
  return 'https://ais-dev-k2izbmn7pw5mntljrz2ps5-490136516285.asia-east1.run.app/api/outlook/callback';
}

// GET /api/outlook/config - Configuration check
router.get('/config', (req: Request, res: Response) => {
  res.json({
    isConfigured: OutlookService.isConfigured(),
    redirectUri: getRedirectUri(req),
  });
});

// POST /api/outlook/connect - Start OAuth flow
router.post('/connect', async (req: Request, res: Response) => {
  try {
    const { teacherId } = req.body;
    if (!teacherId) {
      return res.status(400).json({ error: 'Teacher ID is required to initiate Outlook connection' });
    }

    const redirectUri = getRedirectUri(req);
    const authData = OutlookService.generateAuthUrl(teacherId, redirectUri);

    return res.json({
      authUrl: authData.authUrl,
      state: authData.state,
      isConfigured: authData.isConfigured,
      redirectUri,
    });
  } catch (err: any) {
    console.error('[Outlook Routes] Connect initiation error:', err);
    return res.status(500).json({ error: 'Failed to initiate Outlook connection: ' + err.message });
  }
});

// GET /api/outlook/callback - OAuth redirect handler (renders popup postMessage closer)
router.get('/callback', async (req: Request, res: Response) => {
  const code = (req.query.code as string) || '';
  const state = (req.query.state as string) || '';
  const error = (req.query.error as string) || '';
  const errorDescription = (req.query.error_description as string) || '';

  if (error) {
    const safeDesc = errorDescription.replace(/</g, '&lt;').replace(/>/g, '&gt;');
    return res.send(`
      <!DOCTYPE html>
      <html>
        <head><title>Authentication Cancelled</title></head>
        <body style="background:#020617;color:#f8fafc;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;padding:20px;">
          <div style="text-align:center;max-width:400px;background:#0f172a;border:1px solid #334155;border-radius:12px;padding:24px;">
            <h3 style="color:#ef4444;margin-top:0;">Authentication Cancelled</h3>
            <p style="font-size:13px;color:#94a3b8;">${safeDesc || 'Access was declined or expired.'}</p>
            <button onclick="window.close()" style="background:#334155;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;margin-top:12px;">Close Window</button>
          </div>
          <script>
            try {
              if (window.opener) {
                window.opener.postMessage({ type: 'OUTLOOK_AUTH_ERROR', error: '${encodeURIComponent(error)}' }, '*');
              }
            } catch(e) {}
            setTimeout(() => window.close(), 3000);
          </script>
        </body>
      </html>
    `);
  }

  if (!code) {
    return res.status(400).send('Missing authorization code');
  }

  const redirectUri = getRedirectUri(req);
  const result = await OutlookService.handleOAuthCallback(code, state, redirectUri);

  if (!result.success) {
    return res.send(`
      <!DOCTYPE html>
      <html>
        <head><title>Connection Failed</title></head>
        <body style="background:#020617;color:#f8fafc;font-family:sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;padding:20px;">
          <div style="text-align:center;max-width:400px;background:#0f172a;border:1px solid #ef4444;border-radius:12px;padding:24px;">
            <h3 style="color:#ef4444;margin-top:0;">Connection Failed</h3>
            <p style="font-size:13px;color:#94a3b8;">${result.error || 'Failed to complete Microsoft handshake.'}</p>
            <button onclick="window.close()" style="background:#ef4444;color:#fff;border:none;padding:8px 16px;border-radius:6px;cursor:pointer;margin-top:12px;">Close</button>
          </div>
          <script>
            try {
              if (window.opener) {
                window.opener.postMessage({ type: 'OUTLOOK_AUTH_ERROR', error: '${encodeURIComponent(result.error || '')}' }, '*');
              }
            } catch(e) {}
          </script>
        </body>
      </html>
    `);
  }

  // Success response
  res.send(`
    <!DOCTYPE html>
    <html>
      <head><title>Authentication Complete</title></head>
      <body style="background:#020617;color:#f8fafc;font-family:system-ui,-apple-system,sans-serif;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;padding:20px;">
        <div style="text-align:center;max-width:380px;background:#0f172a;border:1px solid #1e293b;border-radius:16px;padding:32px;box-shadow:0 20px 25px -5px rgba(0,0,0,0.5);">
          <div style="width:48px;height:48px;border-radius:50%;background:#065f46;color:#34d399;display:flex;align-items:center;justify-content:center;margin:0 auto 16px;font-size:24px;">✓</div>
          <h2 style="font-size:18px;font-weight:700;margin:0 0 8px;color:#f8fafc;">Outlook Connected</h2>
          <p style="font-size:13px;color:#94a3b8;margin:0 0 20px;">Your Microsoft Outlook account has been securely authenticated. This window will close automatically.</p>
          <div style="font-size:11px;color:#64748b;">Closing in 1 second...</div>
        </div>
        <script>
          try {
            if (window.opener) {
              window.opener.postMessage({ type: 'OUTLOOK_AUTH_SUCCESS', teacherId: '${result.teacherId}' }, '*');
            }
          } catch(e) {}
          setTimeout(() => {
            window.close();
          }, 1200);
        </script>
      </body>
    </html>
  `);
});

// GET /api/outlook/status - Live check connection status
router.get('/status', async (req: Request, res: Response) => {
  try {
    const teacherId = (req.query.teacherId as string) || '';
    if (!teacherId) {
      return res.status(400).json({ error: 'teacherId parameter is required' });
    }

    const status = await OutlookService.verifyLiveConnection(teacherId);
    return res.json(status);
  } catch (err: any) {
    console.error('[Outlook Routes] Status check error:', err);
    return res.status(500).json({ error: 'Failed to verify Outlook connection: ' + err.message });
  }
});

// POST /api/outlook/verify - Force real-time live verification
router.post('/verify', async (req: Request, res: Response) => {
  try {
    const { teacherId } = req.body;
    if (!teacherId) {
      return res.status(400).json({ error: 'teacherId is required' });
    }

    const status = await OutlookService.verifyLiveConnection(teacherId);
    return res.json(status);
  } catch (err: any) {
    console.error('[Outlook Routes] Force verify error:', err);
    return res.status(500).json({ error: 'Failed to verify connection: ' + err.message });
  }
});

// POST /api/outlook/disconnect - Manual disconnect
router.post('/disconnect', async (req: Request, res: Response) => {
  try {
    const { teacherId } = req.body;
    if (!teacherId) {
      return res.status(400).json({ error: 'teacherId is required' });
    }

    await OutlookService.disconnect(teacherId);
    return res.json({
      success: true,
      status: 'DISCONNECTED',
      message: 'Outlook account disconnected and credentials wiped.',
    });
  } catch (err: any) {
    console.error('[Outlook Routes] Disconnect error:', err);
    return res.status(500).json({ error: 'Failed to disconnect Outlook: ' + err.message });
  }
});

// POST /api/outlook/logout-disconnect - Disconnect on teacher logout
router.post('/logout-disconnect', async (req: Request, res: Response) => {
  try {
    const teacherId = (req.body?.teacherId as string) || (req.query?.teacherId as string) || '';
    if (teacherId) {
      await OutlookService.disconnect(teacherId);
    }
    return res.json({ success: true, status: 'DISCONNECTED' });
  } catch (err) {
    return res.json({ success: true, status: 'DISCONNECTED' });
  }
});

export default router;
