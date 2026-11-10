import React, { useState, useEffect, useCallback, useRef } from 'react';
import { useAuth } from '../../context/AuthContext.tsx';
import {
  Mail,
  CheckCircle2,
  AlertTriangle,
  XCircle,
  RefreshCw,
  Unlink,
  ExternalLink,
  Shield,
  Key,
  Clock,
  UserCheck,
  AlertCircle,
  Sparkles,
} from 'lucide-react';
import {
  OutlookConnectionStatus,
  OutlookStatusResponse,
  getOutlookStatus,
  verifyOutlookConnection,
  initiateOutlookConnect,
  disconnectOutlook,
} from '../../services/outlookApi.ts';

export const OutlookIntegrationModule: React.FC = () => {
  const { user } = useAuth();
  const [status, setStatus] = useState<OutlookConnectionStatus>('CHECKING');
  const [connectionData, setConnectionData] = useState<OutlookStatusResponse | null>(null);
  const [isVerifying, setIsVerifying] = useState(false);
  const [isConnecting, setIsConnecting] = useState(false);
  const [isDisconnecting, setIsDisconnecting] = useState(false);
  const [showDisconnectModal, setShowDisconnectModal] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string | null>(null);
  const [successNotice, setSuccessNotice] = useState<string | null>(null);

  const popupRef = useRef<Window | null>(null);
  const popupCheckIntervalRef = useRef<any>(null);

  const teacherId = user?.id || '';

  // Perform live verification against Microsoft Graph API
  const performCheck = useCallback(async (showIndicator = false) => {
    if (!teacherId) return;
    if (showIndicator) setIsVerifying(true);
    setErrorMessage(null);

    try {
      const data = await getOutlookStatus(teacherId);
      setConnectionData(data);
      setStatus(data.status);
    } catch (err: any) {
      console.error('Failed to verify Outlook connection:', err);
      setStatus('ERROR');
      setErrorMessage(err.message || 'Unable to contact verification service.');
    } finally {
      if (showIndicator) setIsVerifying(false);
    }
  }, [teacherId]);

  // Initial load check & 5-minute periodic verification
  useEffect(() => {
    performCheck(false);

    // Periodic live verification every 5 minutes while viewing tab
    const interval = setInterval(() => {
      performCheck(false);
    }, 5 * 60 * 1000);

    return () => clearInterval(interval);
  }, [performCheck]);

  // Listen for OAuth popup completion postMessage
  useEffect(() => {
    const handleMessage = async (event: MessageEvent) => {
      if (event.data?.type === 'OUTLOOK_AUTH_SUCCESS') {
        if (popupRef.current && !popupRef.current.closed) {
          popupRef.current.close();
        }
        if (popupCheckIntervalRef.current) {
          clearInterval(popupCheckIntervalRef.current);
        }
        setIsConnecting(false);
        setSuccessNotice('Microsoft Outlook successfully connected and verified live!');
        setTimeout(() => setSuccessNotice(null), 6000);

        // Immediate live verification after OAuth handshake
        setStatus('CHECKING');
        await performCheck(true);
      } else if (event.data?.type === 'OUTLOOK_AUTH_ERROR') {
        setIsConnecting(false);
        setErrorMessage(decodeURIComponent(event.data.error || 'Authentication was cancelled or failed.'));
        if (popupCheckIntervalRef.current) {
          clearInterval(popupCheckIntervalRef.current);
        }
      }
    };

    window.addEventListener('message', handleMessage);
    return () => {
      window.removeEventListener('message', handleMessage);
      if (popupCheckIntervalRef.current) {
        clearInterval(popupCheckIntervalRef.current);
      }
    };
  }, [performCheck]);

  // Handle Connect Button Click
  const handleConnect = async () => {
    if (!teacherId) return;
    setIsConnecting(true);
    setErrorMessage(null);
    setSuccessNotice(null);

    try {
      const authData = await initiateOutlookConnect(teacherId);

      // Open OAuth popup window
      const width = 600;
      const height = 720;
      const left = window.screenX + (window.outerWidth - width) / 2;
      const top = window.screenY + (window.outerHeight - height) / 2;

      const popup = window.open(
        authData.authUrl,
        'outlook_oauth_popup',
        `width=${width},height=${height},left=${left},top=${top},status=no,resizable=yes,scrollbars=yes`
      );

      popupRef.current = popup;

      if (!popup || popup.closed || typeof popup.closed === 'undefined') {
        setIsConnecting(false);
        setErrorMessage('Popup was blocked by your browser. Please allow popups for this site and try again.');
        return;
      }

      // Check if user manually closed popup without finishing
      popupCheckIntervalRef.current = setInterval(() => {
        if (popup.closed) {
          clearInterval(popupCheckIntervalRef.current);
          setIsConnecting(false);
        }
      }, 1000);
    } catch (err: any) {
      setIsConnecting(false);
      setErrorMessage(err.message || 'Failed to initiate Microsoft connection');
    }
  };

  // Handle Manual Force Verify
  const handleForceVerify = async () => {
    if (!teacherId) return;
    setIsVerifying(true);
    setErrorMessage(null);

    try {
      const result = await verifyOutlookConnection(teacherId);
      setConnectionData(result);
      setStatus(result.status);
      if (result.status === 'CONNECTED') {
        setSuccessNotice('Connection verified live with Microsoft Graph.');
        setTimeout(() => setSuccessNotice(null), 4000);
      }
    } catch (err: any) {
      setErrorMessage(err.message || 'Verification failed');
      setStatus('ERROR');
    } finally {
      setIsVerifying(false);
    }
  };

  // Handle Disconnect
  const handleConfirmDisconnect = async () => {
    if (!teacherId) return;
    setIsDisconnecting(true);
    try {
      await disconnectOutlook(teacherId);
      setStatus('DISCONNECTED');
      setConnectionData(null);
      setShowDisconnectModal(false);
      setSuccessNotice('Outlook account successfully disconnected and tokens wiped.');
      setTimeout(() => setSuccessNotice(null), 4000);
    } catch (err: any) {
      setErrorMessage(err.message || 'Failed to disconnect account');
    } finally {
      setIsDisconnecting(false);
    }
  };

  // Helper for human-readable relative time
  const formatTimeAgo = (isoString?: string) => {
    if (!isoString) return 'Never';
    const date = new Date(isoString);
    const diffSec = Math.floor((Date.now() - date.getTime()) / 1000);
    if (diffSec < 15) return 'Just now';
    if (diffSec < 60) return `${diffSec} seconds ago`;
    const diffMin = Math.floor(diffSec / 60);
    if (diffMin < 60) return `${diffMin} min${diffMin > 1 ? 's' : ''} ago`;
    const diffHour = Math.floor(diffMin / 60);
    return `${diffHour} hour${diffHour > 1 ? 's' : ''} ago`;
  };

  return (
    <div className="space-y-6">
      {/* Module Overview Header */}
      <div className="bg-gradient-to-r from-slate-900 via-indigo-950/40 to-slate-900 border border-slate-800 rounded-2xl p-6">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
          <div className="flex items-center space-x-3.5">
            <div className="w-12 h-12 rounded-2xl bg-sky-950/80 border border-sky-800/60 text-sky-400 flex items-center justify-center shadow-lg shadow-sky-950/40">
              <Mail className="w-6 h-6" />
            </div>
            <div>
              <div className="flex items-center space-x-2.5">
                <h3 className="text-lg font-bold text-white tracking-tight">
                  Microsoft Outlook Integration
                </h3>
                {connectionData?.isSandbox && (
                  <span className="text-[10px] px-2 py-0.5 rounded-full bg-indigo-950 text-indigo-300 border border-indigo-700/60 font-semibold font-mono">
                    Sandbox Verified
                  </span>
                )}
              </div>
              <p className="text-xs text-slate-400 mt-0.5">
                Individual per-teacher OAuth 2.0 integration for academic correspondence and alerts.
              </p>
            </div>
          </div>

          {/* Live Status Badge */}
          <div className="flex items-center">
            {status === 'CHECKING' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-slate-900 border border-indigo-500/40 text-indigo-300 text-xs font-semibold shadow-inner">
                <RefreshCw className="w-3.5 h-3.5 animate-spin text-indigo-400" />
                <span>Verifying Live Status...</span>
              </div>
            )}

            {status === 'CONNECTED' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-emerald-950/70 border border-emerald-600/60 text-emerald-300 text-xs font-semibold shadow-inner">
                <span className="w-2 h-2 rounded-full bg-emerald-400 animate-pulse" />
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                <span>CONNECTED (LIVE)</span>
              </div>
            )}

            {status === 'DISCONNECTED' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-slate-900 border border-slate-800 text-slate-400 text-xs font-semibold">
                <span className="w-2 h-2 rounded-full bg-slate-500" />
                <span>DISCONNECTED</span>
              </div>
            )}

            {status === 'REAUTHENTICATION_REQUIRED' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-amber-950/80 border border-amber-600/60 text-amber-300 text-xs font-semibold">
                <AlertTriangle className="w-3.5 h-3.5 text-amber-400" />
                <span>REAUTHENTICATION REQUIRED</span>
              </div>
            )}

            {status === 'ERROR' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-rose-950/80 border border-rose-600/60 text-rose-300 text-xs font-semibold">
                <XCircle className="w-3.5 h-3.5 text-rose-400" />
                <span>CONNECTION ERROR</span>
              </div>
            )}

            {status === 'VERIFICATION_UNAVAILABLE' && (
              <div className="flex items-center space-x-2 px-3.5 py-1.5 rounded-xl bg-slate-900 border border-slate-700 text-slate-300 text-xs font-semibold">
                <AlertCircle className="w-3.5 h-3.5 text-slate-400" />
                <span>VERIFICATION UNAVAILABLE</span>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* Notifications */}
      {successNotice && (
        <div className="p-3.5 rounded-xl bg-emerald-950/50 border border-emerald-800/60 flex items-center space-x-2.5 text-xs text-emerald-300">
          <CheckCircle2 className="w-4 h-4 text-emerald-400 shrink-0" />
          <span>{successNotice}</span>
        </div>
      )}

      {errorMessage && (
        <div className="p-3.5 rounded-xl bg-rose-950/50 border border-rose-800/60 flex items-center space-x-2.5 text-xs text-rose-300">
          <AlertCircle className="w-4 h-4 text-rose-400 shrink-0" />
          <span>{errorMessage}</span>
        </div>
      )}

      {/* Main Section */}
      {status === 'CONNECTED' ? (
        <div className="space-y-6">
          {/* Connection Details Card */}
          <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-6 space-y-6 shadow-xl">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between pb-5 border-b border-slate-800 gap-4">
              <div>
                <h4 className="text-sm font-bold text-slate-100 flex items-center space-x-2">
                  <UserCheck className="w-4 h-4 text-emerald-400" />
                  <span>Active Microsoft Session</span>
                </h4>
                <p className="text-xs text-slate-400 mt-0.5">
                  Connected to Microsoft Graph API with individual teacher delegated permissions.
                </p>
              </div>

              {/* Action Buttons */}
              <div className="flex items-center space-x-2.5">
                <button
                  onClick={handleForceVerify}
                  disabled={isVerifying}
                  className="px-3.5 py-2 rounded-xl bg-slate-800 hover:bg-slate-700/80 text-slate-200 border border-slate-700 text-xs font-semibold flex items-center space-x-2 transition-all cursor-pointer disabled:opacity-50"
                  title="Verify live connection with Microsoft Graph right now"
                >
                  <RefreshCw className={`w-3.5 h-3.5 ${isVerifying ? 'animate-spin text-indigo-400' : ''}`} />
                  <span>{isVerifying ? 'Verifying...' : 'Verify Live Now'}</span>
                </button>

                <button
                  onClick={() => setShowDisconnectModal(true)}
                  className="px-3.5 py-2 rounded-xl bg-rose-950/60 hover:bg-rose-900/80 text-rose-300 border border-rose-800/60 text-xs font-semibold flex items-center space-x-2 transition-all cursor-pointer"
                >
                  <Unlink className="w-3.5 h-3.5" />
                  <span>Disconnect</span>
                </button>
              </div>
            </div>

            {/* Profile & Connection Metadata Grid */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4 text-xs">
              <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 space-y-1">
                <span className="text-slate-400 block text-[11px] uppercase tracking-wider font-semibold">
                  Connected Email
                </span>
                <span className="text-slate-100 font-mono font-bold text-sm truncate block" title={connectionData?.email}>
                  {connectionData?.email || 'Authenticated User'}
                </span>
                <span className="text-[11px] text-slate-400">
                  {connectionData?.displayName || user?.name}
                </span>
              </div>

              <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 space-y-1">
                <span className="text-slate-400 block text-[11px] uppercase tracking-wider font-semibold">
                  Account Type
                </span>
                <span className="text-indigo-300 font-semibold text-sm block">
                  {connectionData?.accountType === 'work_school'
                    ? 'Work / School (Entra ID)'
                    : 'Personal Microsoft Account'}
                </span>
                <span className="text-[11px] text-slate-400">
                  Multitenant / Common Authority
                </span>
              </div>

              <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 space-y-1">
                <span className="text-slate-400 block text-[11px] uppercase tracking-wider font-semibold flex items-center space-x-1">
                  <Clock className="w-3 h-3 text-slate-400" />
                  <span>Last Live Verification</span>
                </span>
                <span className="text-emerald-400 font-semibold text-sm block">
                  {formatTimeAgo(connectionData?.lastVerifiedAt)}
                </span>
                <span className="text-[10px] text-slate-400 font-mono">
                  {connectionData?.lastVerifiedAt
                    ? new Date(connectionData.lastVerifiedAt).toLocaleTimeString()
                    : '—'}
                </span>
              </div>

              <div className="bg-slate-950/60 p-4 rounded-xl border border-slate-800 space-y-1">
                <span className="text-slate-400 block text-[11px] uppercase tracking-wider font-semibold flex items-center space-x-1">
                  <Shield className="w-3 h-3 text-slate-400" />
                  <span>Security & Storage</span>
                </span>
                <span className="text-slate-200 font-mono font-semibold text-xs block">
                  AES-256-GCM Encrypted
                </span>
                <span className="text-[11px] text-slate-400">
                  Restricted to Teacher ID: {user?.employeeId || user?.username}
                </span>
              </div>
            </div>

            {/* Delegated Permissions Details */}
            <div className="bg-slate-950/50 p-4 rounded-xl border border-slate-800 space-y-3">
              <h5 className="text-xs font-semibold text-slate-300 flex items-center space-x-2">
                <Key className="w-3.5 h-3.5 text-indigo-400" />
                <span>Active Delegated Microsoft Graph Scopes</span>
              </h5>
              <div className="flex flex-wrap gap-2">
                {(connectionData?.scopes || [
                  'User.Read',
                  'Mail.Read',
                  'Mail.ReadWrite',
                  'Mail.Send',
                  'openid',
                  'profile',
                  'email',
                  'offline_access',
                ]).map((scope) => (
                  <span
                    key={scope}
                    className="px-2.5 py-1 rounded-lg bg-slate-900 border border-slate-800 text-[11px] font-mono text-indigo-300 flex items-center space-x-1"
                  >
                    <CheckCircle2 className="w-2.5 h-2.5 text-emerald-400" />
                    <span>{scope}</span>
                  </span>
                ))}
              </div>
            </div>
          </div>
        </div>
      ) : (
        /* Disconnected / Re-auth Required State */
        <div className="bg-slate-900/60 border border-slate-800 rounded-2xl p-8 space-y-6 shadow-xl max-w-3xl">
          <div className="space-y-2">
            <h4 className="text-base font-bold text-slate-100 flex items-center space-x-2">
              <Mail className="w-5 h-5 text-indigo-400" />
              <span>Connect Your Microsoft Outlook Account</span>
            </h4>
            <p className="text-xs text-slate-400 leading-relaxed">
              Link your personal Microsoft account or institutional Work/School Microsoft 365 mailbox to EduCore.
              Each faculty member connects their own individual account. Your tokens are securely encrypted at rest with AES-256-GCM and never shared with other teachers or administrators.
            </p>
          </div>

          {status === 'REAUTHENTICATION_REQUIRED' && (
            <div className="p-4 rounded-xl bg-amber-950/50 border border-amber-800/60 text-xs text-amber-300 space-y-1">
              <div className="font-semibold flex items-center space-x-1.5">
                <AlertTriangle className="w-4 h-4 text-amber-400" />
                <span>Session Expired or Revoked</span>
              </div>
              <p className="text-amber-300/80">
                Your Microsoft access has expired or was revoked. Please click the button below to sign in again.
              </p>
            </div>
          )}

          {/* Connect Action Box */}
          <div className="p-5 rounded-xl bg-slate-950 border border-slate-800 space-y-4">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
              <div className="space-y-1">
                <span className="text-xs font-semibold text-slate-200 block">
                  OAuth 2.0 Authorization Code Flow
                </span>
                <span className="text-[11px] text-slate-400 block">
                  Authorizes delegated permissions via Microsoft identity platform.
                </span>
              </div>

              <button
                onClick={handleConnect}
                disabled={isConnecting}
                className="px-5 py-2.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold flex items-center justify-center space-x-2 shadow-lg shadow-indigo-600/30 transition-all cursor-pointer disabled:opacity-50"
              >
                {isConnecting ? (
                  <>
                    <RefreshCw className="w-4 h-4 animate-spin text-white" />
                    <span>Connecting in Popup...</span>
                  </>
                ) : (
                  <>
                    <ExternalLink className="w-4 h-4 text-white" />
                    <span>Connect Microsoft Outlook</span>
                  </>
                )}
              </button>
            </div>

            <div className="pt-3 border-t border-slate-900 grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] text-slate-400">
              <div className="flex items-center space-x-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                <span>Supports both Work/School & Personal Outlook.com</span>
              </div>
              <div className="flex items-center space-x-1.5">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-400" />
                <span>Tokens automatically purged upon logout</span>
              </div>
            </div>
          </div>

          {/* Scope Overview Info */}
          <div className="space-y-2">
            <h5 className="text-xs font-semibold text-slate-300">Requested Delegated Permissions</h5>
            <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-[11px] text-slate-400">
              <div className="bg-slate-950/40 p-2.5 rounded-lg border border-slate-800/80">
                <strong className="text-slate-200 block">User.Read & Mail.Read</strong>
                Access your teacher profile name and view academic communications.
              </div>
              <div className="bg-slate-950/40 p-2.5 rounded-lg border border-slate-800/80">
                <strong className="text-slate-200 block">Mail.Send & Mail.ReadWrite</strong>
                Transmit course notices and manage academic message drafts.
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Disconnect Confirmation Modal */}
      {showDisconnectModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm">
          <div className="bg-slate-900 border border-slate-800 rounded-2xl max-w-md w-full p-6 space-y-4 shadow-2xl">
            <div className="w-12 h-12 rounded-2xl bg-rose-950/80 border border-rose-800/60 text-rose-400 flex items-center justify-center">
              <AlertTriangle className="w-6 h-6" />
            </div>

            <div className="space-y-1">
              <h3 className="text-base font-bold text-white">Disconnect Microsoft Outlook?</h3>
              <p className="text-xs text-slate-400 leading-relaxed">
                This will immediately purge all stored access tokens and refresh credentials from the persistent database for your faculty account ({user?.name}). Your connection status will revert to DISCONNECTED.
              </p>
            </div>

            <div className="flex items-center justify-end space-x-3 pt-3">
              <button
                onClick={() => setShowDisconnectModal(false)}
                disabled={isDisconnecting}
                className="px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 text-xs font-semibold cursor-pointer"
              >
                Cancel
              </button>
              <button
                onClick={handleConfirmDisconnect}
                disabled={isDisconnecting}
                className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white text-xs font-semibold flex items-center space-x-1.5 cursor-pointer shadow-lg shadow-rose-600/30"
              >
                {isDisconnecting && <RefreshCw className="w-3.5 h-3.5 animate-spin" />}
                <span>{isDisconnecting ? 'Disconnecting...' : 'Yes, Disconnect'}</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
