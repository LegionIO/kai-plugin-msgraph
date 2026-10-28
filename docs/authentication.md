# Automatic sign-in recovery

In **Teams & Outlook settings → Account**, **Automatically sign in again** controls
whether the plugin starts sign-in when Microsoft rejects an existing session.
It is on by default, including for existing configurations that do not have the
setting. The stored preference is `preferences.autoReconnect` (boolean).

The plugin first tries to refresh access tokens silently. If Microsoft reports
an expired or revoked session, or requires interaction, login, or consent, the
plugin attempts sign-in once. This also applies when restoring a saved session
at startup and recovering from a Graph 401 or a secondary Teams/Outlook token
failure. Concurrent requests share one sign-in window. Saved credentials are
used when available; Microsoft may still require MFA or manual completion.

Turning the setting off leaves silent token refresh enabled, but a rejected
session shows **Sign in to reconnect** instead of opening a new sign-in window.
Changes apply to the next recovery decision without restarting the plugin;
an already-open sign-in window can still complete or be cancelled.

After failed or cancelled sign-in, click **Sign in** to retry. Background polling
does not repeatedly reopen the window, and signing out does not automatically
sign you back in.
