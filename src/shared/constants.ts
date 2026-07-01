// ── Time Constants ──
export const ONE_SEC_IN_MS = 1000;
export const ONE_MIN_IN_MS = 60 * ONE_SEC_IN_MS;

// ── Microsoft Identity Platform ──
export const AAD_TENANT = 'organizations';
export const AAD_AUTHORITY = `https://login.microsoftonline.com/${AAD_TENANT}`;
export const AAD_AUTHORIZE_URL = `${AAD_AUTHORITY}/oauth2/v2.0/authorize`;
export const AAD_TOKEN_URL = `${AAD_AUTHORITY}/oauth2/v2.0/token`;
export const AAD_NATIVE_REDIRECT = 'https://login.microsoftonline.com/common/oauth2/nativeclient';

// FOCI (Family of Client IDs) first-party public clients.
// Interactive login uses the Teams client (proven to bypass consent in locked-down
// tenants); Graph calls redeem the family refresh token as the Office client, whose
// Graph .default preauth includes Chat.ReadWrite / Chat.Create / ChatMember.ReadWrite.
export const CLIENT_ID_TEAMS  = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
export const CLIENT_ID_OFFICE = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';
export const AUTH_CLIENT_ID   = CLIENT_ID_TEAMS;
export const GRAPH_CLIENT_ID  = CLIENT_ID_OFFICE;

export const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
export const AUTH_SCOPES = `${GRAPH_SCOPE} offline_access openid profile`;

// ── Microsoft Graph ──
export const GRAPH_BASE_URL = 'https://graph.microsoft.com/v1.0';

// ── Auth Timeouts ──
export const AUTH_TIMEOUT_MS = 300 * ONE_SEC_IN_MS;
export const TOKEN_REFRESH_BUFFER_MS = 5 * ONE_MIN_IN_MS;
export const AUTH_PARTITION = 'persist:kai-msgraph';

// ── Auto-Login Timeouts (shared with PIM's Azure AD form automation) ──
export const AUTO_LOGIN_POLL_INTERVAL_MS = 500;
export const AUTO_LOGIN_STEP_TIMEOUT_MS = 15 * ONE_SEC_IN_MS;
export const AUTO_LOGIN_MFA_TIMEOUT_MS = 120 * ONE_SEC_IN_MS;

// ── UI IDs ──
export const PANEL_ID = 'teams-panel';
export const NAV_ID = 'teams-nav';
export const SETTINGS_ID = 'msgraph-settings';

// ── Defaults ──
export const DEFAULT_CHAT_LIST_TOP = 50;
export const DEFAULT_MESSAGE_TOP = 25;
export const DEFAULT_POLL_INTERVAL_SECONDS = 30;
