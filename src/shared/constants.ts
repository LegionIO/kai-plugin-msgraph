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
export const CLIENT_ID_TEAMS          = '1fec8e78-bce4-4aaf-ab1b-5451cc387264';
export const CLIENT_ID_OFFICE         = 'd3590ed6-52b3-4102-aeff-aad2292ab01c';
export const CLIENT_ID_OUTLOOK_MOBILE = '27922004-5251-4030-b22d-91ecd9a37ea4'; // Presence.Read.All
export const AUTH_CLIENT_ID   = CLIENT_ID_TEAMS;
export const GRAPH_CLIENT_ID  = CLIENT_ID_OFFICE;

export const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
export const AUTH_SCOPES = `${GRAPH_SCOPE} offline_access openid profile`;

// Teams-internal chat service (IC3) + Trouter push — used to invoke Adaptive
// Card actions on bots and to receive real-time message/typing/read-receipt
// events, none of which Graph exposes. See ic3-client.ts / trouter.ts.
export const IC3_SCOPE      = 'https://ic3.teams.office.com/.default';
export const SPACES_SCOPE   = 'https://api.spaces.skype.com/.default';
export const PRESENCE_SCOPE = 'https://presence.teams.microsoft.com/.default';
export const OUTLOOK_SCOPE  = 'https://outlook.office.com/.default';
export const OUTLOOK_CLOUD_SETTINGS_URL = 'https://outlook.office.com/ows/v1/OutlookCloudSettings/settings/';
export const OUTLOOK_ROAMING_SIG_URL = 'https://outlook.office.com/ows/beta/RoamingSignatures';
export const OWA_SERVICE_URL = 'https://outlook.office.com/owa/service.svc';
export const TEAMS_AUTHSVC_URL       = 'https://teams.microsoft.com/api/authsvc/v1.0/authz';
export const TEAMS_CHATSVC_FALLBACK  = 'https://teams.microsoft.com/api/chatsvc/amer';
export const TEAMS_UPS_FALLBACK      = 'https://teams.microsoft.com/ups/noam';
export const TEAMS_REGISTRAR_FALLBACK = 'https://teams.microsoft.com/registrar/prod/V2/registrations';
export const TROUTER_CONNECT_URL     = 'wss://go.trouter.teams.microsoft.com/v4/c';
export const TROUTER_CLIENT_VERSION  = '2026.20.01.1';

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
export const MAIL_PANEL_ID = 'outlook-panel';
export const MAIL_NAV_ID = 'outlook-nav';
export const SETTINGS_ID = 'msgraph-settings';

// ── Mail ──
export const MAIL_LIST_TOP = 40;
export const MAIL_DELTA_POLL_SECONDS = 45;
export const WELL_KNOWN_FOLDERS = ['inbox', 'drafts', 'sentitems', 'archive', 'deleteditems', 'junkemail'] as const;

// ── Defaults ──
export const DEFAULT_CHAT_LIST_TOP = 50;
export const DEFAULT_MESSAGE_TOP = 25;
export const DEFAULT_POLL_INTERVAL_SECONDS = 30;
