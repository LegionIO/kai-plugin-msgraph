import { CLIENT_ID_OFFICE, CLIENT_ID_TEAMS, CLIENT_ID_OUTLOOK_MOBILE } from '../shared/constants.js';

export interface RequestPolicy {
  workload: string;
  operation: string;
  resource: string;
  readOnly: boolean;
  rotation: boolean;
  /** OR of explicitly supported scope sets. Never infer scope inheritance. */
  scopes: string[][];
  spacing?: number;
  mailbox?: boolean;
  batch?: boolean;
}

// Candidate hints only. Actual tokens and endpoint authorization are authoritative.
// Sources: secureworks/family-of-client-ids-research@57e67fac; ROADtools@3b975e83.
export const FOCI_CLIENTS = [
  { id: CLIENT_ID_OFFICE, alias: 'office', workloads: ['mail', 'directory', 'people', 'teams', 'presence'] },
  { id: CLIENT_ID_TEAMS, alias: 'teams', workloads: ['mail', 'directory', 'people', 'teams'] },
  { id: CLIENT_ID_OUTLOOK_MOBILE, alias: 'outlook-mobile', workloads: ['mail', 'directory', 'people', 'presence'] },
  { id: '00b41c95-dab0-4487-9791-b9d2c32c80f2', alias: 'office-management', workloads: ['mail', 'people'] },
  { id: '57336123-6e14-4acc-8dcf-287b6088aa28', alias: 'whiteboard', workloads: ['mail', 'directory', 'people', 'teams'] },
  { id: '0ec893e0-5785-4de6-99da-4ed124e5296c', alias: 'm365-copilot', workloads: ['directory', 'people', 'presence'] },
  { id: 'd7b530a4-7680-4c23-a8bf-c52c121d2e87', alias: 'edge-new-tab', workloads: ['mail', 'directory'] },
] as const;
export const clientAlias = (id: string): string => FOCI_CLIENTS.find(c => c.id === id)?.alias ?? 'preferred';
const any = (...scopes: string[]): string[][] => scopes.map(s => [s]);

export function graphPolicy(method: string, input: string, body?: unknown): RequestPolicy {
  const url = new URL(input, 'https://graph.microsoft.com/v1.0/');
  const p = url.pathname.replace(/^\/(v1\.0|beta)/, '');
  const read = method === 'GET' || method === 'HEAD';
  const policy: RequestPolicy = { workload: 'graph-other', operation: 'other', resource: '', readOnly: read, rotation: false, scopes: [] };
  if (/\/(drive|drives|shares)(\/|$)/.test(p)) return { ...policy, workload: 'files' };
  if (p === '/$batch') return { ...policy, workload: 'mail', operation: 'batch', resource: 'me', mailbox: true, batch: true };
  if (/^\/(me|users\/[^/]+)\/(mailFolders|messages|sendMail)(\/|$)/.test(p)) {
    const sending = /\/(sendMail|send|reply|replyAll|forward)$/.test(p);
    return { ...policy, workload: 'mail', operation: sending ? 'send' : read ? 'read' : 'write',
      resource: /^\/users\/([^/]+)/.exec(p)?.[1] ?? 'me', mailbox: true, rotation: true,
      scopes: sending ? any('Mail.Send') : read ? any('Mail.Read', 'Mail.ReadWrite') : any('Mail.ReadWrite') };
  }
  if (p === '/communications/getPresencesByUserId') return { ...policy, workload: 'presence', operation: 'lookup', readOnly: true, rotation: true, scopes: any('Presence.Read.All') };
  if (/^\/(me|users\/[^/]+)\/people$/.test(p)) return { ...policy, workload: 'people', operation: 'search', rotation: read, scopes: any('People.Read', 'People.Read.All') };
  if (/^\/users(?:\/[^/]+)?$/.test(p) || p === '/me') {
    // Full profiles require more than User.ReadBasic.All (e.g. jobTitle).
    const basic = (url.searchParams.get('$select') ?? '').split(',').every(s => ['id','displayName','mail','userPrincipalName',''].includes(s));
    return { ...policy, workload: 'directory', operation: p === '/me' ? 'self' : 'users', rotation: read,
      scopes: p === '/me' ? any('User.Read','User.ReadWrite','User.Read.All','Directory.Read.All') :
        any(...(basic && url.searchParams.has('$select') ? ['User.ReadBasic.All'] : []), 'User.Read.All', 'User.ReadWrite.All', 'Directory.Read.All') };
  }
  if (/\/(photos|photo)(\/|$)/.test(p)) return { ...policy, workload: 'photos', operation: 'photo' };
  const chat = /^\/(?:me\/|users\/[^/]+\/)?chats(?:\/([^/]+))?/.exec(p);
  if (chat) {
    const op = /\/messages(?:\/|$)/.test(p) ? (read ? 'messages-read' : 'messages-write') : chat[1] ? 'chat' : 'list-chats';
    const scopes = read ? any('Chat.Read','Chat.ReadWrite') : /\/messages$/.test(p) ? any('ChatMessage.Send','Chat.ReadWrite') : [];
    return { ...policy, workload: 'teams', operation: op, resource: chat[1] ?? '', rotation: scopes.length > 0, scopes,
      spacing: op === 'list-chats' || op === 'messages-read' || op === 'messages-write' ? 1500 : undefined };
  }
  if (/\/joinedTeams$/.test(p)) return { ...policy, workload: 'teams', operation: 'joined-teams', rotation: read, scopes: any('Team.ReadBasic.All','Team.Read.All','Team.ReadWrite.All') };
  if (/^\/teams\//.test(p)) {
    const channel = /\/channels\/([^/]+)/.exec(p)?.[1];
    const messages = /\/messages(?:\/|$)/.test(p);
    const scopes = messages ? read ? any('ChannelMessage.Read.All') : method === 'POST' && /\/(messages|replies)$/.test(p) ? any('ChannelMessage.Send') : [] :
      read && /\/channels(?:\/[^/]+)?$/.test(p) ? any('Channel.ReadBasic.All','ChannelSettings.Read.All','ChannelSettings.ReadWrite.All') :
      read && /^\/teams\/[^/]+$/.test(p) ? any('Team.ReadBasic.All','Team.Read.All','Team.ReadWrite.All') : [];
    return { ...policy, workload: 'teams', operation: messages ? (read ? 'channel-read' : 'channel-write') : 'channels', resource: channel ?? p.split('/')[2],
      rotation: scopes.length > 0, scopes, spacing: channel ? 1500 : undefined };
  }
  if (p === '/search/query') {
    const entities = (body as { requests?: Array<{ entityTypes?: string[] }> })?.requests?.flatMap(r => r.entityTypes ?? []) ?? [];
    return { ...policy, workload: 'teams', operation: 'search', readOnly: true,
      rotation: entities.length > 0 && entities.every(e => e === 'chatMessage'), scopes: any('Chat.Read','Chat.ReadWrite') };
  }
  return policy;
}

export function servicePolicy(workload: string, method: string, readOnly = method === 'GET'): RequestPolicy {
  return { workload, operation: 'service', resource: '', readOnly, rotation: false, scopes: [] };
}
