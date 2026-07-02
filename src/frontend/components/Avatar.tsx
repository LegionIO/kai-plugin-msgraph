import React from 'react';
import type { Presence } from '../../shared/types.ts';

const COLORS = [
  'bg-rose-500', 'bg-orange-500', 'bg-amber-500', 'bg-emerald-500',
  'bg-teal-500', 'bg-sky-500', 'bg-indigo-500', 'bg-violet-500', 'bg-fuchsia-500',
];

function initials(name: string): string {
  let first = '';
  let last = '';
  const comma = name.indexOf(',');
  if (comma >= 0) {
    last = name.slice(0, comma).trim();
    first = name.slice(comma + 1).trim().split(/\s+/)[0] ?? '';
  } else {
    const parts = name.split(/\s+/).filter(Boolean);
    first = parts[0] ?? '';
    last = parts.length > 1 ? parts[parts.length - 1] : '';
  }
  if (!first && !last) return '?';
  if (!last) return first.slice(0, 2).toUpperCase();
  return ((first[0] ?? '') + (last[0] ?? '')).toUpperCase();
}

function colorFor(id: string): string {
  let h = 0;
  for (let i = 0; i < id.length; i++) h = ((h << 5) - h + id.charCodeAt(i)) | 0;
  return COLORS[Math.abs(h) % COLORS.length];
}

type PresenceKind = 'available' | 'busy' | 'dnd' | 'away' | 'oof' | 'offline' | 'unknown';

function presenceInfo(p: Presence | undefined): { kind: PresenceKind; label: string } | null {
  if (!p) return null;
  const a = (p.availability ?? '').toLowerCase();
  const act = (p.activity ?? '').toLowerCase();
  const label = (p.activity && p.activity !== p.availability ? p.activity : p.availability)
    .replace(/([A-Z])/g, ' $1').trim();
  if (a === 'donotdisturb' || act === 'donotdisturb' || act === 'focusing' || act === 'presenting' || act === 'urgentinterruptionsonly')
    return { kind: 'dnd', label: label || 'Do not disturb' };
  if (a.startsWith('busy') || act === 'inacall' || act === 'inaconferencecall' || act === 'inameeting')
    return { kind: 'busy', label: label || 'Busy' };
  if (a.startsWith('available'))
    return { kind: 'available', label: 'Available' };
  if (a === 'away' || a === 'berightback' || act === 'away' || act === 'berightback' || act === 'inactive')
    return { kind: 'away', label };
  if (act === 'outofoffice')
    return { kind: 'oof', label: 'Out of office' };
  if (a === 'offline' || act === 'offline' || act === 'offwork')
    return { kind: 'offline', label: 'Offline' };
  if (a === 'presenceunknown' || a === '')
    return null;
  return { kind: 'unknown', label: label || p.availability };
}

function PresenceBadge({ kind, size }: { kind: PresenceKind; size: number }) {
  const c = {
    available: '#6bb700',
    busy: '#c4314b',
    dnd: '#c4314b',
    away: '#ffaa44',
    oof: '#b4009e',
    offline: '#8a8886',
    unknown: '#8a8886',
  }[kind];
  const outline = kind === 'offline' || kind === 'unknown';
  return (
    <svg
      viewBox="0 0 16 16"
      width={size}
      height={size}
      style={{ position: 'absolute', bottom: 0, right: 0, transform: 'translate(15%, 15%)' }}
    >
      {outline ? (
        <circle cx="8" cy="8" r="6.5" fill="var(--background, #fff)" stroke={c} strokeWidth="2" />
      ) : (
        <circle cx="8" cy="8" r="7.5" fill={c} />
      )}
      {kind === 'available' && (
        <path d="M4.8 8.3l2 2 4.2-4.6" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      )}
      {kind === 'dnd' && <rect x="4" y="7" width="8" height="2" rx="1" fill="#fff" />}
      {kind === 'away' && (
        <path d="M8 4.2V8l2.4 1.6" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      )}
      {kind === 'oof' && (
        <path d="M9.5 5L6.5 8l3 3" fill="none" stroke="#fff" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
      )}
      {(kind === 'offline' || kind === 'unknown') && (
        <path d="M5.8 5.8l4.4 4.4M10.2 5.8l-4.4 4.4" stroke={c} strokeWidth="2" strokeLinecap="round" />
      )}
    </svg>
  );
}

export function Avatar({
  id,
  name,
  photo,
  presence,
  size = 8,
  className = '',
}: {
  id: string;
  name: string;
  photo?: string | null;
  presence?: Presence;
  /** Tailwind size unit (8 → w-8 h-8). */
  size?: 6 | 7 | 8 | 9 | 10;
  className?: string;
}) {
  const dim = { 6: 'w-6 h-6 text-[9px]', 7: 'w-7 h-7 text-[10px]', 8: 'w-8 h-8 text-[11px]', 9: 'w-9 h-9 text-xs', 10: 'w-10 h-10 text-sm' }[size];
  const badgePx = { 6: 8, 7: 9, 8: 10, 9: 11, 10: 12 }[size];
  const p = presenceInfo(presence);
  const core = photo ? (
    <img src={photo} alt={name} className="w-full h-full rounded-full object-cover" />
  ) : (
    <div className={`w-full h-full ${colorFor(id)} rounded-full ring-1 ring-background flex items-center justify-center font-semibold text-white`}>
      {initials(name)}
    </div>
  );
  return (
    <div className={`relative shrink-0 ${dim} ${className}`} title={p ? `${name} — ${p.label}` : name}>
      {core}
      {p && <PresenceBadge kind={p.kind} size={badgePx} />}
    </div>
  );
}

export function AvatarStack({
  members,
  photos,
  presence,
  max = 3,
}: {
  members: Array<{ id: string; displayName: string }>;
  photos: Record<string, string | null>;
  presence?: Record<string, Presence>;
  max?: number;
}) {
  const shown = members.slice(0, max);
  const extra = members.length - shown.length;
  return (
    <div className="flex -space-x-1.5 shrink-0">
      {shown.map((m) => (
        <Avatar
          key={m.id}
          id={m.id}
          name={m.displayName}
          photo={photos[m.id]}
          presence={presence?.[m.id]}
          size={8}
        />
      ))}
      {extra > 0 && (
        <div className="w-8 h-8 rounded-full bg-muted text-muted-foreground ring-1 ring-background flex items-center justify-center text-[10px] font-medium shrink-0">
          +{extra}
        </div>
      )}
    </div>
  );
}
