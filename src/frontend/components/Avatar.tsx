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

function presenceColor(p: Presence | undefined): { cls: string; label: string } | null {
  if (!p) return null;
  const a = (p.availability ?? '').toLowerCase();
  const act = (p.activity ?? '').toLowerCase();
  const label = (p.activity && p.activity !== p.availability ? p.activity : p.availability)
    .replace(/([A-Z])/g, ' $1').trim();
  if (a === 'donotdisturb' || act === 'donotdisturb' || act === 'presenting' || act === 'urgentinterruptionsonly')
    return { cls: 'bg-rose-600', label };
  if (a.startsWith('busy') || act === 'inacall' || act === 'inaconferencecall' || act === 'inameeting')
    return { cls: 'bg-red-500', label };
  if (a.startsWith('available'))
    return { cls: 'bg-emerald-500', label: 'Available' };
  if (a === 'away' || a === 'berightback' || act === 'away' || act === 'berightback' || act === 'inactive')
    return { cls: 'bg-amber-400', label };
  if (act === 'outofoffice')
    return { cls: 'bg-fuchsia-500', label: 'Out of office' };
  if (a === 'offline' || act === 'offline' || act === 'offwork')
    return { cls: 'bg-zinc-400', label: 'Offline' };
  if (a === 'presenceunknown' || a === '')
    return null;
  return { cls: 'bg-zinc-400', label: label || p.availability };
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
  const dot = { 6: 'w-2 h-2', 7: 'w-2 h-2', 8: 'w-2.5 h-2.5', 9: 'w-3 h-3', 10: 'w-3 h-3' }[size];
  const p = presenceColor(presence);
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
      {p && (
        <span
          className={`absolute bottom-0 right-0 translate-x-[15%] translate-y-[15%] ${dot} rounded-full ${p.cls}`}
        />
      )}
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
