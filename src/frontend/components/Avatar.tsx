import React from 'react';

const COLORS = [
  'bg-rose-500', 'bg-orange-500', 'bg-amber-500', 'bg-emerald-500',
  'bg-teal-500', 'bg-sky-500', 'bg-indigo-500', 'bg-violet-500', 'bg-fuchsia-500',
];

function initials(name: string): string {
  const parts = name.replace(/,/g, ' ').split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
}

function colorFor(id: string): string {
  let h = 0;
  for (let i = 0; i < id.length; i++) h = ((h << 5) - h + id.charCodeAt(i)) | 0;
  return COLORS[Math.abs(h) % COLORS.length];
}

export function Avatar({
  id,
  name,
  photo,
  size = 8,
  className = '',
}: {
  id: string;
  name: string;
  photo?: string | null;
  /** Tailwind size unit (8 → w-8 h-8). */
  size?: 6 | 7 | 8 | 9 | 10;
  className?: string;
}) {
  const dim = { 6: 'w-6 h-6 text-[9px]', 7: 'w-7 h-7 text-[10px]', 8: 'w-8 h-8 text-[11px]', 9: 'w-9 h-9 text-xs', 10: 'w-10 h-10 text-sm' }[size];
  if (photo) {
    return (
      <img
        src={photo}
        alt={name}
        className={`${dim} rounded-full object-cover shrink-0 ${className}`}
      />
    );
  }
  return (
    <div
      className={`${dim} ${colorFor(id)} rounded-full shrink-0 flex items-center justify-center font-semibold text-white ${className}`}
      title={name}
    >
      {initials(name)}
    </div>
  );
}

export function AvatarStack({
  members,
  photos,
  max = 3,
}: {
  members: Array<{ id: string; displayName: string }>;
  photos: Record<string, string | null>;
  max?: number;
}) {
  const shown = members.slice(0, max);
  const extra = members.length - shown.length;
  return (
    <div className="flex -space-x-2 shrink-0">
      {shown.map((m) => (
        <Avatar
          key={m.id}
          id={m.id}
          name={m.displayName}
          photo={photos[m.id]}
          size={8}
          className="ring-2 ring-background"
        />
      ))}
      {extra > 0 && (
        <div className="w-8 h-8 rounded-full bg-muted text-muted-foreground ring-2 ring-background flex items-center justify-center text-[10px] font-medium shrink-0">
          +{extra}
        </div>
      )}
    </div>
  );
}
