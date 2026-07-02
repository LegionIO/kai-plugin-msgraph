import React, { useState } from 'react';
import { EMOJI_CATEGORIES } from '../emoji-data.ts';

export const TEAMS_REACTIONS = [
  { type: 'like', emoji: '👍', label: 'Like' },
  { type: 'heart', emoji: '❤️', label: 'Heart' },
  { type: 'laugh', emoji: '😆', label: 'Laugh' },
  { type: 'surprised', emoji: '😮', label: 'Surprised' },
  { type: 'sad', emoji: '😢', label: 'Sad' },
  { type: 'angry', emoji: '😡', label: 'Angry' },
] as const;

export function ReactionPicker({
  onSelect,
  onClose,
  align = 'left',
}: {
  onSelect: (reactionType: string) => void;
  onClose: () => void;
  align?: 'left' | 'right';
}) {
  const [expanded, setExpanded] = useState(false);
  const [custom, setCustom] = useState('');

  const submitCustom = () => {
    const v = custom.trim();
    if (!v) return;
    onSelect(Array.from(v)[0] ?? v);
    setCustom('');
  };

  const anchor: React.CSSProperties = { position: 'absolute', top: -4, right: 0, transform: 'translateY(-100%)' };
  void align;

  if (!expanded) {
    return (
      <div
        style={anchor}
        className="z-20 flex items-center gap-0.5 rounded-full border border-border/50 bg-card shadow-lg px-1.5 py-1"
        onMouseDown={(e) => e.stopPropagation()}
      >
        {TEAMS_REACTIONS.map((r) => (
          <button
            key={r.type}
            type="button"
            title={r.label}
            onClick={() => onSelect(r.type)}
            className="flex h-7 w-7 items-center justify-center rounded-full text-base transition-transform hover:scale-125 hover:bg-muted/50"
          >
            {r.emoji}
          </button>
        ))}
        <button
          type="button"
          title="More reactions"
          onClick={() => setExpanded(true)}
          className="flex h-7 w-7 items-center justify-center rounded-full text-muted-foreground hover:bg-muted/50 hover:text-foreground"
        >
          <svg className="w-4 h-4" viewBox="0 0 24 24" fill="currentColor">
            <circle cx="5" cy="12" r="1.6" /><circle cx="12" cy="12" r="1.6" /><circle cx="19" cy="12" r="1.6" />
          </svg>
        </button>
        <button
          type="button"
          onClick={onClose}
          className="ml-0.5 flex h-6 w-6 items-center justify-center rounded-full text-muted-foreground/60 hover:bg-muted/50 hover:text-muted-foreground"
        >
          ×
        </button>
      </div>
    );
  }

  return (
    <div
      style={{ ...anchor, width: 260 }}
      className="z-20 rounded-xl border border-border/50 bg-card shadow-2xl p-2"
      onMouseDown={(e) => e.stopPropagation()}
    >
      <div className="flex items-center justify-between px-1 pb-1.5">
        <button
          type="button"
          onClick={() => setExpanded(false)}
          className="text-[11px] text-muted-foreground hover:text-foreground"
        >
          ‹ Back
        </button>
        <button
          type="button"
          onClick={onClose}
          className="flex h-5 w-5 items-center justify-center rounded-full text-muted-foreground/60 hover:bg-muted/50 hover:text-muted-foreground"
        >
          ×
        </button>
      </div>
      <div style={{ maxHeight: 260 }} className="overflow-y-auto overscroll-contain pr-0.5">
        {EMOJI_CATEGORIES.map((cat) => (
          <div key={cat.name}>
            <div className="px-1 pt-1.5 pb-1 text-[10px] font-semibold uppercase tracking-wide text-muted-foreground">
              {cat.name}
            </div>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(8, minmax(0, 1fr))', gap: 2 }}>
              {cat.emoji.map((e, i) => (
                <button
                  key={`${cat.name}-${i}`}
                  type="button"
                  onClick={() => onSelect(e)}
                  style={{ width: 28, height: 28 }}
                  className="flex items-center justify-center rounded-md text-base transition-transform hover:scale-110 hover:bg-muted/50"
                >
                  {e}
                </button>
              ))}
            </div>
          </div>
        ))}
      </div>
      <div className="mt-2 flex gap-1.5">
        <input
          value={custom}
          onChange={(e) => setCustom(e.target.value)}
          onKeyDown={(e) => e.key === 'Enter' && submitCustom()}
          placeholder="Any emoji…"
          className="flex-1 min-w-0 px-2 py-1 text-xs bg-muted border border-border rounded-md text-foreground placeholder:text-muted-foreground focus:border-primary transition-colors"
        />
        <button
          type="button"
          onClick={submitCustom}
          disabled={!custom.trim()}
          className="px-2 py-1 text-[11px] font-medium bg-primary text-primary-foreground rounded-md hover:bg-primary/90 disabled:opacity-50"
        >
          React
        </button>
      </div>
    </div>
  );
}
