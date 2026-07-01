import React from 'react';

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
  return (
    <div
      className={`absolute -top-9 z-20 flex items-center gap-0.5 rounded-full border border-border/50 bg-card shadow-lg px-1.5 py-1 ${
        align === 'right' ? 'right-0' : 'left-0'
      }`}
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
        onClick={onClose}
        className="ml-0.5 flex h-6 w-6 items-center justify-center rounded-full text-muted-foreground/60 hover:bg-muted/50 hover:text-muted-foreground"
      >
        ×
      </button>
    </div>
  );
}
