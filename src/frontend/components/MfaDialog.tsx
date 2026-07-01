import React, { useState } from 'react';

type MfaDialogProps = {
  type: 'sms' | 'totp';
  onSubmit: (code: string) => void;
  onCancel: () => void;
};

export function MfaDialog({ type, onSubmit, onCancel }: MfaDialogProps) {
  const [code, setCode] = useState('');

  const title = type === 'sms' ? 'Enter SMS Code' : 'Enter Authenticator Code';
  const subtitle = type === 'sms'
    ? 'Enter the verification code sent to your phone'
    : 'Enter the code from your authenticator app';

  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50">
      <div className="bg-card rounded-xl shadow-2xl w-full max-w-sm p-6 border border-border">
        <h3 className="text-lg font-semibold text-foreground mb-1">{title}</h3>
        <p className="text-sm text-muted-foreground mb-4">{subtitle}</p>
        <input
          type="text"
          value={code}
          onChange={(e: React.ChangeEvent<HTMLInputElement>) => setCode(e.target.value)}
          onKeyDown={(e: React.KeyboardEvent) => { if (e.key === 'Enter' && code.trim()) onSubmit(code.trim()); }}
          autoFocus
          placeholder="000000"
          className="w-full px-3 py-2 bg-muted border border-border rounded-lg text-center text-lg font-mono text-foreground focus:border-primary transition-colors tracking-widest"
          maxLength={8}
        />
        <div className="flex justify-end gap-3 mt-4">
          <button
            type="button"
            onClick={onCancel}
            className="px-4 py-2 text-sm font-medium text-muted-foreground bg-muted border border-border rounded-lg hover:bg-muted/80 transition-colors"
          >
            Cancel
          </button>
          <button
            type="button"
            onClick={() => code.trim() && onSubmit(code.trim())}
            disabled={!code.trim()}
            className="px-4 py-2 text-sm font-medium bg-primary text-primary-foreground rounded-lg hover:bg-primary/90 disabled:opacity-50 transition-colors"
          >
            Verify
          </button>
        </div>
      </div>
    </div>
  );
}
