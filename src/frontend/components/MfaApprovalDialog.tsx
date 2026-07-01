import React from 'react';

type MfaApprovalDialogProps = {
  approvalNumber: string | null;
};

export function MfaApprovalDialog({ approvalNumber }: MfaApprovalDialogProps) {
  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50">
      <div className="bg-card rounded-xl shadow-2xl w-full max-w-sm p-6 border border-border text-center">
        <div className="mb-4">
          <svg
            className="mx-auto w-12 h-12 text-primary animate-pulse"
            fill="none"
            viewBox="0 0 24 24"
            stroke="currentColor"
            strokeWidth={1.5}
          >
            <path
              strokeLinecap="round"
              strokeLinejoin="round"
              d="M10.5 1.5H8.25A2.25 2.25 0 006 3.75v16.5a2.25 2.25 0 002.25 2.25h7.5A2.25 2.25 0 0018 20.25V3.75a2.25 2.25 0 00-2.25-2.25H13.5m-3 0V3h3V1.5m-3 0h3m-3 18.75h3"
            />
          </svg>
        </div>
        <h3 className="text-lg font-semibold text-foreground mb-2">Check Your Phone</h3>
        <p className="text-sm text-muted-foreground mb-4">
          Approve the sign-in request in your Microsoft Authenticator app
        </p>
        {approvalNumber ? (
          <div className="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-primary/15 mb-2">
            <span className="text-3xl font-bold text-primary">{approvalNumber}</span>
          </div>
        ) : null}
        {approvalNumber ? (
          <p className="text-xs text-muted-foreground">Enter this number in the Authenticator app</p>
        ) : null}
      </div>
    </div>
  );
}
