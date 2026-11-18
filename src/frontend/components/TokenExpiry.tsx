import React, { useEffect, useState } from 'react';

export function TokenExpiry({ expiresAt }: { expiresAt: number | null }) {
  const [now, setNow] = useState(Date.now);
  useEffect(() => {
    const timer = setInterval(() => setNow(Date.now()), 1_000);
    return () => clearInterval(timer);
  }, []);

  if (!expiresAt || !Number.isFinite(expiresAt)) return null;
  const remaining = Math.max(0, Math.ceil((expiresAt - now) / 60_000));
  const expiry = new Date(expiresAt);
  const time = expiry.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  const duration = remaining >= 60
    ? `${Math.floor(remaining / 60)}h ${remaining % 60}m`
    : `${remaining}m`;

  return (
    <span className="text-[10px] text-muted-foreground"
      title={`Access token expires ${expiry.toLocaleString()}. Silent refresh can extend this time; Microsoft may require sign-in sooner.`}>
      {remaining > 0 ? `Token: ${time} (${duration} left)` : 'Token expired'}
    </span>
  );
}
