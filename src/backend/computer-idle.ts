import { execFile } from 'node:child_process';
import { win32 } from 'node:path';

type Run = (file: string, args: string[]) => Promise<string | null>;
const run: Run = (file, args) => new Promise(resolve => {
  execFile(file, args, { timeout: 5_000, maxBuffer: 256 * 1024, windowsHide: true }, (error, stdout) => resolve(error ? null : stdout));
});

const WINDOWS_IDLE_SCRIPT = `
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
public static class KaiInputIdle {
  [StructLayout(LayoutKind.Sequential)] struct LastInput { public uint size; public uint time; }
  [DllImport("user32.dll")] static extern bool GetLastInputInfo(ref LastInput input);
  public static double Seconds() {
    var input = new LastInput(); input.size = (uint)Marshal.SizeOf(input);
    if (!GetLastInputInfo(ref input)) throw new InvalidOperationException("Input idle time unavailable");
    return unchecked((uint)Environment.TickCount - input.time) / 1000.0;
  }
}
'@
[KaiInputIdle]::Seconds().ToString([System.Globalization.CultureInfo]::InvariantCulture)
`;

type Probe = { file: string; args: string[]; parse: (stdout: string) => number | null };
function numeric(value: string | undefined, scale = 1): number | null {
  if (!value || !/^\d+(?:\.\d+)?$/.test(value)) return null;
  const number = Number(value) / scale;
  return Number.isFinite(number) && number >= 0 ? number : null;
}

/** Read elapsed time since input, never the input itself. No shell or user-provided command. */
export function createComputerIdleReader(
  platform: string = process.platform,
  environment: Record<string, string | undefined> = process.env,
  execute: Run = run,
) {
  const probes: Probe[] = [];
  if (platform === 'darwin') probes.push({
    file: '/usr/sbin/ioreg', args: ['-r', '-c', 'IOHIDSystem'],
    parse: text => numeric(/"HIDIdleTime"\s*=\s*(\d+)/.exec(text)?.[1], 1e9),
  });
  if (platform === 'win32') probes.push({
    file: win32.join(environment.SystemRoot || 'C:\\Windows', 'System32', 'WindowsPowerShell', 'v1.0', 'powershell.exe'),
    args: ['-NoProfile', '-NonInteractive', '-Command', WINDOWS_IDLE_SCRIPT],
    parse: text => numeric(text.trim()),
  });
  if (platform === 'linux') {
    const dbus = (dest: string, path: string, method: string, scale: number): Probe => ({
      file: 'gdbus', args: ['call', '--session', '--dest', dest, '--object-path', path, '--method', method],
      parse: text => numeric(/^\(\s*(?:uint64\s+|uint32\s+)?(\d+)\s*,\s*\)\s*$/.exec(text)?.[1], scale),
    });
    probes.push(dbus('org.gnome.Mutter.IdleMonitor', '/org/gnome/Mutter/IdleMonitor/Core', 'org.gnome.Mutter.IdleMonitor.GetIdletime', 1000));
    probes.push(dbus('org.freedesktop.ScreenSaver', '/ScreenSaver', 'org.freedesktop.ScreenSaver.GetSessionIdleTime', 1));
    // XWayland only sees X11-client input and can falsely report a permanently
    // active session. Never use it to authorize MFA in a Wayland desktop.
    if (environment.XDG_SESSION_TYPE !== 'wayland' && !environment.WAYLAND_DISPLAY && environment.DISPLAY) {
      probes.push({ file: 'xprintidle', args: [], parse: text => numeric(text.trim(), 1000) });
    }
  }
  let selected: Probe | undefined;
  let retryAt = 0;
  return async (): Promise<number | null> => {
    if (Date.now() < retryAt) return null;
    const candidates = selected ? [selected, ...probes.filter(probe => probe !== selected)] : probes;
    for (const probe of candidates) {
      let output: string | null;
      try { output = await execute(probe.file, probe.args); } catch { output = null; }
      const value = output === null ? null : probe.parse(output);
      if (value !== null) { selected = probe; return value; }
    }
    selected = undefined;
    retryAt = Date.now() + 60_000;
    return null;
  };
}

export const readComputerIdleSeconds = createComputerIdleReader();
