import { test } from 'node:test';
import assert from 'node:assert/strict';
import { createComputerIdleReader } from '../src/backend/computer-idle';

test('macOS converts HID nanoseconds to seconds', async () => {
  const read = createComputerIdleReader('darwin', {}, async (file, args) => {
    assert.equal(file, '/usr/sbin/ioreg'); assert.deepEqual(args, ['-r', '-c', 'IOHIDSystem']);
    return '"HIDIdleTime" = 301000000000';
  });
  assert.equal(await read(), 301);
});

test('Windows queries the current session input API without a shell and parses seconds', async () => {
  const read = createComputerIdleReader('win32', { SystemRoot: 'D:\\Windows' }, async (file, args) => {
    assert.equal(file, 'D:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe');
    assert.deepEqual(args.slice(0, 3), ['-NoProfile', '-NonInteractive', '-Command']);
    assert.match(args[3], /GetLastInputInfo/);
    assert.match(args[3], /unchecked\(\(uint\)Environment.TickCount - input.time\)/);
    return '301.25\r\n';
  });
  assert.equal(await read(), 301.25);
});

test('GNOME on Wayland reads Mutter idle milliseconds', async () => {
  const read = createComputerIdleReader('linux', { XDG_SESSION_TYPE: 'wayland', DISPLAY: ':0' }, async (file, args) => {
    assert.equal(file, 'gdbus'); assert.ok(args.includes('org.gnome.Mutter.IdleMonitor.GetIdletime'));
    return '(uint64 301000,)\n';
  });
  assert.equal(await read(), 301);
});

test('Linux desktop ScreenSaver service is a seconds fallback and the working probe is retained', async () => {
  const calls: string[] = [];
  const read = createComputerIdleReader('linux', { WAYLAND_DISPLAY: 'wayland-0' }, async (_file, args) => {
    const method = args[args.length - 1]; calls.push(method);
    return method === 'org.freedesktop.ScreenSaver.GetSessionIdleTime' ? '(uint32 301,)' : null;
  });
  assert.equal(await read(), 301);
  assert.equal(await read(), 301);
  assert.deepEqual(calls, ['org.gnome.Mutter.IdleMonitor.GetIdletime', 'org.freedesktop.ScreenSaver.GetSessionIdleTime', 'org.freedesktop.ScreenSaver.GetSessionIdleTime']);
});

test('X11 can use xprintidle, but Wayland must never trust XWayland idle time', async () => {
  let x11Calls = 0;
  const execute = async (file: string) => { if (file === 'xprintidle') { x11Calls++; return '301000\n'; } return null; };
  const x11 = createComputerIdleReader('linux', { DISPLAY: ':0', XDG_SESSION_TYPE: 'x11' }, execute);
  assert.equal(await x11(), 301); assert.equal(x11Calls, 1);
  const wayland = createComputerIdleReader('linux', { DISPLAY: ':0', XDG_SESSION_TYPE: 'wayland' }, execute);
  assert.equal(await wayland(), null); assert.equal(x11Calls, 1);
});

test('missing commands, unsupported sessions and malformed output never report active', async () => {
  for (const output of [null, 'Error 0', '-1', 'NaN', '']) {
    assert.equal(await createComputerIdleReader('win32', {}, async () => output)(), null);
  }
  assert.equal(await createComputerIdleReader('linux', {}, async () => { throw new Error('Command not found'); })(), null);
  assert.equal(await createComputerIdleReader('unknown', {}, async () => assert.fail('No probe expected'))(), null);
});
