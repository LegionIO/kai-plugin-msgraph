# MFA activity policy

Azure username/password submission and cached-account selection remain automatic, including while the computer is idle. A valid session can therefore complete without prompting. Microsoft can send an initial Authenticator push immediately after these steps; the plugin cannot prevent that push while still submitting credentials automatically.

The plugin pauses before selecting an MFA method or displaying a code/approval prompt. It rechecks the page when activity returns so expired or already completed challenges are not reused. MFA verification is never approved automatically.

Each plugin's settings contain **MFA activity**:

- **Activity source**: keyboard/mouse input anywhere on the computer (default), or interaction with this plugin.
- **Idle timeout**: 1–60 minutes, default 5 minutes.
- **Interactive agent use**: enabled by default. Counts a currently executing plugin tool with an interactive tool context. Scheduled/headless tools, background polling, mounted panels, and completed/cancelled tools do not count.

UI activity requires trusted input in a visible, focused plugin view. Blur/unmount clears that view's activity. Computer-wide input checks use macOS HIDIdleTime, Windows GetLastInputInfo, or available Linux desktop idle services. GNOME/Mutter reports milliseconds; the freedesktop ScreenSaver service reports seconds where implemented. X11 can fall back to xprintidle if installed. XWayland readings are deliberately excluded on Wayland. Missing/failed probes are treated as unknown, never as active; plugin UI interaction remains usable.

A paused login resumes in the same window when activity returns, subject to the existing authentication timeout. If the window expires, choose Sign in again when active. Once an unattended window reaches MFA (including a cookie-only refresh), its partition is recorded in mfaPendingPartitions. Background attempts cannot repeatedly submit credentials or reopen that partition until activity returns; a successful sign-in clears the record. This record survives plugin restarts and contains no credentials or verification codes.

The common policy, idle reader, UI controls, and policy tests are intentionally identical in kai-plugin-pim, kai-plugin-msgraph, and kai-plugin-skynet; keep them synchronized when changing this behavior. PIM carries the gate into its Azure PIM, AWS SSO, Azure CLI, GCP, and ServiceNow adapters.
