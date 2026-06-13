# Ask Marcel Companion - Browser Extension

## How It Works

Teams uses **authorization code flow**: the browser first obtains an authorization code, then the Teams frontend exchanges the code for an access_token via `/api/authsvc/v1.0/authz`.

The extension uses the `chrome.debugger` API to monitor OAuth server responses (`login.microsoftonline.com`), extracts the Graph access_token from the response body, and sends it to the CLI.

**Flow:**
1. When the CLI runs the `login` command, it opens an incognito/inprivate browser window and navigates to `teams.microsoft.com/?ask_marcel_port=PORT`
2. The extension detects the `ask_marcel_port` parameter and automatically attaches the debugger to that tab
3. Monitors JSON responses from the OAuth server, waiting for the `Network.loadingFinished` event
4. Retrieves the response body, parses the JWT token, and verifies it's a Graph token (audience must be `graph.microsoft.com` or `00000003-0000-0000-c000-000000000000`)
5. Sends the token to the CLI's localhost callback server
6. Automatically detaches the debugger after 2 seconds

**Advantages:**
- Uses the user's default browser, no need to install Playwright separately
- Supports both already-logged-in and re-login scenarios
- Captures the real token directly from the OAuth response

**Trade-offs:**
- A blue notification bar briefly appears at the top of the browser: "Ask Marcel Companion is debugging this browser"
- The bar disappears automatically after the token is captured (about 2 seconds)

## Installation

### Chrome / Edge

1. Open the browser extension management page:
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`
2. Enable **"Developer mode"** in the top-right corner
3. Click **"Load unpacked"**
4. Select the `browser-extension/` directory from this project
5. Confirm the extension is enabled, version 3.1.0
6. **Enable incognito/inprivate mode access:** Click the extension's "Details" button, then enable **"Allow in incognito"** (Chrome) or **"Allow in InPrivate"** (Edge). **This step is required** — the CLI opens an incognito/inprivate window, and extensions are disabled in this mode by default.

### Verify Installation

After installation, you can see "Ask Marcel Companion" on the extension management page.

## Usage

```bash
# Default: use Playwright browser
ask-marcel login

# Use browser extension (requires extension to be installed)
ask-marcel login --use-extension
```

## Troubleshooting

### Command line hangs

1. Check if the extension is enabled
2. Open browser developer tools (F12), check the Console tab
3. You should see logs starting with `[Ask Marcel]`
4. If no logs appear, the extension is not loaded correctly

### Blue notification bar doesn't disappear

- Normally, it disappears automatically 2 seconds after capturing the token
- If it stays visible, the token may have failed to send
- Check if the CLI is still running and if the localhost server is listening properly

### Reload extension

If you modified the extension code, you need to click the refresh button on the extension management page to reload it.

## Security

- The extension only works under `teams.microsoft.com` and `teams.live.com` domains
- Only captures Graph tokens (audience must be `graph.microsoft.com`)
- Tokens are only sent to `127.0.0.1` (localhost), never to external servers
- The debugger is detached immediately after capturing the token, minimizing debugging time
