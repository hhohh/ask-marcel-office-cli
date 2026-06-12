/**
 * Ask Marcel Companion v3.1 - Background Service Worker
 * 
 * 通过 chrome.debugger 监听 OAuth 服务器的响应，
 * 等待 Network.loadingFinished 后再获取响应体，
 * 从中提取 Graph access_token 并发送给 CLI。
 */

console.log('[Ask Marcel] ✓ Service Worker 已启动 v3.1');

let callbackPort = null;
let debuggedTabId = null;
let tokenSent = false;

// 跟踪待处理的请求 ID
let pendingRequestIds = new Set();

// 使用 webRequest.onBeforeRequest 拦截请求，捕获 ask_marcel_port
chrome.webRequest.onBeforeRequest.addListener(
  async (details) => {
    if (tokenSent) return;
    if (details.tabId <= 0) return;

    console.log('[Ask Marcel] webRequest 检测到请求: ' + details.url.substring(0, 120));

    try {
      const url = new URL(details.url);
      if (url.hostname !== 'teams.microsoft.com' && url.hostname !== 'teams.live.com') return;

      const port = url.searchParams.get('ask_marcel_port');
      if (!port) return;

      console.log('[Ask Marcel] ✓ webRequest 捕获到 ask_marcel_port=' + port + ', tab=' + details.tabId);
      await handlePortCapture(port, details.tabId);
    } catch (e) {
      console.error('[Ask Marcel] webRequest 错误:', e.message);
    }
  },
  { urls: ['*://teams.microsoft.com/*', '*://teams.live.com/*'] },
  []
);

async function handlePortCapture(port, tabId) {
  if (callbackPort && debuggedTabId === tabId) return;

  callbackPort = port;
  debuggedTabId = tabId;

  await attachDebugger(tabId);
}

async function attachDebugger(tabId) {
  try {
    await chrome.debugger.attach({ tabId }, '1.3');
    console.log('[Ask Marcel] ✓ Debugger attached to tab ' + tabId);

    await chrome.debugger.sendCommand({ tabId }, 'Network.enable', {});
    console.log('[Ask Marcel] ✓ Network monitoring enabled, waiting for OAuth response...');
  } catch (e) {
    console.error('[Ask Marcel] ✗ Debugger attach failed:', e.message);
  }
}

// 检查是否是 OAuth 服务器
function isOAuthHost(url) {
  return url.includes('login.microsoftonline.com') ||
         url.includes('login.live.com') ||
         url.includes('login.microsoft.com');
}

// 监听 debugger 事件
chrome.debugger.onEvent.addListener((source, method, params) => {
  if (tokenSent || !callbackPort) return;
  if (source.tabId !== debuggedTabId) return;

  // 监听 OAuth 响应，记录 requestId 等待 loadingFinished
  if (method === 'Network.responseReceived') {
    const url = params.response?.url;
    if (!url) return;

    // 检查是否是 OAuth 服务器的 JSON 响应
    if (!isOAuthHost(url)) return;
    
    const contentType = params.response?.headers?.['content-type'] || 
                        params.response?.headers?.['Content-Type'] || '';
    if (!contentType.includes('json')) return;

    console.log('[Ask Marcel] 检测到 OAuth 响应: ' + url.substring(0, 80));
    
    // 记录 requestId，等待 loadingFinished
    pendingRequestIds.add(params.requestId);
    console.log('[Ask Marcel] 等待请求完成: ' + params.requestId);
  }

  // 响应完全加载后获取 body
  if (method === 'Network.loadingFinished') {
    const requestId = params.requestId;
    if (!pendingRequestIds.has(requestId)) return;

    pendingRequestIds.delete(requestId);
    console.log('[Ask Marcel] 请求完成，获取响应体: ' + requestId);

    chrome.debugger.sendCommand(
      { tabId: debuggedTabId },
      'Network.getResponseBody',
      { requestId: requestId },
      (result) => {
        if (chrome.runtime.lastError || !result) {
          console.log('[Ask Marcel] 获取响应体失败:', chrome.runtime.lastError?.message);
          return;
        }

        try {
          const body = result.body;
          if (!body.includes('access_token')) {
            console.log('[Ask Marcel] OAuth 响应中没有 access_token');
            return;
          }

          const data = JSON.parse(body);
          const token = data.access_token;
          if (!token || !token.startsWith('eyJ')) {
            console.log('[Ask Marcel] access_token 格式无效');
            return;
          }

          // 验证是否是 Graph token
          const payload = JSON.parse(atob(token.split('.')[1].replace(/-/g, '+').replace(/_/g, '/')));
          console.log('[Ask Marcel] OAuth 响应中的 token, aud:', payload.aud);

          if (payload.aud === 'https://graph.microsoft.com' ||
              payload.aud === '00000003-0000-0000-c000-000000000000') {
            console.log('[Ask Marcel] ✓ 从 OAuth 响应捕获到 Graph token');
            sendTokenToCli(token);
          } else {
            console.log('[Ask Marcel] token 不是 Graph token, aud:', payload.aud);
          }
        } catch (e) {
          console.log('[Ask Marcel] OAuth 响应解析失败:', e.message);
        }
      }
    );
  }
});

chrome.debugger.onDetach.addListener((source) => {
  if (source.tabId === debuggedTabId) {
    console.log('[Ask Marcel] Debugger detached');
    debuggedTabId = null;
  }
});

async function sendTokenToCli(accessToken) {
  if (tokenSent) return;
  tokenSent = true;

  const url = 'http://127.0.0.1:' + callbackPort + '/token';

  try {
    console.log('[Ask Marcel] 发送 token 到 CLI...');
    const response = await fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ access_token: accessToken })
    });

    if (response.ok) {
      console.log('[Ask Marcel] ✓ Token 发送成功！');
    } else {
      console.error('[Ask Marcel] ✗ 发送失败:', response.status);
      tokenSent = false;
    }
  } catch (e) {
    console.error('[Ask Marcel] ✗ 发送错误:', e.message);
    tokenSent = false;
  }

  setTimeout(async () => {
    if (debuggedTabId !== null) {
      try {
        await chrome.debugger.detach({ tabId: debuggedTabId });
        console.log('[Ask Marcel] ✓ Debugger detached (清理完成)');
      } catch (e) {}
      debuggedTabId = null;
    }
  }, 2000);
}

chrome.runtime.onInstalled.addListener(() => {
  console.log('[Ask Marcel] ✓ Extension v3.1 installed');
});
