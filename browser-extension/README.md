# Ask Marcel Companion - 浏览器插件

## 工作原理

Teams 使用 **authorization code flow**：浏览器先获取 authorization code，然后 Teams 前端用 code 向 `/api/authsvc/v1.0/authz` 交换 access_token。

插件通过 `chrome.debugger` API 监听 OAuth 服务器 (`login.microsoftonline.com`) 的响应，从响应体中提取 Graph access_token 并发送给 CLI。

**流程：**
1. CLI 运行 `login` 命令时，打开浏览器无痕窗口并访问 `teams.microsoft.com/?ask_marcel_port=PORT`
2. 插件检测到 `ask_marcel_port` 参数后，自动 attach debugger 到该 tab
3. 监听 OAuth 服务器的 JSON 响应，等待 `Network.loadingFinished` 事件
4. 获取响应体，解析 JWT token，验证是否是 Graph token（audience 为 `graph.microsoft.com` 或 `00000003-0000-0000-c000-000000000000`）
5. 发送 token 到 CLI 的 localhost callback server
6. 2 秒后自动 detach debugger

**优势：**
- 使用用户默认浏览器，无需额外安装 Playwright
- 支持已登录和重新登录两种场景
- 直接从 OAuth 响应中捕获真实 token

**代价：**
- 浏览器顶部会短暂显示蓝色提示条："Ask Marcel Companion is debugging this browser"
- 获取 token 后自动消失（约 2 秒）

## 安装步骤

### Chrome / Edge

1. 打开浏览器扩展管理页面：
   - Chrome: `chrome://extensions/`
   - Edge: `edge://extensions/`
2. 开启右上角的 **"开发者模式"**
3. 点击 **"加载已解压的扩展程序"**
4. 选择本项目的 `browser-extension/` 目录
5. 确认插件已启用，版本号为 3.1.0
6. **启用无痕模式访问权限：** 点击插件的"详情"按钮，开启 **"在无痕模式下启用"**（Chrome）或 **"在 InPrivate 模式下启用"**（Edge）。**此步骤为必需** — CLI 会打开无痕/InPrivate 窗口，而浏览器默认禁止扩展在此模式下运行。

### 验证安装

插件安装后，可以在扩展管理页面看到 "Ask Marcel Companion"。

## 使用方式

```bash
# 默认使用 Playwright 浏览器
ask-marcel login

# 使用浏览器插件（需要已安装插件）
ask-marcel login --use-extension
```

## 故障排查

### 命令行一直卡着

1. 检查插件是否已启用
2. 打开浏览器开发者工具（F12），查看 Console 标签
3. 应该能看到 `[Ask Marcel]` 开头的日志
4. 如果没有日志，说明插件未正确加载

### 蓝色提示条不消失

- 正常情况下，获取 token 后 2 秒会自动消失
- 如果一直显示，可能是 token 发送失败
- 检查 CLI 是否还在运行，localhost server 是否正常监听

### 重新加载插件

如果修改了插件代码，需要在扩展管理页面点击刷新按钮重新加载。

## 安全说明

- 插件只在 `teams.microsoft.com` 和 `teams.live.com` 域名下工作
- 只捕获 Graph token（audience 必须是 `graph.microsoft.com`）
- token 只发送到 `127.0.0.1`（localhost），不会发送到外部
- 获取 token 后立即 detach debugger，最小化调试时间
