# AI 办公助手 - 完整 API 参考

## 📦 前端引擎 API (`js/ppt-engine.js`)

> 前端只需引入 `js/ppt-engine.js`，调用以下函数即可。

### 🚀 初始化

```js
PPTEngine.init({
  provider: 'gemini',   // 默认服务商
  apikey: '',           // API密钥
  style: 'business',    // PPT风格
  onReady: ({ host, settings }) => {
    // 初始化完成回调
  }
}).then(({ host, settings }) => {
  // host: 'PowerPoint' | 'Word' | 'Browser'
});
```

### ⚙️ 配置

```js
PPTEngine.getSettings()                    // → { provider, apikey, endpoint, model, style }
PPTEngine.saveSettings({ provider:'gemini', apikey:'xxx', style:'tech' })
PPTEngine.getProviders()                   // → [{ id:'gemini', endpoint:'...', model:'...', type:'gemini' }, ...]
PPTEngine.getStyles()                      // → [{ id:'business', label:'商务专业风格...' }, ...]
PPTEngine.checkConnection()                // → Promise<boolean>
```

### 📄 Office 文档操作

```js
PPTEngine.isOfficeReady()                  // → boolean
PPTEngine.getHost()                        // → 'PowerPoint' | 'Word' | 'Browser'
PPTEngine.getSelection()                   // → Promise<string>  读取选中文字
PPTEngine.writeToOffice(text)              // → Promise<void>    替换选中文档
```

### 📊 PPT 编写（无需选中文字）

```js
PPTEngine.generateOutline(topic)           // → Promise<string>  生成完整大纲
PPTEngine.generateFullContent(topic)       // → Promise<string>  生成每页标题+要点+备注
PPTEngine.generateSlideContent(title)      // → Promise<string>  生成单页详细内容
```

### ✏️ PPT 修改（需先选中文字）

```js
PPTEngine.polishText(text)                 // → Promise<string>  精修润色
PPTEngine.expandContent(text)              // → Promise<string>  内容扩展
PPTEngine.condenseContent(text)            // → Promise<string>  精简缩写
PPTEngine.translate(text)                  // → Promise<string>  中英互译
PPTEngine.generateSpeakerNotes(text)       // → Promise<string>  生成演讲备注
PPTEngine.getDesignTips(text)              // → Promise<string>  设计建议
```

### 💬 自由对话

```js
PPTEngine.chat(message, context?)          // → Promise<string>
// context 为选中的 PPT 文字上下文（可选）
```

### 💾 对话历史

```js
PPTEngine.getChatHistory()                 // → [{sender, text, time}, ...]
PPTEngine.clearChatHistory()
PPTEngine.exportChatHistory()              // → string (Markdown格式)
PPTEngine.addChatMessage(sender, text)     // 手动添加消息
```

---

## 🖥️ 后端服务 API (`server/`)

> 服务端代理模式：API Key 存储在服务端，前端调用后端代理请求 AI。

### 模式切换

```js
// 切换到服务端代理模式（API Key 在服务端，安全！）
PPTEngine.useServer('http://localhost:3456', 'jwt-token-here');

// 切换回直连模式（API Key 在前端）
PPTEngine.useDirect();

// 获取当前模式
PPTEngine.getMode(); // → { mode: 'proxy', serverUrl: '...', authenticated: true }
```

### 🔐 用户认证

```js
// 注册
const { token, user } = await PPTEngine.register('username', 'password', '昵称');

// 登录
const { token, user } = await PPTEngine.login('username', 'password');

// 获取用户信息
const userInfo = await PPTEngine.getUserInfo();

// 登出
PPTEngine.logout();
```

### ⚙️ 服务端设置

```js
// 从服务端加载设置（覆盖本地设置）
await PPTEngine.loadServerSettings();

// 保存设置到服务端
await PPTEngine.saveServerSettings({ provider: 'deepseek', style: 'tech' });
```

### 💬 服务端聊天历史

```js
// 加载服务端聊天历史
const { messages, pagination } = await PPTEngine.loadChatHistory(page, limit);

// 清空服务端聊天历史
await PPTEngine.clearServerChat();

// 导出服务端聊天记录（Markdown）
const md = await PPTEngine.exportServerChat();
```

### 📄 文档管理

```js
// 获取文档列表
const { documents, pagination } = await PPTEngine.getDocuments(page, limit, type);

// 创建文档
const { id } = await PPTEngine.createDocument('标题', 'outline', '内容', '摘要');

// 获取文档详情
const doc = await PPTEngine.getDocument(id);

// 更新文档
await PPTEngine.updateDocument(id, { title: '新标题', content: '新内容' });

// 删除文档
await PPTEngine.deleteDocument(id);
```

### 🔗 后端 REST API 端点

| 方法 | 端点 | 说明 | 认证 |
|------|------|------|------|
| GET | `/api/health` | 健康检查 | 否 |
| POST | `/api/auth/register` | 用户注册 | 否 |
| POST | `/api/auth/login` | 用户登录 | 否 |
| GET | `/api/auth/me` | 当前用户信息 | 是 |
| POST | `/api/ai/chat` | 自由对话 | 可选 |
| POST | `/api/ai/generate-outline` | 生成大纲 | 可选 |
| POST | `/api/ai/generate-content` | 生成完整内容 | 可选 |
| POST | `/api/ai/generate-slide` | 生成单页 | 可选 |
| POST | `/api/ai/polish` | 精修润色 | 可选 |
| POST | `/api/ai/expand` | 内容扩展 | 可选 |
| POST | `/api/ai/condense` | 精简缩写 | 可选 |
| POST | `/api/ai/translate` | 翻译 | 可选 |
| POST | `/api/ai/speaker-notes` | 演讲备注 | 可选 |
| POST | `/api/ai/design-tips` | 设计建议 | 可选 |
| GET | `/api/ai/providers` | AI供应商列表 | 否 |
| GET | `/api/ai/styles` | 风格列表 | 否 |
| POST | `/api/ai/check-connection` | 测试连接 | 否 |
| GET | `/api/chat` | 聊天历史 | 是 |
| DELETE | `/api/chat` | 清空聊天 | 是 |
| GET | `/api/chat/export` | 导出聊天 | 是 |
| GET | `/api/documents` | 文档列表 | 是 |
| POST | `/api/documents` | 创建文档 | 是 |
| GET/PUT/DELETE | `/api/documents/:id` | 文档CRUD | 是 |
| GET/PUT | `/api/settings` | 用户设置 | 是 |

### 📦 完整调用示例（服务端模式）

```js
// 1. 初始化
await PPTEngine.init();

// 2. 切换服务端模式
PPTEngine.useServer('http://localhost:3456');

// 3. 注册/登录
await PPTEngine.register('myuser', 'mypassword');

// 4. 调用 AI（自动走服务端代理，API Key 不暴露）
const outline = await PPTEngine.generateOutline('人工智能发展报告');

// 5. 保存文档
await PPTEngine.createDocument('AI报告大纲', 'outline', outline);

// 6. 以上所有聊天记录自动保存到服务端
const { messages } = await PPTEngine.loadChatHistory();
```
