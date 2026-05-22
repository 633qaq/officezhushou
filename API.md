# PPTEngine API 参考

> 前端只需引入 `js/ppt-engine.js`，调用以下函数即可。

## 🚀 初始化

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

## ⚙️ 配置

```js
PPTEngine.getSettings()                    // → { provider, apikey, endpoint, model, style }
PPTEngine.saveSettings({ provider:'gemini', apikey:'xxx', style:'tech' })
PPTEngine.getProviders()                   // → [{ id:'gemini', endpoint:'...', model:'...', type:'gemini' }, ...]
PPTEngine.getStyles()                      // → [{ id:'business', label:'商务专业风格...' }, ...]
PPTEngine.checkConnection()                // → Promise<boolean>
```

## 📄 Office 文档操作

```js
PPTEngine.isOfficeReady()                  // → boolean
PPTEngine.getHost()                        // → 'PowerPoint' | 'Word' | 'Browser'
PPTEngine.getSelection()                   // → Promise<string>  读取选中文字
PPTEngine.writeToOffice(text)              // → Promise<void>    替换选中文档
```

## 📊 PPT 编写（无需选中文字）

```js
PPTEngine.generateOutline(topic)           // → Promise<string>  生成完整大纲
PPTEngine.generateFullContent(topic)       // → Promise<string>  生成每页标题+要点+备注
PPTEngine.generateSlideContent(title)      // → Promise<string>  生成单页详细内容
```

## ✏️ PPT 修改（需先选中文字）

```js
PPTEngine.polishText(text)                 // → Promise<string>  精修润色
PPTEngine.expandContent(text)              // → Promise<string>  内容扩展
PPTEngine.condenseContent(text)            // → Promise<string>  精简缩写
PPTEngine.translate(text)                  // → Promise<string>  中英互译
PPTEngine.generateSpeakerNotes(text)       // → Promise<string>  生成演讲备注
PPTEngine.getDesignTips(text)              // → Promise<string>  设计建议
```

## 💬 自由对话

```js
PPTEngine.chat(message, context?)          // → Promise<string>
// context 为选中的 PPT 文字上下文（可选）
```

## 💾 对话历史

```js
PPTEngine.getChatHistory()                 // → [{sender, text, time}, ...]
PPTEngine.clearChatHistory()
PPTEngine.exportChatHistory()              // → string (Markdown格式)
PPTEngine.addChatMessage(sender, text)     // 手动添加消息
```

## 📦 完整调用示例

```js
// 初始化
await PPTEngine.init({ onReady: ({host}) => console.log('就绪:', host) });

// 生成大纲
const outline = await PPTEngine.generateOutline('人工智能发展报告');

// 修改选中文字
const sel = await PPTEngine.getSelection();
if (sel) {
  const polished = await PPTEngine.polishText(sel);
  await PPTEngine.writeToOffice(polished);
}

// 导出对话
const md = PPTEngine.exportChatHistory();
// 下载为 .md 文件
```
