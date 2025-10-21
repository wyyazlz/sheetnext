# SheetNext AI 中转配置指南

> 一句话总结超级简单：写一个接口将前端传入的message消息分发给你想对接的大模型，然后在前端配置好接口地址即可开始工作！

## 文档导航

- [← 返回 README](https://github.com/wyyazlz/sheetnext/blob/master/README.md) - 快速开始和特点介绍
- [← API 文档](https://github.com/wyyazlz/sheetnext/blob/master/DOCS.md) - 详细的类、方法和属性说明

---

## 目录

- [SheetNext AI 中转配置指南](#sheetnext-ai-中转配置指南)
  - [文档导航](#文档导航)
  - [目录](#目录)
  - [功能说明](#功能说明)
    - [核心功能](#核心功能)
  - [核心架构](#核心架构)
    - [工作流程](#工作流程)
  - [完整示例](#完整示例)
    - [安装依赖](#安装依赖)
    - [完整代码](#完整代码)
    - [配置说明](#配置说明)
  - [消息格式](#消息格式)
    - [请求格式](#请求格式)
    - [响应格式](#响应格式)

---

## 功能说明

AI 服务中转层是连接 SheetNext 前端与大模型 API 的桥梁，主要负责以下核心功能：

### 核心功能

1. **消息格式转换** - 将 SheetNext 提供的通用消息结构转换为目标大模型（如 Claude、GPT 等）所需的标准格式
2. **流式数据处理** - 实现 AI 响应的流式接收与转发，提升用户交互体验
3. **安全隔离** - 在服务端隐藏真实的 API Key，避免密钥泄露风险
4. **使用统计** - 企业可在中转层统计 Token 消耗、请求次数等关键数据

---

## 核心架构

```
┌─────────────┐         ┌──────────────┐         ┌─────────────┐
│             │         │              │         │             │
│  SheetNext  │────────▶│  中转服务器   │────────▶│各种大模型API│
│   前端      │  HTTP   │  (您的服务器) │  HTTPS  │  (Claude等)  │
│             │◀────────│              │◀────────│             │
└─────────────┘  SSE流  └──────────────┘  Stream └─────────────┘
                                  │
                                  ▼
                          ┌──────────────┐
                          │ 使用统计/日志 │
                          └──────────────┘
```

### 工作流程

1. **前端请求** - SheetNext 发送包含 `messages` 数组的 POST 请求到中转服务器
2. **格式转换** - 中转服务器将通用格式转换为目标大模型的专用格式
3. **API 调用** - 使用服务端存储的 API Key 调用大模型 API
4. **流式响应** - 接收大模型的流式响应，转换后通过 SSE (Server-Sent Events) 返回前端

---

## 完整示例

通用中转完整实现示例：

### 安装依赖

```bash
npm install @anthropic-ai/sdk openai
```

### 完整代码

```javascript
/**
 * SheetNext AI & claude/openai 中转服务器示例 Node.js 版本
 * 2025.10.17 v1.0.0
 */

const http = require('http');
const Anthropic = require('@anthropic-ai/sdk');
const OpenAI = require('openai');

// ======= 配置 =======
const CONFIG = {
    model: 'claude-sonnet-4-5-20250929', // 设置模型名称，自动判断使用 claude 还是 openai
    claude: {
        apiKey: 'sk-xWp4TFA81arQCudIbLRmE0h1TtmM0lQWz4Lt7lKryUhk5HhN',
        baseURL: 'https://m5.aitoo.fun/'
    },
    openai: {
        apiKey: 'sk-UWq0SWmDIEFTI5hswtg5uONjQ95ECUneWtp46Pp9Kdujg9py',
        baseURL: 'https://m5.aitoo.fun/v1'
    }
};

const anthropic = new Anthropic({ apiKey: CONFIG.claude.apiKey, baseURL: CONFIG.claude.baseURL });
const openai = new OpenAI({ apiKey: CONFIG.openai.apiKey, baseURL: CONFIG.openai.baseURL });

// ======= message默认是openai格式，claude请求时转为它适配格式 =======
const convertToClaudeMessages = (messages) => {
    const system = [];
    const claudeMessages = [];
    let isFirstSystem = true;

    // 转换内容部分的辅助函数
    const convertContent = (content) => {
        const parts = Array.isArray(content) ? content : [{ type: 'text', text: content }];
        return parts.map(part => {
            if (part.type === 'text') {
                return { type: 'text', text: part.text };
            }
            if (part.type === 'image_url') {
                const [, mediaType, base64Data] = part.image_url.url.match(/data:(.*?);base64,(.*)/) || [];
                if (base64Data) {
                    return { type: 'image', source: { type: 'base64', media_type: mediaType || 'image/jpeg', data: base64Data } };
                }
            }
            return null;
        }).filter(Boolean);
    };

    for (const msg of messages) {
        if (msg.role === 'system') {
            if (isFirstSystem) {
                // 第一个 system：提取文本作为 system 参数（约定无图片）
                const text = typeof msg.content === 'string' ? msg.content : msg.content[0]?.text || '';
                if (text) system.push({ type: 'text', text });
                isFirstSystem = false;
            } else {
                // 其他 system：转为 user
                claudeMessages.push({ role: 'user', content: convertContent(msg.content) });
            }
        } else {
            // user/assistant 消息
            claudeMessages.push({ role: msg.role, content: convertContent(msg.content) });
        }
    }

    return { system, messages: claudeMessages };
};

// ======= Claude SDK =======
async function callClaudeSDK(messages, model, onChunk) {
    const { system, messages: claudeMessages } = convertToClaudeMessages(messages);

    // 打印请求结构（省略 base64 数据）
    const printableRequest = {
        system: system.map(s => s.type === 'image'
            ? { type: 'image', source: { ...s.source, data: `[${s.source.data?.length || 0} chars]` } }
            : s
        ),
        messages: claudeMessages.map(msg => ({
            role: msg.role,
            content: typeof msg.content === 'string' ? msg.content :
                msg.content.map(c => c.type === 'image'
                    ? { type: 'image', source: { ...c.source, data: `[${c.source.data?.length || 0} chars]` } }
                    : c
                )
        }))
    };

    const stream = await anthropic.messages.create({
        model: model,
        max_tokens: 8192,
        system,
        messages: claudeMessages,
        stream: true,
        thinking: { type: "enabled", budget_tokens: 2000 }
    });

    for await (const event of stream) {
        if (event.type === 'content_block_delta') {
            const { delta } = event;
            if (delta?.type === 'thinking_delta' && delta.thinking) {
                onChunk({ type: 'think', delta: delta.thinking });
            } else if (delta?.type === 'text_delta') {
                onChunk({ type: 'text', delta: delta.text });
            }
        }
    }
}

// ======= OpenAI SDK =======
async function callOpenAISDK(messages, model, onChunk) {
    const stream = await openai.chat.completions.create({
        model: model,
        messages: messages, // 直接使用 OpenAI 格式的 messages
        stream: true
    });

    for await (const chunk of stream) {
        const delta = chunk.choices[0]?.delta;
        if (delta?.content) {
            onChunk({ type: 'text', delta: delta.content });
        }
    }
}

// ======= HTTP 处理 =======
async function handleChat(messages, res) {
    res.writeHead(200, {
        'Content-Type': 'text/event-stream',
        'Cache-Control': 'no-cache',
        'Connection': 'keep-alive',
        'Access-Control-Allow-Origin': '*'
    });

    let ended = false;
    const write = (data) => !ended && !res.writableEnded && res.write(data);
    const onChunk = (chunk) => write(`data: ${JSON.stringify(chunk)}\n\n`);

    try {
        // 根据模型名称自动判断使用哪个 provider
        const provider = CONFIG.model.toLowerCase().includes('claude') ? 'claude' : 'openai';
        if (provider === 'openai') {
            await callOpenAISDK(messages, CONFIG.model, onChunk);
        } else {
            await callClaudeSDK(messages, CONFIG.model, onChunk);
        }
        write(`data: [DONE]\n\n`);
    } catch (error) {
        write(`data: ${JSON.stringify({ error: error.message })}\n\n`);
    } finally {
        ended = true;
        res.end();
    }
}

// ======= HTTP 服务器 =======
http.createServer(async (req, res) => {
    const corsHeaders = {
        'Access-Control-Allow-Origin': '*',
        'Access-Control-Allow-Methods': 'POST, OPTIONS',
        'Access-Control-Allow-Headers': 'Content-Type'
    };

    if (req.method === 'OPTIONS') {
        res.writeHead(200, corsHeaders);
        return res.end();
    }

    if (req.url === '/sheetnextAI' && req.method === 'POST') {
        let body = '';
        req.on('data', chunk => body += chunk);
        req.on('end', async () => {
            try {
                const { messages } = JSON.parse(body);
                if (!Array.isArray(messages)) throw new Error('Invalid messages');
                await handleChat(messages, res);
            } catch (error) {
                res.writeHead(400, { 'Content-Type': 'application/json' });
                res.end(JSON.stringify({ error: error.message }));
            }
        });
    } else {
        res.writeHead(404);
        res.end('Not Found');
    }
}).listen(3000, () => console.log('🚀 Server running on http://localhost:3000'));
```

### 配置说明

**判断规则:**
- 如果模型名称包含 `claude`(不区分大小写) → 使用 Claude SDK
- 其他情况 → 使用 OpenAI SDK

---

## 消息格式

### 请求格式

SheetNext 发送的请求体格式：

```json
{
  "messages": [
    {
      "role": "system",
      "content": "你是一个电子表格助手..."
    },
    {
      "role": "user",
      "content": "帮我分析销售数据"
    },
    {
      "role": "assistant",
      "content": "好的，我来帮您分析..."
    },
    {
      "role": "user",
      "content": [
        {
          "type": "text",
          "text": "某区域图片"
        },
        {
          "type": "image_url",
          "image_url": {
            "url": "data:image/png;base64,iVBORw0KGgoAAAANS..."
          }
        }
      ]
    }
  ]
}
```

### 响应格式

您的服务器应该返回 SSE 流：

```
data: {"type":"text","delta":"我"}

data: {"type":"text","delta":"来"}

data: {"type":"text","delta":"帮"}

data: [DONE]
```
