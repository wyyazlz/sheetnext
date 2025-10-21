<div align="center">
  <div><img src="docs/logo.png" alt="SheetNext Logo" width="100" style="vertical-align: middle;"/></div>
  <div>✨ 数行代码可集成一个纯前端高性能 Excel 编辑器。</div>
  <div>🤖 内置超级AI工作流程，目标用 AI 操纵 Excel 表格完成所有任务。</div>
  <div>
    <a href="https://www.sheetnext.com/">🏠 官网</a> |
    <a href="https://www.sheetnext.com/editor">🎯 在线体验</a> |
    <a href="https://github.com/wyyazlz/sheetnext/blob/main/AGENT.md">🤖 AI中转文档</a> |
    <a href="https://github.com/wyyazlz/sheetnext/blob/main/DOCS.md">📖 编辑器操作文档</a>
  </div>
</div>

---

[English](./README.md) | 简体中文

## ✨ 特点

- 📊 **电子表格功能** - 支持电子表格核心功能如：单元格编辑、样式、公式引擎、图表、排序、筛选等
- 🤖 **AI 工作流** - 内置 AI 全自动操作工作流，简单配置可以生成模板、数据分析、公式编写、跨表逻辑操纵等
- 📁 **导入导出** - 原生支持 Excel (.xlsx) 文件的导入和导出，无需插件前端秒操作
- 🚀 **开箱即用** - 不用单独配置任何库和插件，安装后简单配置即用所有功能
- 🔄 **快速迭代** - 版本快速迭代更新，有问题可随时联系或提交 issue

## 🚀 快速开始

### 📦 使用 npm 安装

```bash
npm install sheetnext
```

```html
<div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
```

```javascript
import SheetNext from 'sheetnext';
import 'sheetnext/dist/sheetnext.css';

// 注意设置容器#SNContainer宽高
const SN = new SheetNext(document.querySelector('#SNContainer'));
```

### 🌐 浏览器直接引入

```html
<!-- 引入样式 -->
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sheetnext/dist/sheetnext.css">

<!-- 编辑器容器 -->
<div id="SNContainer" style="width: 100vw; height: 100vh;padding:0 7px 7px"></div>

<!-- 引入脚本 -->
<script src="https://cdn.jsdelivr.net/npm/sheetnext/dist/sheetnext.umd.js"></script>

<!-- 初始化，注意设置宽高 -->
<script>
  const SN = new SheetNext(document.querySelector('#SNContainer'));
</script>
```

## ⚙️ 初始化配置

```javascript
const SN = new SheetNext(document.querySelector('#container'), {
  AI_URL: "http://localhost:3000/sheetnextAI",  // AI 中转地址（可选）
  AI_TOKEN: "your-token"                        // 中转 token（可选）
});
```

## 📚 文档

- [DOCS.md](./DOCS.md) - 完整 API 文档
- [AGENT.md](./AGENT.md) - AI Agent 集成指南
- [官方网站](https://www.sheetnext.com) - 更多示例和教程

## 🤝 贡献

欢迎提交 Issue 和参与讨论！详见 [CONTRIBUTING.md](./CONTRIBUTING.md)

## 📄 许可证

[Apache-2.0](./LICENSE)

## 🌟 社区版说明

这是 SheetNext 的社区版本，包含编译后的分发文件。核心源代码因商业原因保持私有。

您可以：
- ✅ 免费使用和分发
- ✅ 用于商业项目
- ✅ 修改和定制（通过 API）
- ✅ 报告问题和提出建议

## 🔗 相关链接

- 🏠 [官网](https://www.sheetnext.com)
- 📦 [npm 包地址](https://www.npmjs.com/package/sheetnext)
- 💬 [问题反馈](https://github.com/wyyazlz/sheetnext/issues)
