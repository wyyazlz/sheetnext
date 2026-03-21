<p align="center">
  <img src="docs/logo.png" alt="SheetNext Logo" width="80" />
</p>

<p align="center">
  A pure front-end spreadsheet component with Excel-like capabilities, built-in native AI workflows, and flexible LLM integration for data operations.
</p>

<p align="center">
  English | <a href="./README_CN.md">简体中文</a>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/sheetnext"><img src="https://img.shields.io/npm/v/sheetnext.svg" alt="npm version" /></a>
  <a href="https://www.npmjs.com/package/sheetnext"><img src="https://img.shields.io/npm/dm/sheetnext.svg" alt="npm downloads" /></a>
  <a href="./LICENSE"><img src="https://img.shields.io/badge/license-Apache%202.0-blue.svg" alt="license" /></a>
  <img src="https://img.shields.io/badge/rendering-Canvas-ff8c00.svg" alt="Canvas rendering" />
  <img src="https://img.shields.io/badge/workflow-AI%20Ready-00A67E.svg" alt="AI ready" />
  <img src="https://img.shields.io/badge/file%20support-XLSX%20%7C%20CSV%20%7C%20JSON-1f6feb.svg" alt="file support" />
</p>

<p align="center">
  <video src="docs/ys2.mp4" controls muted playsinline preload="metadata" poster="docs/image_en.png" width="100%"></video>
</p>

<p align="center">
  <a href="docs/ys2.mp4">Watch demo video</a>
</p>

- SheetNext is a pure front-end, high-performance spreadsheet engine that provides enterprises with a ready-to-use intelligent spreadsheet foundation.
- With the AI-driven development approach, a single developer + AI can integrate and deliver complex enterprise spreadsheet solutions.
- Common scenarios like ledgers, budgets, analytics, data entry, and approvals can produce a first version in minutes.

## ✨ Key Features

- 📊 Full Spreadsheet Capabilities — Formula engine, charts, pivot tables, super tables, slicers, conditional formatting, data validation, sparklines, freeze panes, sorting & filtering, and more
- 🤖 AI-Powered Workflow — Built-in AI automation for template generation, data analysis, formula writing, and cross-sheet logic
- 📁 Native File Support — Import/export Excel (.xlsx), CSV, and JSON out of the box, no extra plugins needed
- 🚀 Zero-Config Setup — All features built in, no additional dependencies required
- ⚡ High-Performance Rendering — Canvas-based virtual scrolling handles large datasets with ease

## 🚀 Quick Start

SheetNext can be integrated with just a few lines of code and works with any front-end framework (Vue, React, Angular, etc.).

### Option 1: Traditional Integration

#### Install via npm

```bash
npm install sheetnext
```

```html
<!-- Container for the editor -->
<div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
```

```javascript
import SheetNext from 'sheetnext';
import 'sheetnext.css';

const SN = new SheetNext(document.querySelector('#SNContainer'));
```

#### Browser Direct Import (CDN)

```html
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>SheetNext Demo</title>
  <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sheetnext@0.2.0/dist/sheetnext.css">
</head>
<body>
  <div id="SNContainer" style="width:100vw;height:100vh;padding:0 7px 7px"></div>
  <script src="https://cdn.jsdelivr.net/npm/sheetnext@0.2.0/dist/sheetnext.umd.js"></script>
  <script>
    const SN = new SheetNext(document.querySelector('#SNContainer'));
  </script>
</body>
</html>
```

#### Internationalization (i18n)

The default language is English (en-US). A built-in Chinese (zh-CN) locale is available.

Import locale via npm:

```javascript
import SheetNext from 'sheetnext';
import zhCN from 'sheetnext/locales/zh-CN.js';

SheetNext.registerLocale('zh-CN', zhCN);

const SN = new SheetNext(document.querySelector('#SNContainer'), {
  locale: 'zh-CN'
});
```

Import locale via CDN:

```html
<script src="https://cdn.jsdelivr.net/npm/sheetnext@0.2.0/dist/sheetnext.umd.js"></script>
<script src="https://cdn.jsdelivr.net/npm/sheetnext@0.2.0/dist/sheetnext.locale.zh-CN.umd.js"></script>
<script>
  const SN = new SheetNext(document.querySelector('#SNContainer'), {
    locale: 'zh-CN'
  });
</script>
```

### Option 2: AI-Driven Development (Recommended)

#### Step 1: Download the AI Development Reference

- Open `docs/detail/` in the repository.
- The core references are `docs/detail/core-api.md`, `docs/detail/events.md`, and `docs/detail/enums.md`.
- Protocol supplements are `docs/detail/ai-relay.md` and `docs/detail/json-format.md`.

#### Step 2: Feed `docs/detail` to Your AI Tool

Use Cursor / Claude / ChatGPT / Copilot or any AI coding assistant. Provide the `docs/detail/` reference set first, then describe your requirements.

Recommended prompt template:

```text
You are a senior SheetNext AI development expert. Please read and understand the documentation I provide, then give a directly implementable solution.
Execution order:
1) Read: docs/detail/core-api.md
2) Read as needed: docs/detail/events.md, docs/detail/enums.md, docs/detail/ai-relay.md, and docs/detail/json-format.md
3) Identify user goals (business goals + technical goals)
4) Output a minimum viable implementation (get it running first, then optimize)
5) All APIs and code must strictly follow the documentation
6) Provide verification steps and risk points
Constraints:
- Do not fabricate APIs
- Do not skip edge cases
- Prioritize reusing existing capabilities, avoid over-engineering
```

#### Step 3: Describe Your Business Goal

For example:

- "Build a sales pivot analysis template with charts and slicers"
- "Build a multi-sheet budget entry system with permissions and printing"
- "Migrate an existing Excel template to an online editable version"

## 🎯 Use Cases

- Online reporting systems, BI analytics front-ends, business dashboards
- Spreadsheet engine modules in ERP / CRM / Finance / Supply Chain systems
- Complex business forms for budgets, settlements, reconciliation, planning, and scheduling
- AI-powered scenarios: auto-generate tables, analysis, templates, and logic

## Browser Support

| Chrome | Firefox | Safari | Edge |
|--------|---------|--------|------|
| 80+ | 75+ | 13+ | 80+ |

## License

Apache-2.0. See [LICENSE](./LICENSE).
