# 🧭 学术罗盘 Office 插件 (Academic Compass Office Add-in)

这是一个专为 **Microsoft Word** 设计的任务窗格 (Task Pane) 插件。它将 AI 驱动的学术与职业生涯分析功能无缝集成到 Word 界面中，帮助用户在撰写简历或学术文档时，随时获取职业规划洞察。

This is a Task Pane Add-in designed specifically for **Microsoft Word**. It seamlessly integrates AI-powered academic and career analysis features into the Word interface, allowing users to obtain career planning insights while drafting resumes or academic documents.

## 核心功能 / Core Features

* **深度集成 Word / Deep Word Integration:** 插件作为 Word 的侧边任务窗格运行，并通过 Office.js 确保在 Word 环境下正常运行。
* **简历/专业分析 / Resume & Major Analysis:** 用户可以直接在任务窗格内输入或粘贴专业、技能和简历文本，一键启动生涯分析。
* **多语言 UI / Multilingual UI:** 界面支持简体中文、繁体中文和英文，并能保存用户选择的语言设置。
* **可调整面板 / Resizable Panel:** 任务窗格 UI 具备拖动分割线以调整输入和输出区域高度的功能，优化用户体验。
* **AI 报告与引用 / AI Reporting & Citation:** 插件连接到 Academic Compass 后端 API，获取 Gemini 生成的结构化报告，并使用 Marked.js 和 DOMPurify 安全地渲染报告内容和引用来源。
* **设置持久化 / Settings Persistence:** 使用 Office.js 的 `document.settings` 和 `localStorage` 来保存主题、语言和面板高度等设置。

## 技术栈 / Tech Stack

| 模块 / Module | 组件 / Component | 描述 / Description |
| :--- | :--- | :--- |
| **平台 / Platform** | Microsoft Office Add-in, Office.js | 任务窗格运行环境和与 Word 宿主应用交互的 API。/ Task Pane environment and API for interaction with the Word host. |
| **构建工具 / Build Tools** | Webpack, Babel | 用于打包和转译 JavaScript 代码，确保兼容旧版 Office 运行时环境。/ Used to bundle and transpile JavaScript for compatibility with older Office runtimes. |
| **UI 基础 / UI Foundation** | HTML, CSS, Vanilla JavaScript | 任务窗格界面的构建。/ Building the Task Pane UI. |
| **报告处理 / Report Handling** | Marked.js, DOMPurify | 客户端 Markdown 渲染和 HTML 安全净化。/ Client-side Markdown rendering and HTML sanitization. |
| **后端通信 / Backend Communication**| Fetch API | 用于调用外部部署的 Academic Compass 后端服务。/ Used to call the external Academic Compass backend service. |

## 安装与部署 / Installation and Deployment (Sideloading)

此插件通过标准 Office 插件清单文件 (`manifest.xml`) 进行安装和部署。/ This add-in is installed and deployed via the standard Office Add-in manifest file (`manifest.xml`).

1.  **准备环境 / Prerequisites:** 需要安装 Node.js 和 Office Add-in 开发工具。/ Requires Node.js and Office Add-in development tools.
2.  **构建 / Build:** 运行 `npm run build` 或 `npm run build:dev` 使用 Webpack 生成最终的 `taskpane.html` 和 `taskpane.js` 等文件。
3.  **旁加载 / Sideloading:** 使用 Office Add-in 工具链，通过 `manifest.xml` 文件在 Word 中进行本地调试和旁加载。/ Use the Office Add-in tooling to sideload and debug the plugin in Word using `manifest.xml`.
