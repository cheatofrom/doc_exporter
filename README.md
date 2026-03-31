# Word 报告自动生成系统 (Word Report Generation System)

这是一个基于 React (前端) 和 Node.js + Express (后端) 的全栈应用，用于通过前端动态表单收集数据，并将其导出填充到指定的 Word (`.docx`) 模板中。

## ✨ 核心特性 (Features)

- **动态复杂表单**：支持多页报告数据收集、面板折叠与展开。
- **动态表格增删行**：支持附录（如样件信息、零部件信息、试验时间、设备列表等）的动态表格行添加与删除。
- **图片自动插入**：支持在前端上传图片（转为 Base64），后端自动将其按指定尺寸渲染到 Word 模板的对应位置（如：附录 B3 样件照片，附录 C3 设备备注图片）。
- **健壮的模板引擎**：基于 `docxtemplater`，采用 `[[ ]]` 作为自定义占位符边界，有效避免了 Word 底层 XML 标签将默认 `{{ }}` 拆分导致解析失败的问题。

## 📁 目录结构 (Project Structure)

本系统采用 npm workspaces 组织为一个 Monorepo：

```text
.
├── frontend/             # 前端项目 (React + TypeScript + Vite)
│   ├── src/
│   │   ├── App.tsx       # 核心 UI 及表单状态逻辑
│   │   ├── types.ts      # 接口与表单字段类型定义
│   │   └── styles.css    # 界面样式
│   └── package.json
├── backend/              # 后端项目 (Node.js + Express + TypeScript)
│   ├── src/
│   │   ├── index.ts      # Express 服务入口及路由
│   │   ├── template.ts   # 模板渲染逻辑与图片处理器配置
│   │   └── types.ts      # 后端类型定义
│   ├── templates/
│   │   └── test_word.docx # 默认的 Word 报告模板文件
│   └── package.json
└── package.json          # 根目录配置，统筹 workspaces 及启动脚本
```

## 🚀 快速开始 (Getting Started)

### 1. 安装依赖

在项目根目录下执行：

```bash
npm install
```

### 2. 启动后端服务

由于默认端口可能被占用，建议指定 `39011` 等端口启动：

```bash
PORT=39011 npm run dev:backend
```

后端服务启动后，默认运行在 `http://localhost:39011`。

### 3. 启动前端服务

在新的终端窗口中，运行：

```bash
npm run dev:frontend
```

如果需要指定端口，可执行类似于 `npm run dev --workspace frontend -- --port 5176` 的命令。

### 4. 使用说明

1. 打开前端页面。
2. 按需填写第一页、第二页以及各个附录（B1, B2, B3, B4, C, D）的表单内容。
3. 表格部分支持点击 `+ 添加行` 和上传照片。
4. 填写完毕后，点击最下方的 **"导出 Word"** 按钮。
5. 系统会自动下载名为 `报告导出_xxxx.docx` 的文件。

## 📝 模板占位符语法 (Template Syntax)

为了防止 Word 文档将花括号拆分，本系统将 `docxtemplater` 的界定符改为了双中括号 `[[ ]]`。在修改 `backend/templates/test_word.docx` 时，请遵循以下规则：

- **普通文本填充**：`[[field_name]]` (例如：`[[report_no]]`, `[[test_leader]]`)
- **多行文本填充**：直接在文本框输入多行内容，Word 模板中使用 `[[test_steps]]` 即可，系统已开启 `linebreaks: true`。
- **表格/循环渲染**：
  ```word
  [[#equipment_list]]
  设备名称: [[name]]
  设备编号: [[equip_no]]
  [[/equipment_list]]
  ```
- **图片渲染**：`[[%remark_image]]` (前缀 `%` 表示这是一个图片模块标签，后台需要接收 Base64 格式的字符串)。

## 🛠️ 常见问题 (Troubleshooting)

- **MulterError: Field value too long**
  *原因*：上传的高清图片转 Base64 后超出了 Express 默认大小限制。
  *解决*：已在 `backend/src/index.ts` 中将 `express.json` 和 `multer` 的 `limit` 大小配置为了 `50mb`。
- **TemplateError: Duplicate open tag**
  *原因*：使用 `{{ }}` 时，Word 的拼写检查或格式变更导致 XML 将大括号拆开了。
  *解决*：必须使用 `[[ ]]` 作为占位符。若还报错，建议用纯文本编辑器或者 Python 脚本清洗该段落，或者删除该占位符后以“无格式”方式重新手打一遍。
