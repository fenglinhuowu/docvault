# DocVault

完全离线、跨平台的文档管理、编辑与转换软件（轻量版 WPS）。

## 特性

- 完全离线，数据不上传
- 跨平台：Windows / macOS / iOS / Android
- Word 文档编辑与转 PDF
- Excel 表格处理（集成 Luckysheet/FortuneSheet）
- PDF 查看（集成 PDF.js）

## 技术栈

- 前端：Vue 3 + TypeScript + Vite
- 后端：Tauri v2.0 + Rust

## 快速开始

```bash
# 安装依赖
npm install

# 开发模式
npm run tauri:dev

# 构建发布版
npm run tauri:build
```

## 项目结构

```
docvault/
├── src/                    # 前端 (Vue3 + TypeScript)
│   ├── App.vue             # 主应用组件
│   ├── main.ts             # 入口文件
│   ├── components/         # 公共组件
│   ├── views/              # 页面视图
│   ├── stores/             # 状态管理
│   └── composables/        # 组合式函数
├── src-tauri/              # Rust 后端 (Tauri)
│   ├── Cargo.toml          # Rust 依赖
│   ├── src/main.rs         # Rust 核心代码
│   ├── tauri.conf.json     # Tauri 配置
│   └── capabilities/       # 权限配置
├── package.json
└── vite.config.ts
```

## 核心功能

### Rust Tauri Commands

| Command | 说明 |
|---------|------|
| `open_local_file` | 离线读取本地文档（Base64） |
| `save_local_file` | 离线保存文档到本地 |
| `convert_docx_to_pdf` | Word 转 PDF（完全离线） |
| `scan_directory` | 扫描目录获取文件列表 |
| `get_file_info` | 获取文件详细信息 |

## License

MIT
