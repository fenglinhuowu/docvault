# DocVault

完全离线、跨平台的文档管理、编辑与转换软件（轻量版 WPS）。

## 下载

### macOS 安装包

| 文件 | 说明 |
|------|------|
| [DocVault.app](releases/DocVault.app) | 直接运行的应用程序 |
| [DocVault_0.1.0_aarch64.dmg](releases/DocVault_0.1.0_aarch64.dmg) | macOS 安装镜像 |

**安装方式：**

1. **DMG 安装**：双击 `DocVault_0.1.0_aarch64.dmg`，将 DocVault 拖到 Applications 文件夹
2. **直接运行**：双击 `DocVault.app` 即可运行

> **注意**：首次运行需要在 `系统设置 → 隐私与安全性 → 辅助功能` 中授权 DocVault。

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
# 克隆项目
git clone https://github.com/fenglinhuowu/docvault.git
cd docvault

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
├── releases/               # 安装包
│   ├── DocVault.app        # macOS 应用程序
│   └── DocVault_0.1.0_aarch64.dmg  # macOS 安装镜像
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

## Word 功能

| 功能 | 说明 |
|------|------|
| 富文本编辑 | 加粗/斜体/下划线/删除线 |
| 字体设置 | 宋体/黑体/楷体/微软雅黑/Arial 等 |
| 字号设置 | 小五(12px) ~ 一号(36px) |
| 标题格式 | H1/H2/H3 |
| 颜色 | 文字颜色/高亮背景色 |
| 段落格式 | 左对齐/居中/右对齐 |
| 列表 | 有序列表/无序列表 |
| 表格 | 插入表格 |
| 图片 | 插入图片 |
| 撤销重做 | Ctrl+Z / Ctrl+Y |

## License

MIT
