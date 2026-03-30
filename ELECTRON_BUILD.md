# 舱单文件处理系统 - 打包为 Windows 可执行文件

## 项目结构

```
├── electron/                # Electron 相关代码
│   ├── main.js             # 主进程
│   └── preload.js          # 预加载脚本
├── public/                  # 静态文件
│   └── index.html          # 前端页面
├── templates/               # 文档模板
├── server.js                # Express 后端服务器
├── package.json             # 项目配置
└── build-electron.bat       # Windows 打包脚本
```

## 环境要求

- **Node.js** 18.x 或更高版本
- **pnpm** 9.x (推荐) 或 npm
- **Windows** 操作系统 (用于打包 Windows exe)

## 快速开始

### 方法一：使用打包脚本 (推荐)

双击运行 `build-electron.bat` 文件，脚本会自动完成以下步骤：
1. 检查并安装依赖
2. 打包 Electron 应用
3. 生成可执行文件

### 方法二：手动打包

```bash
# 1. 安装依赖
pnpm install

# 2. 打包 Electron 应用 (Windows 64位)
pnpm electron:build:win

# 或者打包其他平台
pnpm electron:build:mac    # macOS
pnpm electron:build:linux  # Linux
```

## 输出文件

打包完成后，在 `dist` 目录下会生成以下文件：

| 文件名 | 说明 |
|--------|------|
| `舱单文件处理系统 Setup 1.0.0.exe` | 安装程序，带安装向导 |
| `舱单文件处理系统 1.0.0.exe` | 便携版，无需安装 |

## 技术栈

- **Express** - Node.js Web 框架
- **Electron** - 桌面应用框架
- **electron-builder** - 打包工具
- **Tailwind CSS** (CDN) - CSS 框架
- **Lucide Icons** (CDN) - 图标库

## 开发模式

```bash
# 启动开发服务器
pnpm dev

# 或使用 Electron
pnpm electron:dev
```

## 自定义图标 (可选)

如需自定义应用图标，请将图标文件放入 `public` 目录：

- **Windows**: `icon.ico` (256x256 推荐)
