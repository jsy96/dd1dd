# 部署到 Vercel 指南

本项目是一个舱单文件处理系统。以下是将项目部署到 Vercel 的完整步骤。

## 项目结构

```
.
├── api/process.js              # Vercel 无服务器函数（主处理逻辑）
├── public/index.html          # 前端界面
├── templates/                 # Word/Excel 模板文件
├── package.json              # 依赖配置
├── vercel.json               # Vercel 部署配置
├── .env.example              # 环境变量示例
└── README_VERCEL.md          # 本文件
```

## 部署步骤

### 1. 准备代码仓库

将代码推送到 GitHub、GitLab 或 Bitbucket 仓库。

### 2. 在 Vercel 中导入项目

1. 访问 [Vercel](https://vercel.com) 并登录
2. 点击 "Add New" → "Project"
3. 导入你的代码仓库
4. Vercel 会自动检测项目配置

### 3. 部署配置

项目已包含 `vercel.json` 配置文件，无需额外设置：

```json
{
  "builds": [
    {
      "src": "api/*.js",
      "use": "@vercel/node"
    },
    {
      "src": "public/*",
      "use": "@vercel/static"
    }
  ],
  "routes": [
    {
      "src": "/api/(.*)",
      "dest": "/api/$1"
    },
    {
      "src": "/(.*)",
      "dest": "/public/$1"
    }
  ],
  "headers": [
    {
      "source": "/api/(.*)",
      "headers": [
        {
          "key": "Access-Control-Allow-Origin",
          "value": "*"
        }
      ]
    }
  ]
}
```

### 4. 开始部署

点击 "Deploy" 按钮开始部署。Vercel 会自动：
- 安装 Node.js 依赖
- 构建项目
- 部署无服务器函数和静态文件

### 5. 测试部署

部署完成后，访问 Vercel 提供的域名（如 `https://your-project.vercel.app`）测试功能：

1. **前端界面**：`https://your-project.vercel.app`
2. **API 端点**：`https://your-project.vercel.app/api/process`

## 功能验证

### HS 编码处理

部署后，所有商品的 HS 编码统一使用固定值 `12345678`。

### 文件处理功能验证

上传舱单文件（.xls/.xlsx 格式）后，系统会生成：
- 提单确认件（Word 文档）
- 装箱单发票（Excel 文档）
- 并单保函（Word 文档）
- 总提单 OK 件（带 HS 编码）
- 总提单 OK 件（无 HS 编码）

所有文件打包为 ZIP 格式下载。

## 常见问题

### 1. 模板文件找不到

**错误**：`无法加载 Excel 模板`

**解决**：
1. 确保 `templates/` 目录中的模板文件存在
2. 检查文件路径是否正确
3. 模板文件名必须与代码中的引用一致

### 2. 文件大小限制

Vercel 无服务器函数有 50MB 的部署包限制。本项目模板文件较小，不会超过此限制。

## 本地开发

### 解决常见本地测试问题

#### 1. Tailwind CSS CDN 警告
已修复：将 `cdn.tailwindcss.com` 替换为 `unpkg.com/tailwindcss` 稳定版本。

#### 2. API 404 错误
在本地直接打开 `index.html` 文件会导致 `/api/process` 请求失败。请使用本地开发服务器：

```bash
# 安装依赖
npm install

# 启动本地开发服务器（模拟 Vercel 环境）
npm run dev

# 访问 http://localhost:3000
```

#### 3. JSON 解析错误
如果 API 返回 HTML 错误页面而不是 JSON，请确保：
1. 使用本地开发服务器（而非直接打开 HTML 文件）
2. 检查浏览器控制台错误信息

### 本地开发服务器功能
- 提供静态文件服务（`public/` 目录）
- 提供 API 端点 `/api/process`（与 Vercel 无服务器函数兼容）
- 支持多文件上传和 ZIP 打包下载
- 健康检查端点 `/api/health`
- 模板文件调试端点 `/api/templates`

### 使用 Vercel CLI 测试（可选）
```bash
# 全局安装 Vercel CLI
npm i -g vercel

# 在项目目录中启动 Vercel 开发服务器
vercel dev
```

## 更新部署

代码更新后，Vercel 会自动重新部署。如需手动触发：

1. 在 Vercel 控制台选择项目
2. 点击 "Redeploy"
3. 或推送代码到连接的仓库

## 技术支持

如遇部署问题，请检查：
1. Vercel 部署日志
2. 浏览器开发者工具控制台

项目保持原有功能不变，所有商品的 HS 编码统一使用 12345678。
