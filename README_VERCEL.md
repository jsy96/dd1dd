# 部署到 Vercel 指南

本项目是一个舱单文件处理系统，支持从飞书表格查询HS编码功能。以下是将项目部署到 Vercel 的完整步骤。

## 项目结构

```
.
├── api/process.js              # Vercel 无服务器函数（主处理逻辑）
├── hs-encoder.js              # HS编码查询模块（飞书表格集成）
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

### 3. 配置环境变量

在 Vercel 项目设置中设置以下环境变量：

| 变量名 | 说明 | 示例值 |
|--------|------|--------|
| `FEISHU_APP_ID` | 飞书应用 ID | `cli_a9c794dadcb9dcc5` |
| `FEISHU_APP_SECRET` | 飞书应用密钥 | `w9JwfI4t4EiJsppHNjNSLcvx1EkLdau7` |
| `FEISHU_BASE_URL` | （可选）飞书多维表格链接 | `https://my.feishu.cn/base/OoNybRydGaN6Wwspy41cnQQCnGe?table=tbl2uWivrvboRe2a` |

**如何获取飞书应用凭证：**
1. 访问 [飞书开放平台](https://open.feishu.cn/)
2. 创建企业自建应用
3. 在「凭证与基础信息」中获取 App ID 和 App Secret
4. 为应用添加「多维表格」权限

### 4. 部署配置

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

### 5. 开始部署

点击 "Deploy" 按钮开始部署。Vercel 会自动：
- 安装 Node.js 依赖
- 构建项目
- 部署无服务器函数和静态文件

### 6. 测试部署

部署完成后，访问 Vercel 提供的域名（如 `https://your-project.vercel.app`）测试功能：

1. **前端界面**：`https://your-project.vercel.app`
2. **API 端点**：`https://your-project.vercel.app/api/process`

## 功能验证

### HS 编码查询功能验证

部署后，HS 编码查询功能应正常工作：

1. **自动查询**：处理舱单文件时，系统会自动查询飞书表格中的 HS 编码
2. **自动创建**：如果商品不存在，会自动在飞书表格中创建新记录（HS 编码为 12345678）
3. **错误处理**：如果飞书 API 不可用，会使用默认 HS 编码 12345678

### 文件处理功能验证

上传舱单文件（.xls/.xlsx 格式）后，系统会生成：
- 提单确认件（Word 文档）
- 装箱单发票（Excel 文档）
- 并单保函（Word 文档）
- 总提单 OK 件（带 HS 编码）
- 总提单 OK 件（无 HS 编码）

所有文件打包为 ZIP 格式下载。

## 常见问题

### 1. 飞书 API 权限问题

**错误**：`飞书API错误: 无权限访问该数据表`

**解决**：
1. 在飞书开放平台为应用添加「多维表格」权限
2. 将应用发布到企业
3. 在飞书多维表格中分享表格给该应用

### 2. 环境变量未生效

**错误**：`缺少飞书应用配置`

**解决**：
1. 在 Vercel 项目设置中确认环境变量已正确设置
2. 重新部署项目
3. 检查环境变量名称是否正确

### 3. 模板文件找不到

**错误**：`无法加载 Excel 模板`

**解决**：
1. 确保 `templates/` 目录中的模板文件存在
2. 检查文件路径是否正确
3. 模板文件名必须与代码中的引用一致

### 4. 文件大小限制

Vercel 无服务器函数有 50MB 的部署包限制。本项目模板文件较小，不会超过此限制。

## 本地开发

如需在本地测试 Vercel 无服务器函数：

```bash
# 安装依赖
npm install

# 设置环境变量
cp .env.example .env
# 编辑 .env 文件填入真实值

# 启动开发服务器（使用原始 Express 服务器）
npm run dev

# 或者使用 Vercel CLI 测试无服务器函数
npm i -g vercel
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
3. 飞书开放平台应用权限

项目保持原有功能不变，HS 编码查询功能完全保留。