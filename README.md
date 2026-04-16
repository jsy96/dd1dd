# 舱单文件处理系统

这是一个基于 Express 的 Web 应用，用于处理舱单文件并自动生成提单确认件和装箱单发票。

## 功能特性

- 📤 **文件上传**：支持上传舱单 Excel 文件
- 📊 **智能解析**：自动提取舱单中的关键信息
- ✏️ **数据编辑**：支持查看和修改提取的数据
- 🔄 **重新生成**：修改数据后可重新生成文档
- 📄 **文档生成**：
  - 生成提单确认件（Word 格式）
  - 生成装箱单发票（Excel 格式）
- 📥 **文件下载**：一键下载生成的文档

## 技术栈

- **Express** - Node.js Web 框架
- **Multer** - 文件上传中间件
- **XLSX** - Excel 文件解析
- **ExcelJS** - Excel 文档生成
- **Docxtemplater** - Word 文档生成
- **PizZip** - ZIP 文件处理（用于 Word 模板）

## 快速开始

### 安装依赖

```bash
pnpm install
```

### 启动服务器

```bash
pnpm start
```

或使用开发模式：

```bash
pnpm dev
```

服务器默认运行在 [http://localhost:5000](http://localhost:5000)

## 使用说明

### 1. 上传舱单文件

在首页点击"舱单文件"区域，上传舱单的格式.xls文件。

### 2. 开始处理

点击"开始处理"按钮，系统会自动：
1. 解析舱单 Excel 文件
2. 提取关键信息
3. 使用内置模板生成 Word 和 Excel 文档

### 3. 查看数据

处理完成后，在"提取数据"标签页可以查看：
- **基本信息**：船名、航次、目的港、提单号、箱号、封号、箱型
- **货物信息**：英文品名、件数、毛重、体积、唛头
- **发货人信息**：名称、地址、电话
- **收货人信息**：名称、地址、电话、联系人
- **通知人信息**：名称、地址、电话

### 4. 编辑数据（可选）

点击"编辑数据"按钮，可以修改提取的数据：
- 修改任意字段的内容
- 点击"重新生成"按钮生成新文档

### 5. 下载文件

在"下载文件"标签页下载：
- 提单确认件.doc
- 装箱单发票.xls

## 舱单格式支持

系统支持标准预配舱单格式，自动提取以下数据：

| 字段 | 说明 |
|------|------|
| 船名 | 船舶名称 |
| 航次 | 航次编号 |
| 目的港 | 卸货港口 |
| 提单号 | 提单号码 |
| 箱号 | 集装箱号 |
| 封号 | 铅封号 |
| 箱型 | 集装箱类型（如40HQ） |
| 英文品名 | 商品英文名称（逗号分隔） |
| 件数 | 货物件数 |
| 毛重 | 货物毛重（KGS） |
| 体积 | 货物体积（CBM） |
| 发货人 | 发货人信息 |
| 收货人 | 收货人信息 |
| 通知人 | 通知人信息 |

## 模板文件

系统已预置以下模板文件：
- `templates/提单确认件的格式.docx` - 提单确认件模板
- `templates/装箱单发票的格式.xlsx` - 装箱单发票模板

无需用户上传，系统会自动使用这些模板生成文档。

## API 接口

### POST /api/process

处理舱单文件并生成文档

**请求参数：**
- `manifest`: 舱单 Excel 文件（multipart/form-data）

**响应：**
```json
{
  "success": true,
  "message": "文件处理成功",
  "data": { /* 提取的数据 */ },
  "wordFileUrl": "/api/download?file=提单确认件_xxx.doc",
  "excelFileUrl": "/api/download?file=装箱单发票_xxx.xls"
}
```

### POST /api/regenerate

根据修改的数据重新生成文档

**请求参数：**
```json
{
  "data": { /* 舱单数据 */ }
}
```

**响应：**
```json
{
  "success": true,
  "wordFileUrl": "/api/download?file=提单确认件_xxx.doc",
  "excelFileUrl": "/api/download?file=装箱单发票_xxx.xls"
}
```

### GET /api/download

下载生成的文件

**查询参数：**
- `file`: 文件名

## 项目结构

```
├── public/                  # 静态文件
│   └── index.html           # 前端页面
├── templates/               # 文档模板
│   ├── 提单确认件的格式.docx
│   ├── 舱单的格式.xls
│   ├── 装箱单发票的格式.xls
│   └── 装箱单发票的格式.xlsx
├── server.js                # Express 后端服务器
├── package.json             # 项目配置
└── README.md                # 项目文档
```

## 环境变量

- `PORT`: 服务器端口（默认：5000）

## 开发规范

**必须使用 pnpm 作为包管理器**

```bash
# ✅ 正确
pnpm install
pnpm add <package>

# ❌ 错误
npm install
yarn add
```

## 依赖说明

主要依赖：
- `express` (^4.21.2) - Web 框架
- `multer` (^1.4.5-lts.1) - 文件上传
- `xlsx` (^0.18.5) - Excel 解析
- `exceljs` (^4.4.0) - Excel 生成
- `docxtemplater` (^3.68.3) - Word 生成
- `pizzip` (^3.2.0) - ZIP 处理

## HS 编码查询功能

系统集成了飞书多维表格查询功能，可自动为商品添加 HS 编码：

1. **自动查询**：处理舱单时，自动查询飞书表格中的 HS 编码
2. **自动创建**：如果商品不存在，自动在飞书表格创建新记录
3. **错误处理**：飞书 API 不可用时使用默认编码 12345678

### 配置说明

1. 设置飞书应用环境变量：
   - `FEISHU_APP_ID`: 飞书应用 ID
   - `FEISHU_APP_SECRET`: 飞书应用密钥
   - `FEISHU_BASE_URL`: （可选）飞书多维表格链接

2. 飞书应用需要「多维表格」权限

### 相关文件
- `hs-encoder.js`: HS 编码查询核心模块
- `api/process.js`: 无服务器函数版本（支持 Vercel 部署）

## Vercel 部署

项目已配置支持 Vercel 无服务器部署：

### 部署方式
1. 将代码推送到 Git 仓库
2. 在 Vercel 中导入项目
3. 配置环境变量（见 `.env.example`）
4. 自动部署

### 项目结构适配
- `api/process.js`: Vercel 无服务器函数
- `public/index.html`: 静态前端界面
- `templates/`: 文档模板（包含在部署包中）

### 详细部署指南
参见 [README_VERCEL.md](README_VERCEL.md) 文件。

## 本地开发与测试

### 常见问题解决

#### 1. Tailwind CSS 警告
已修复：使用稳定的 unpkg Tailwind CSS 版本替代开发 CDN。

#### 2. API 404 错误
不要在浏览器中直接打开 `index.html` 文件。请使用本地开发服务器：

```bash
# 安装依赖
npm install

# 复制环境变量示例文件
cp .env.example .env
# 编辑 .env 文件，填入飞书应用凭证

# 启动本地开发服务器
npm run dev

# 访问 http://localhost:3000
```

#### 3. JSON 解析错误
如果遇到 "Unexpected token" JSON 解析错误：
1. 确保使用本地开发服务器（端口 3000）
2. 检查浏览器控制台网络标签，确认 API 返回 JSON 而非 HTML
3. 验证飞书应用环境变量设置正确

### 本地开发服务器
项目包含 `dev-server.js` 文件，提供完整的本地测试环境：
- 静态文件服务（前端界面）
- API 端点 `/api/process`（处理舱单文件）
- 健康检查 `/api/health`
- 模板文件调试 `/api/templates`

### 测试流程
1. 启动开发服务器：`npm run dev`
2. 访问 http://localhost:3000
3. 上传舱单文件测试处理功能
4. 检查 HS 编码查询是否正常工作（需正确配置飞书应用）
