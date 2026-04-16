// 本地开发服务器
// 用于在本地测试项目，模拟Vercel无服务器函数环境

const express = require('express');
const path = require('path');
const fs = require('fs').promises;

// 导入api/process.js中的函数
const apiProcess = require('./api/process.js');

const app = express();
const PORT = process.env.PORT || 3000;

// 中间件
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json({ limit: '50mb' }));

// 提供templates目录的静态访问（用于调试）
app.use('/templates', express.static(path.join(__dirname, 'templates')));

// API路由 - 处理舱单文件
app.post('/api/process', async (req, res) => {
  try {
    // 设置CORS头
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    // 使用api/process.js中的parseFormData函数处理multipart/form-data
    const { parseFormData } = require('./api/process.js');
    const formData = await parseFormData(req);

    const manifestFiles = formData.files.manifest;
    if (!manifestFiles || (Array.isArray(manifestFiles) && manifestFiles.length === 0) || manifestFiles.length === 0) {
      return res.status(400).json({ success: false, message: '请上传舱单文件' });
    }

    // 确保是数组
    const files = Array.isArray(manifestFiles) ? manifestFiles : [manifestFiles];

    console.log(`开始批量处理 ${files.length} 个文件`);

    // 导入api/process.js中的其他函数
    const { parseManifestExcel, generateWordDocument, generateExcelDocument, generateCombinedLetter, generateOKBillWithHS, generateOKBillWithoutHS, bufferToBase64 } = require('./api/process.js');
    const archiver = require('archiver');
    const stream = require('stream');

    // 收集所有舱单数据用于生成汇总文件
    const allCargoData = [];
    let firstCargoData = null;

    // 创建 ZIP 归档
    const archive = archiver('zip', { zlib: { level: 9 } });
    const chunks = [];
    archive.on('data', (chunk) => chunks.push(chunk));

    const zipPromise = new Promise((resolve, reject) => {
      archive.on('end', () => {
        const zipBuffer = Buffer.concat(chunks);
        resolve(zipBuffer);
      });
      archive.on('error', reject);
    });

    // 处理每个文件
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      console.log(`处理第 ${i + 1} 个文件: ${file.filename || '未命名文件'}`);

      try {
        const cargoData = parseManifestExcel(file.buffer);

        // 收集数据用于汇总文件
        allCargoData.push(cargoData);
        if (i === 0) {
          firstCargoData = cargoData;
        }

        const wordBuffer = await generateWordDocument(cargoData);
        const excelBuffer = await generateExcelDocument(cargoData);

        // 生成安全文件名
        const safeBillName = cargoData.提单号 ? cargoData.提单号.replace(/[^a-zA-Z0-9]/g, '_') : `bill_${i + 1}`;
        const safeContainerName = cargoData.箱号 ? cargoData.箱号.replace(/[^a-zA-Z0-9]/g, '_') : `container_${i + 1}`;

        // 添加到 ZIP，按文件夹结构组织
        archive.append(wordBuffer, { name: `A/B/${safeBillName}.docx` });
        archive.append(excelBuffer, { name: `A/C/${safeContainerName}.xlsx` });

        console.log(`文件 ${i + 1} 处理完成: 提单号=${cargoData.提单号}, 箱号=${cargoData.箱号}`);
        console.log(`  文件名: B/${safeBillName}.docx, C/${safeContainerName}.xlsx`);
      } catch (fileError) {
        console.error(`处理文件 ${i + 1} 失败:`, fileError);
        // 跳过失败的文件，继续处理其他文件
        continue;
      }
    }

    // 生成三个汇总文件（使用第一个文件的数据）
    if (firstCargoData) {
      try {
        const safeBillNumber = firstCargoData.提单号 ? firstCargoData.提单号.replace(/[^a-zA-Z0-9]/g, '_') : '汇总';

        // 生成并单保函
        const combinedLetterBuffer = await generateCombinedLetter(firstCargoData, allCargoData);
        archive.append(combinedLetterBuffer, { name: `A/${safeBillNumber}并单保函的格式.docx` });
        console.log(`生成汇总文件: A/${safeBillNumber}并单保函的格式.docx`);

        // 生成总提单OK件（带HS）
        const okWithHSBuffer = await generateOKBillWithHS(firstCargoData, allCargoData);
        archive.append(okWithHSBuffer, { name: `A/${safeBillNumber}总提单OK件的格式(带HS的.xlsx` });
        console.log(`生成汇总文件: A/${safeBillNumber}总提单OK件的格式(带HS的.xlsx`);

        // 生成总提单OK件（无HS）
        const okWithoutHSBuffer = await generateOKBillWithoutHS(firstCargoData, allCargoData);
        archive.append(okWithoutHSBuffer, { name: `A/${safeBillNumber}总提单OK件的格式(无HS的.xlsx` });
        console.log(`生成汇总文件: A/${safeBillNumber}总提单OK件的格式(无HS的.xlsx`);
      } catch (summaryError) {
        console.error('生成汇总文件失败，跳过:', summaryError);
      }
    }

    // 完成 ZIP 归档
    await archive.finalize();
    const zipBuffer = await zipPromise;
    const zipBase64 = bufferToBase64(zipBuffer);

    res.json({
      success: true,
      message: `批量处理完成，共处理 ${files.length} 个文件`,
      zipFileBase64: zipBase64,
      fileCount: files.length,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    res.status(500).json({ success: false, message: '处理文件失败：' + error.message });
  }
});

// OPTIONS预检请求
app.options('/api/process', (req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*');
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  res.status(200).end();
});

// 健康检查端点
app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    timestamp: new Date().toISOString(),
    environment: process.env.NODE_ENV || 'development'
  });
});

// 模板文件列表（用于调试）
app.get('/api/templates', async (req, res) => {
  try {
    const templatesDir = path.join(__dirname, 'templates');
    const files = await fs.readdir(templatesDir);
    res.json({
      success: true,
      files: files.map(file => ({
        name: file,
        path: `/templates/${file}`,
        size: '-'
      }))
    });
  } catch (error) {
    res.status(500).json({
      success: false,
      message: '无法读取模板文件'
    });
  }
});

// 主页
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 启动服务器
app.listen(PORT, '0.0.0.0', () => {
  console.log(`本地开发服务器已启动: http://localhost:${PORT}`);
  console.log(`API端点: http://localhost:${PORT}/api/process`);
  console.log(`健康检查: http://localhost:${PORT}/api/health`);
  console.log(`模板文件: http://localhost:${PORT}/api/templates`);
  console.log(`\n注意：此服务器仅用于本地测试，生产环境请使用Vercel部署。`);
  console.log(`环境变量检查:`);
  console.log(`  - FEISHU_APP_ID: ${process.env.FEISHU_APP_ID ? '已设置' : '未设置'}`);
  console.log(`  - FEISHU_APP_SECRET: ${process.env.FEISHU_APP_SECRET ? '已设置' : '未设置'}`);
  console.log(`  - FEISHU_BASE_URL: ${process.env.FEISHU_BASE_URL ? '已设置' : '未设置'}`);
});