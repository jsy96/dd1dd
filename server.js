const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const os = require('os');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

const app = express();
const PORT = process.env.PORT || 5000;

// 中间件
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.json());

// 临时目录（兼容 Windows 和 Unix）
const tmpDir = path.join(os.tmpdir(), 'manifest-uploads');
const outputDir = path.join(os.tmpdir(), 'manifest-output');

// 确保临时目录存在
fs.mkdir(tmpDir, { recursive: true }).catch(() => {});
fs.mkdir(outputDir, { recursive: true }).catch(() => {});

// 配置文件上传
const upload = multer({
  dest: tmpDir,
  limits: { fileSize: 10 * 1024 * 1024 } // 10MB 限制
});

// 解析舱单 Excel 文件
function parseManifestExcel(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];

  const jsonData = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: ''
  });

  const getCellValue = (row, col) => {
    if (row < 0 || row >= jsonData.length) return '';
    const rowData = jsonData[row];
    if (!rowData || col < 0 || col >= rowData.length) return '';
    return String(rowData[col] || '').trim();
  };

  // 1. 预配舱单：固定位置读取
  const common = {
    船名: getCellValue(3, 1),
    航次: getCellValue(3, 4),
    目的港: getCellValue(3, 7),
    总提单号: getCellValue(4, 1),
  };

  // 2. 搜索"明细品名及数据"区域
  let detailTitleRow = -1;
  for (let r = 0; r < jsonData.length; r++) {
    if (String(jsonData[r][0] || '').includes('明细品名及数据')) {
      detailTitleRow = r;
      break;
    }
  }
  if (detailTitleRow === -1) {
    throw new Error('找不到"明细品名及数据"区域');
  }

  // 3. 解析表头行，确定各字段所在列号
  const headerRowIndex = detailTitleRow + 1;
  const headerRow = jsonData[headerRowIndex] || [];
  const colMap = {};
  const headerFieldMap = {
    '提单号': '提单号',
    '箱号': '箱号',
    '封号': '封号',
    '箱型': '箱型',
    '英文品名': '英文品名',
    '件数': '件数',
    '包装单位': '包装单位',
    '毛重(KGS)': '毛重',
    '体积(CBM)': '体积',
    '唛头': '唛头',
  };
  headerRow.forEach((val, idx) => {
    const name = String(val || '').trim();
    if (headerFieldMap[name]) {
      colMap[headerFieldMap[name]] = idx;
    }
  });

  // 4. 读取数据行
  const items = [];
  for (let r = headerRowIndex + 1; r < jsonData.length; r++) {
    const row = jsonData[r];
    const firstCell = String(row[0] || '').trim();
    if (firstCell === '') break;
    if (firstCell.includes('VGM') || firstCell.includes('发货人') || firstCell.includes('收货人') || firstCell.includes('通知人')) break;

    items.push({
      提单号: getCellValue(r, colMap['提单号']),
      箱号: getCellValue(r, colMap['箱号']),
      封号: getCellValue(r, colMap['封号']),
      箱型: getCellValue(r, colMap['箱型']),
      英文品名: getCellValue(r, colMap['英文品名']),
      件数: getCellValue(r, colMap['件数']),
      包装单位: getCellValue(r, colMap['包装单位']),
      毛重: getCellValue(r, colMap['毛重']),
      体积: getCellValue(r, colMap['体积']),
      唛头: getCellValue(r, colMap['唛头']),
    });
  }

  if (items.length === 0) {
    throw new Error('"明细品名及数据"区域没有数据行');
  }

  console.log('DEBUG parseManifestExcel: 总提单号:', common.总提单号, '明细品名行数:', items.length);
  items.forEach((item, i) => {
    console.log(`  item[${i}]: 提单号=${item.提单号}, 英文品名=${item.英文品名}`);
  });

  return { common, items };
}

// 生成 Word 文档
async function generateWordDocument(data) {
  const templatePath = path.join(__dirname, 'templates', '提单确认件的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);
  
  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // 商品列表 - 严格按舱单文件中的英文品名数量处理
  const englishNames = data.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const goodsData = {};
  for (let i = 1; i <= 22; i++) {
    // 只使用舱单文件中存在的商品，不存在则设置为空字符串
    goodsData[`商品${i}`] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 调试日志
  console.log('DEBUG Word生成: 英文品名原始值:', JSON.stringify(data.英文品名));
  console.log('DEBUG Word生成: 解析后商品列表:', JSON.stringify(goodsList));
  console.log('DEBUG Word生成: 商品数据:', JSON.stringify(goodsData));

  doc.setData({
    船名: data.船名,
    航次: data.航次,
    目的港: data.目的港,
    提单号: data.提单号,
    箱号: data.箱号,
    封号: data.封号,
    箱型: data.箱型,
    发货人: data.发货人,
    收货人: data.收货人,
    通知人: data.通知人,
    件数: data.件数,
    毛重: data.毛重,
    体积: data.体积,
    公司名: data.发货人名称,
    公司地址: data.发货人地址,
    电话: data.发货人电话,
    传真: '',
    电子邮箱: '',
    许可证号: '',
    收货地址: data.收货人地址,
    邮编: '',
    手机号: '',
    电话号码: data.收货人电话,
    姓名: data.通知人名称,
    地址: data.通知人地址,
    ...goodsData,
  });

  doc.render();
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// 生成 Excel 文档
async function generateExcelDocument(data) {
  const templatePath = path.join(__dirname, 'templates', '装箱单发票的格式.xlsx');
  const templateBuffer = await fs.readFile(templatePath);
  
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);
  const worksheet = workbook.worksheets[0];

  if (!worksheet) {
    throw new Error('无法加载 Excel 模板');
  }

  // 生成日期
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  // 获取单元格文本
  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

  // 替换占位符
  const replacePlaceholder = (cell, placeholder, replacement) => {
    const text = getCellText(cell);
    if (text.includes(placeholder)) {
      let font = {};
      if (cell.value?.richText && cell.value.richText.length > 0) {
        font = cell.value.richText[0].font || {};
      }
      cell.value = {
        richText: [{
          font: font,
          text: replacement
        }]
      };
      return true;
    }
    return false;
  };

  // 填充发票日期
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell) => {
      replacePlaceholder(cell, '{发票日期}', formattedDate);
    });
  });

  // 填充商品列表 - 严格按舱单文件中的英文品名数量处理
  const englishNames = data.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  for (let i = 0; i < 22; i++) {
    const rowNum = 12 + i;
    const row = worksheet.getRow(rowNum);
    const cell = row.getCell(5);
    const placeholder = `{商品${i + 1}}`;

    // 只使用舱单文件中存在的商品，不存在则设置为空字符串
    const goodsValue = i < goodsList.length ? goodsList[i] : '';
    replacePlaceholder(cell, placeholder, goodsValue);
  }

  // 调试日志
  console.log('DEBUG Excel生成: 英文品名原始值:', JSON.stringify(data.英文品名));
  console.log('DEBUG Excel生成: 解析后商品列表:', JSON.stringify(goodsList));

  return workbook.xlsx.writeBuffer();
}

// API: 处理舱单文件
app.post('/api/process', upload.single('manifest'), async (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ success: false, message: '请上传舱单文件' });
    }

    // 读取上传的文件
    const manifestBuffer = await fs.readFile(req.file.path);

    // 解析舱单数据
    const parsed = parseManifestExcel(manifestBuffer);

    const results = [];
    // 处理每个明细行
    for (let j = 0; j < parsed.items.length; j++) {
      const item = parsed.items[j];
      const cargoData = { ...parsed.common, ...item };

      const wordBuffer = await generateWordDocument(cargoData);
      const excelBuffer = await generateExcelDocument(cargoData);

      const timestamp = Date.now();
      const safeBillName = cargoData.提单号 ? cargoData.提单号.replace(/[^a-zA-Z0-9]/g, '_') : `bill_${j + 1}`;
      const wordFileName = `提单确认件_${safeBillName}_${timestamp}.doc`;
      const excelFileName = `装箱单发票_${safeBillName}_${timestamp}.xls`;

      const wordFilePath = path.join(outputDir, wordFileName);
      const excelFilePath = path.join(outputDir, excelFileName);

      await fs.writeFile(wordFilePath, wordBuffer);
      await fs.writeFile(excelFilePath, excelBuffer);

      results.push({
        提单号: cargoData.提单号,
        wordFileUrl: `/api/download?file=${encodeURIComponent(wordFileName)}`,
        excelFileUrl: `/api/download?file=${encodeURIComponent(excelFileName)}`,
      });
    }

    // 清理上传的临时文件
    await fs.unlink(req.file.path).catch(() => {});

    res.json({
      success: true,
      message: `文件处理成功，共生成 ${results.length} 组文件`,
      data: parsed,
      results,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    res.status(500).json({ success: false, message: '处理文件失败，请检查文件格式' });
  }
});

// API: 重新生成文件
app.post('/api/regenerate', async (req, res) => {
  try {
    const cargoData = req.body.data;

    if (!cargoData) {
      return res.status(400).json({ success: false, message: '缺少数据' });
    }

    // 生成文件
    const wordBuffer = await generateWordDocument(cargoData);
    const excelBuffer = await generateExcelDocument(cargoData);

    // 保存文件
    const timestamp = Date.now();
    const wordFileName = `提单确认件_${timestamp}.doc`;
    const excelFileName = `装箱单发票_${timestamp}.xls`;
    
    const wordFilePath = path.join(outputDir, wordFileName);
    const excelFilePath = path.join(outputDir, excelFileName);

    await fs.writeFile(wordFilePath, wordBuffer);
    await fs.writeFile(excelFilePath, excelBuffer);

    res.json({
      success: true,
      message: '文件重新生成成功',
      wordFileUrl: `/api/download?file=${encodeURIComponent(wordFileName)}`,
      excelFileUrl: `/api/download?file=${encodeURIComponent(excelFileName)}`,
    });
  } catch (error) {
    console.error('重新生成文件失败:', error);
    res.status(500).json({ success: false, message: '重新生成文件失败' });
  }
});

// API: 下载文件
app.get('/api/download', async (req, res) => {
  try {
    const filename = req.query.file;
    if (!filename) {
      return res.status(400).json({ success: false, message: '缺少文件名' });
    }

    const filePath = path.join(outputDir, filename);

    // 检查文件是否存在
    try {
      await fs.access(filePath);
    } catch {
      console.error('文件不存在:', filePath);
      return res.status(404).json({ success: false, message: '文件不存在' });
    }

    // 读取文件
    const fileBuffer = await fs.readFile(filePath);

    // 根据文件扩展名设置 MIME 类型
    const ext = path.extname(filename).toLowerCase();
    let contentType = 'application/octet-stream';
    if (ext === '.doc' || ext === '.docx') {
      contentType = 'application/msword';
    } else if (ext === '.xls' || ext === '.xlsx') {
      contentType = 'application/vnd.ms-excel';
    }

    // 设置响应头（兼容手机浏览器）
    res.setHeader('Content-Type', contentType);
    res.setHeader('Content-Length', fileBuffer.length);

    // Content-Disposition 支持 UTF-8 文件名 (RFC 5987)
    const encodedFileName = encodeURIComponent(filename);
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"; filename*=UTF-8''${encodedFileName}`);

    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Accept-Ranges', 'bytes');

    res.send(fileBuffer);
  } catch (error) {
    console.error('下载文件失败:', error);
    res.status(500).json({ success: false, message: '下载文件失败' });
  }
});

// 启动服务器
app.listen(PORT, '0.0.0.0', () => {
  console.log(`服务器已启动: http://localhost:${PORT}`);
  console.log(`模板目录: ${path.join(__dirname, 'templates')}`);
});
