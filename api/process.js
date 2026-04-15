// Vercel serverless function for processing manifest files
const path = require('path');
const fs = require('fs').promises;
const stream = require('stream');
const XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');
const archiver = require('archiver');

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

  const data = {
    船名: getCellValue(3, 1),
    航次: getCellValue(3, 4),
    目的港: getCellValue(3, 7),
    提单号: getCellValue(4, 1),
    箱号: getCellValue(12, 0),
    封号: getCellValue(12, 1),
    箱型: getCellValue(12, 2),
    英文品名: getCellValue(20, 4),
    件数: getCellValue(20, 6),
    包装单位: getCellValue(20, 7),
    毛重: getCellValue(20, 8),
    体积: getCellValue(20, 9),
    唛头: getCellValue(20, 10),
    发货人名称: getCellValue(27, 2),
    发货人地址: getCellValue(28, 2),
    发货人电话: getCellValue(30, 2),
    收货人名称: getCellValue(34, 2),
    收货人地址: getCellValue(35, 2),
    收货人电话: getCellValue(37, 2),
    收货人联系人: getCellValue(39, 2),
    通知人名称: getCellValue(43, 2),
    通知人地址: getCellValue(44, 2),
    通知人电话: getCellValue(46, 2),
  };

  data.发货人 = [
    data.发货人名称,
    data.发货人地址,
    `TEL: ${data.发货人电话}`
  ].filter(Boolean).join('\n');

  data.收货人 = [
    data.收货人名称,
    data.收货人地址,
    `TEL: ${data.收货人电话}`,
    data.收货人联系人 ? `Contact: ${data.收货人联系人}` : ''
  ].filter(Boolean).join('\n');

  data.通知人 = [
    data.通知人名称,
    data.通知人地址,
    `TEL: ${data.通知人电话}`
  ].filter(Boolean).join('\n');

  // 调试日志
  console.log('DEBUG parseManifestExcel: 英文品名原始值:', JSON.stringify(data.英文品名));
  console.log('DEBUG parseManifestExcel: 解析后商品列表长度:', data.英文品名 ? data.英文品名.split(',').map(s => s.trim()).filter(item => item !== '').length : 0);

  return data;
}

// 生成 Word 文档
async function generateWordDocument(data) {
  const templatePath = path.join(__dirname, '../templates/提单确认件的格式.docx');
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
  const templatePath = path.join(__dirname, '../templates/装箱单发票的格式.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

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

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = data.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  console.log('替换数据:', {
    英文品名: data.英文品名,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    占位符数量: Object.keys(replacementData).length,
    占位符列表: Object.keys(replacementData)
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    // 遍历所有行和单元格
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        // 尝试替换所有可能的占位符
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });
    console.log(`Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符`);
  });

  // 专门处理 PACKING LIST 工作表的 B11-B32 单元格
  const packingListSheet = workbook.getWorksheet('PACKING LIST') || workbook.worksheets[1];
  if (packingListSheet) {
    console.log(`专门处理 PACKING LIST 工作表: ${packingListSheet.name}`);
    let specificReplaced = 0;
    for (let row = 11; row <= 32; row++) {
      const cell = packingListSheet.getCell(`B${row}`);
      const placeholderIndex = row - 10; // B11 对应商品1, B12 对应商品2...
      const placeholder = `{商品${placeholderIndex}}`;
      const replacement = replacementData[placeholder] || '';

      if (replacePlaceholder(cell, placeholder, replacement)) {
        specificReplaced++;
      } else {
        // 如果没找到占位符，直接设置单元格值
        const cellText = getCellText(cell);
        if (cellText.includes('{商品')) {
          // 单元格包含其他商品占位符，尝试替换所有可能的占位符
          for (const [ph, repl] of Object.entries(replacementData)) {
            if (replacePlaceholder(cell, ph, repl)) {
              specificReplaced++;
              break;
            }
          }
        }
      }
    }
    console.log(`PACKING LIST 工作表 B11-B32 替换了 ${specificReplaced} 个单元格`);
  }

  return workbook.xlsx.writeBuffer();
}

// 生成并单保函 Word 文档
async function generateCombinedLetter(firstData, allCargoData) {
  const templatePath = path.join(__dirname, '../templates/并单保函的格式.docx');
  const templateBuffer = await fs.readFile(templatePath);

  const zip = new PizZip(templateBuffer);
  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  // 商品列表 - 严格按舱单文件中的英文品名数量处理（使用第一个文件的数据）
  const englishNames = firstData.英文品名 || '';
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

  // 生成所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, ...
  const containerData = {};
  const maxContainers = 20; // 假设模板最多支持20个舱单

  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      containerData[`提单号${suffix}`] = cargo.提单号 || '';
      containerData[`箱号${suffix}`] = cargo.箱号 || '';
      containerData[`箱型${suffix}`] = cargo.箱型 || '';
      containerData[`封号${suffix}`] = cargo.封号 || '';
      containerData[`件数${suffix}`] = cargo.件数 || '';
      containerData[`毛重${suffix}`] = cargo.毛重 || '';
      containerData[`体积${suffix}`] = cargo.体积 || '';
    } else {
      // 填充空的占位符
      containerData[`提单号${suffix}`] = '';
      containerData[`箱号${suffix}`] = '';
      containerData[`箱型${suffix}`] = '';
      containerData[`封号${suffix}`] = '';
      containerData[`件数${suffix}`] = '';
      containerData[`毛重${suffix}`] = '';
      containerData[`体积${suffix}`] = '';
    }
  }

  console.log('并单保函舱单数据（前3个）:', {
    提单号1: containerData['提单号1'],
    箱号1: containerData['箱号1'],
    箱型1: containerData['箱型1'],
    封号1: containerData['封号1'],
    件数1: containerData['件数1'],
    毛重1: containerData['毛重1'],
    体积1: containerData['体积1'],
  });

  doc.setData({
    船名: firstData.船名,
    航次: firstData.航次,
    目的港: firstData.目的港,
    提单号: firstData.提单号,
    箱号: firstData.箱号,
    封号: firstData.封号,
    箱型: firstData.箱型,
    件数: firstData.件数,
    毛重: firstData.毛重,
    体积: firstData.体积,
    公司名: firstData.发货人名称,
    公司地址: firstData.发货人地址,
    电话: firstData.发货人电话,
    传真: '',
    电子邮箱: '',
    许可证号: '',
    收货地址: firstData.收货人地址,
    邮编: '',
    手机号: '',
    电话号码: firstData.收货人电话,
    姓名: firstData.通知人名称,
    地址: firstData.通知人地址,
    ...goodsData,
    ...containerData,
  });

  doc.render();
  return doc.getZip().generate({ type: 'nodebuffer' });
}

// 生成总提单OK件（带HS）Excel 文档
async function generateOKBillWithHS(firstData, allCargoData) {
  const templatePath = path.join(__dirname, '../templates/总提单OK件的格式(带HS的.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  // 使用与现有 Excel 生成相同的替换逻辑
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

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

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = firstData.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate,
    '{船名}': firstData.船名 || '',
    '{航次}': firstData.航次 || '',
    '{目的港}': firstData.目的港 || '',
    '{提单号}': firstData.提单号 || '',
    '{箱号}': firstData.箱号 || '',
    '{封号}': firstData.封号 || '',
    '{箱型}': firstData.箱型 || '',
    '{件数}': firstData.件数 || '',
    '{毛重}': firstData.毛重 || '',
    '{体积}': firstData.体积 || '',
    '{发货人名称}': firstData.发货人名称 || '',
    '{发货人地址}': firstData.发货人地址 || '',
    '{发货人电话}': firstData.发货人电话 || '',
    '{收货人名称}': firstData.收货人名称 || '',
    '{收货人地址}': firstData.收货人地址 || '',
    '{收货人电话}': firstData.收货人电话 || '',
    '{通知人名称}': firstData.通知人名称 || '',
    '{通知人地址}': firstData.通知人地址 || '',
    '{通知人电话}': firstData.通知人电话 || '',
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 添加所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, ...
  const maxContainers = 20; // 假设模板最多支持20个舱单
  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      replacementData[`{提单号${suffix}}`] = cargo.提单号 || '';
      replacementData[`{箱号${suffix}}`] = cargo.箱号 || '';
      replacementData[`{箱型${suffix}}`] = cargo.箱型 || '';
      replacementData[`{封号${suffix}}`] = cargo.封号 || '';
      replacementData[`{件数${suffix}}`] = cargo.件数 || '';
      replacementData[`{毛重${suffix}}`] = cargo.毛重 || '';
      replacementData[`{体积${suffix}}`] = cargo.体积 || '';
    } else {
      // 填充空的占位符
      replacementData[`{提单号${suffix}}`] = '';
      replacementData[`{箱号${suffix}}`] = '';
      replacementData[`{箱型${suffix}}`] = '';
      replacementData[`{封号${suffix}}`] = '';
      replacementData[`{件数${suffix}}`] = '';
      replacementData[`{毛重${suffix}}`] = '';
      replacementData[`{体积${suffix}}`] = '';
    }
  }

  console.log('总提单OK件（带HS）替换数据:', {
    提单号: firstData.提单号,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    提单号总数: allCargoData.length,
    所有提单号: allCargoData.map(d => d.提单号),
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });
    console.log(`总提单OK件（带HS） Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符`);
  });

  return workbook.xlsx.writeBuffer();
}

// 生成总提单OK件（无HS）Excel 文档
async function generateOKBillWithoutHS(firstData, allCargoData) {
  const templatePath = path.join(__dirname, '../templates/总提单OK件的格式(无HS的.xlsx');
  const templateBuffer = await fs.readFile(templatePath);

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBuffer);

  if (workbook.worksheets.length === 0) {
    throw new Error('无法加载 Excel 模板');
  }

  // 使用与现有 Excel 生成相同的替换逻辑
  const today = new Date();
  const months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC'];
  const formattedDate = `${months[today.getMonth()]}. ${String(today.getDate()).padStart(2, '0')}. ${today.getFullYear()}`;

  const getCellText = (cell) => {
    if (!cell.value) return '';
    if (typeof cell.value === 'string') return cell.value;
    if (cell.value.richText) {
      return cell.value.richText.map(rt => rt.text || '').join('');
    }
    return '';
  };

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

  // 准备替换数据 - 严格按舱单文件中的英文品名数量处理
  const englishNames = firstData.英文品名 || '';
  const goodsList = englishNames.split(',').map(s => s.trim()).filter(item => item !== '');
  // 确保商品数量不超过22个，如果超过则截断并记录警告
  if (goodsList.length > 22) {
    console.warn(`警告：舱单文件中有 ${goodsList.length} 个英文品名，但模板只支持22个商品。将截断超出的部分。`);
  }
  const replacementData = {
    '{发票日期}': formattedDate,
    '{船名}': firstData.船名 || '',
    '{航次}': firstData.航次 || '',
    '{目的港}': firstData.目的港 || '',
    '{提单号}': firstData.提单号 || '',
    '{箱号}': firstData.箱号 || '',
    '{封号}': firstData.封号 || '',
    '{箱型}': firstData.箱型 || '',
    '{件数}': firstData.件数 || '',
    '{毛重}': firstData.毛重 || '',
    '{体积}': firstData.体积 || '',
    '{发货人名称}': firstData.发货人名称 || '',
    '{发货人地址}': firstData.发货人地址 || '',
    '{发货人电话}': firstData.发货人电话 || '',
    '{收货人名称}': firstData.收货人名称 || '',
    '{收货人地址}': firstData.收货人地址 || '',
    '{收货人电话}': firstData.收货人电话 || '',
    '{通知人名称}': firstData.通知人名称 || '',
    '{通知人地址}': firstData.通知人地址 || '',
    '{通知人电话}': firstData.通知人电话 || '',
  };

  // 添加商品占位符替换数据 - 只使用舱单文件中存在的商品
  for (let i = 1; i <= 22; i++) {
    const placeholder = `{商品${i}}`;
    replacementData[placeholder] = i <= goodsList.length ? goodsList[i - 1] : '';
  }

  // 添加所有舱单字段映射：提单号1, 箱号1, 箱型1, 封号1, 件数1, 毛重1, 体积1, ...
  const maxContainers = 20; // 假设模板最多支持20个舱单
  for (let i = 0; i < maxContainers; i++) {
    const suffix = i + 1;
    if (i < allCargoData.length) {
      const cargo = allCargoData[i];
      replacementData[`{提单号${suffix}}`] = cargo.提单号 || '';
      replacementData[`{箱号${suffix}}`] = cargo.箱号 || '';
      replacementData[`{箱型${suffix}}`] = cargo.箱型 || '';
      replacementData[`{封号${suffix}}`] = cargo.封号 || '';
      replacementData[`{件数${suffix}}`] = cargo.件数 || '';
      replacementData[`{毛重${suffix}}`] = cargo.毛重 || '';
      replacementData[`{体积${suffix}}`] = cargo.体积 || '';
    } else {
      // 填充空的占位符
      replacementData[`{提单号${suffix}}`] = '';
      replacementData[`{箱号${suffix}}`] = '';
      replacementData[`{箱型${suffix}}`] = '';
      replacementData[`{封号${suffix}}`] = '';
      replacementData[`{件数${suffix}}`] = '';
      replacementData[`{毛重${suffix}}`] = '';
      replacementData[`{体积${suffix}}`] = '';
    }
  }

  console.log('总提单OK件（无HS）替换数据:', {
    提单号: firstData.提单号,
    商品列表长度: goodsList.length,
    商品列表内容: goodsList,
    提单号总数: allCargoData.length,
    所有提单号: allCargoData.map(d => d.提单号),
  });

  // 处理所有 sheet
  workbook.worksheets.forEach((worksheet, sheetIndex) => {
    let replacedCount = 0;
    worksheet.eachRow((row, rowNumber) => {
      row.eachCell((cell) => {
        for (const [placeholder, replacement] of Object.entries(replacementData)) {
          if (replacePlaceholder(cell, placeholder, replacement)) {
            replacedCount++;
          }
        }
      });
    });
    console.log(`总提单OK件（无HS） Sheet ${sheetIndex + 1} "${worksheet.name}" 替换了 ${replacedCount} 个占位符`);
  });

  return workbook.xlsx.writeBuffer();
}

// Buffer to base64
function bufferToBase64(buffer) {
  return buffer.toString('base64');
}

module.exports = async (req, res) => {
  // Handle CORS preflight
  if (req.method === 'OPTIONS') {
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ success: false, message: 'Method not allowed' });
  }

  try {
    const formData = await parseFormData(req);

    const manifestFiles = formData.files.manifest;
    if (!manifestFiles || (Array.isArray(manifestFiles) && manifestFiles.length === 0) || manifestFiles.length === 0) {
      return res.status(400).json({ success: false, message: '请上传舱单文件' });
    }

    // 确保是数组
    const files = Array.isArray(manifestFiles) ? manifestFiles : [manifestFiles];

    console.log(`开始批量处理 ${files.length} 个文件`);

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

    res.setHeader('Access-Control-Allow-Origin', '*');
    res.json({
      success: true,
      message: `批量处理完成，共处理 ${files.length} 个文件`,
      zipFileBase64: zipBase64,
      fileCount: files.length,
    });
  } catch (error) {
    console.error('处理文件失败:', error);
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.status(500).json({ success: false, message: '处理文件失败：' + error.message });
  }
};

// Helper to parse form data
async function parseFormData(req) {
  const Busboy = require('busboy');
  return new Promise((resolve, reject) => {
    const busboy = Busboy({ headers: req.headers });
    const files = {};
    const fields = {};

    busboy.on('file', (fieldname, file, info) => {
      const chunks = [];
      file.on('data', (chunk) => chunks.push(chunk));
      file.on('end', () => {
        const fileData = { buffer: Buffer.concat(chunks), ...info };
        if (!files[fieldname]) {
          files[fieldname] = [fileData];
        } else if (Array.isArray(files[fieldname])) {
          files[fieldname].push(fileData);
        } else {
          // 如果已存在但不是数组，转换为数组
          files[fieldname] = [files[fieldname], fileData];
        }
      });
    });

    busboy.on('field', (fieldname, value) => {
      fields[fieldname] = value;
    });

    busboy.on('finish', () => resolve({ files, fields }));
    busboy.on('error', reject);

    req.pipe(busboy);
  });
}

module.exports.parseManifestExcel = parseManifestExcel;
module.exports.generateWordDocument = generateWordDocument;
module.exports.generateExcelDocument = generateExcelDocument;
module.exports.generateCombinedLetter = generateCombinedLetter;
module.exports.generateOKBillWithHS = generateOKBillWithHS;
module.exports.generateOKBillWithoutHS = generateOKBillWithoutHS;
