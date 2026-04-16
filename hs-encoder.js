/**
 * 将无HS的商品列表转换为有HS的商品列表
 * 输入：逗号分隔的英文品名字符串
 * 输出：每个英文品名后加空格和8位HS编码，逗号分隔的字符串
 * HS编码来自飞书多维表格，如果不存在则添加新记录（HS编码填充为12345678）
 *
 * 使用说明：
 * 1. 需要先配置飞书多维表格链接（直接链接或appToken和tableId）
 * 2. 可以通过以下方式使用：
 *    - 直接调用 addHSCodes() 函数
 *    - 使用 lark-cli 技能手动查询和创建记录
 */

const path = require('path');

/**
 * 解析飞书多维表格的直接链接，提取 app_token 和 table_id
 * @param {string} url - 飞书多维表格直接链接，格式：https://{domain}.feishu.cn/base/{appToken}?table={tableId}
 * @returns {object} { appToken, tableId, viewId }
 */
function parseFeishuBaseUrl(url) {
  try {
    const urlObj = new URL(url);
    const pathParts = urlObj.pathname.split('/');
    const appToken = pathParts[pathParts.length - 1];
    const tableId = urlObj.searchParams.get('table');
    const viewId = urlObj.searchParams.get('view');
    return { appToken, tableId, viewId };
  } catch (error) {
    throw new Error('无效的飞书多维表格链接格式。请提供类似 https://yourdomain.feishu.cn/base/{appToken}?table={tableId} 的链接');
  }
}

/**
 * 解析飞书Wiki链接，提取 wiki_token 和 table_id
 * @param {string} url - 飞书Wiki链接，格式：https://{domain}.feishu.cn/wiki/{wikiToken}?table={tableId}
 * @returns {object} { wikiToken, tableId, viewId, isWikiLink: true }
 */
function parseFeishuWikiUrl(url) {
  try {
    const urlObj = new URL(url);
    const pathParts = urlObj.pathname.split('/');
    const wikiToken = pathParts[pathParts.length - 1];
    const tableId = urlObj.searchParams.get('table');
    const viewId = urlObj.searchParams.get('view');
    return { wikiToken, tableId, viewId, isWikiLink: true };
  } catch (error) {
    throw new Error('无效的飞书Wiki链接格式。请提供类似 https://yourdomain.feishu.cn/wiki/{wikiToken}?table={tableId} 的链接');
  }
}

/**
 * 智能解析飞书链接，自动判断是Base链接还是Wiki链接
 * @param {string} url - 飞书链接
 * @returns {object} 解析结果
 */
function parseFeishuUrl(url) {
  if (!url || typeof url !== 'string') {
    throw new Error('链接不能为空');
  }

  if (url.includes('/wiki/')) {
    return parseFeishuWikiUrl(url);
  } else if (url.includes('/base/')) {
    return parseFeishuBaseUrl(url);
  } else {
    throw new Error('无法识别的飞书链接格式，必须是/base/或/wiki/链接');
  }
}

/**
 * 模拟查询飞书多维表格记录（实际使用时需要替换为lark-cli调用）
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {string} fieldName - 查询字段名（默认为"英文品名"）
 * @param {string} value - 查询值
 * @returns {object|null} 记录数据，如果不存在返回 null
 */
function queryFeishuRecord(appToken, tableId, fieldName, value) {
  // 这里是模拟实现，实际使用时需要调用 lark-cli 技能
  // lark-cli base list-records {appToken} {tableId} --filter '{fieldName}="{value}"' --output json
  console.log(`模拟查询: ${fieldName} = "${value}"`);
  console.log(`实际应执行: lark-cli base list-records ${appToken} ${tableId} --filter '${fieldName}="${value}"' --output json`);
  return null; // 模拟返回空，表示记录不存在
}

/**
 * 模拟在飞书多维表格中添加新记录（实际使用时需要替换为lark-cli调用）
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {object} fields - 字段键值对，例如 { "英文品名": "Apple", "HS编码": "12345678" }
 * @returns {object} 新创建的记录
 */
function createFeishuRecord(appToken, tableId, fields) {
  // 这里是模拟实现，实际使用时需要调用 lark-cli 技能
  // lark-cli base create-record {appToken} {tableId} --fields '{fieldsJson}' --output json
  const fieldsJson = JSON.stringify(fields);
  console.log(`模拟创建记录: ${fieldsJson}`);
  console.log(`实际应执行: lark-cli base create-record ${appToken} ${tableId} --fields '${fieldsJson}' --output json`);
  return { record_id: 'mock_record_id', fields: fields };
}

/**
 * 获取商品的HS编码，如果不存在则创建新记录
 * @param {string} goodsName - 英文品名
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @returns {string} 8位HS编码
 */
async function getOrCreateHSCode(goodsName, appToken, tableId) {
  // 查询是否存在该商品
  const record = queryFeishuRecord(appToken, tableId, '英文品名', goodsName);
  if (record) {
    // 从记录中提取HS编码字段
    const hsCode = record.fields['HS编码'] || record.fields['hs_code'] || '';
    if (hsCode && hsCode.toString().length === 8) {
      return hsCode.toString();
    } else {
      console.warn(`商品"${goodsName}"的HS编码格式不正确: ${hsCode}，使用默认值12345678`);
      return '12345678';
    }
  } else {
    // 创建新记录
    console.log(`商品"${goodsName}"不存在于飞书多维表格中，添加新记录...`);
    const fields = {
      '英文品名': goodsName,
      'HS编码': '12345678'
    };
    const newRecord = createFeishuRecord(appToken, tableId, fields);
    console.log(`新记录添加成功，记录ID: ${newRecord.record_id}`);
    return '12345678';
  }
}

/**
 * 主函数：将无HS的商品列表转换为有HS的商品列表
 * @param {string} goodsListWithoutHS - 逗号分隔的英文品名字符串，例如 "Apple, Banana, Orange"
 * @param {object} options - 配置选项
 * @param {string} options.feishuBaseUrl - 飞书多维表格直接链接（必需）
 * @param {string} options.goodsNameField - 商品名字段名，默认为"英文品名"
 * @param {string} options.hsCodeField - HS编码字段名，默认为"HS编码"
 * @returns {Promise<string>} 有HS的商品列表字符串，例如 "Apple 12345678, Banana 87654321, Orange 12345678"
 */
async function addHSCodes(goodsListWithoutHS, options = {}) {
  if (!goodsListWithoutHS || typeof goodsListWithoutHS !== 'string') {
    return '';
  }

  // 解析飞书多维表格链接
  if (!options.feishuBaseUrl) {
    throw new Error('缺少飞书多维表格链接，请提供 options.feishuBaseUrl');
  }
  const { appToken, tableId } = parseFeishuBaseUrl(options.feishuBaseUrl);
  const goodsNameField = options.goodsNameField || '英文品名';
  const hsCodeField = options.hsCodeField || 'HS编码';

  // 分割商品列表
  const goodsNames = goodsListWithoutHS.split(',')
    .map(name => name.trim())
    .filter(name => name.length > 0);

  if (goodsNames.length === 0) {
    return '';
  }

  // 处理每个商品
  const results = [];
  for (const goodsName of goodsNames) {
    // 获取或创建HS编码
    const hsCode = await getOrCreateHSCode(goodsName, appToken, tableId);
    results.push(`${goodsName} ${hsCode}`);
  }

  // 重新组合为逗号分隔的字符串
  return results.join(', ');
}

/**
 * 使用lark-cli技能的完整工作流程示例
 * @param {string} goodsListWithoutHS - 逗号分隔的英文品名字符串
 * @param {string} feishuBaseUrl - 飞书多维表格直接链接
 * @returns {string} 转换后的商品列表和操作指南
 */
function getWorkflowInstructions(goodsListWithoutHS, feishuUrl) {
  const parsedUrl = parseFeishuUrl(feishuUrl);
  const goodsNames = goodsListWithoutHS.split(',')
    .map(name => name.trim())
    .filter(name => name.length > 0);

  let instructions = `# HS编码转换工作流程\n\n`;

  if (parsedUrl.isWikiLink) {
    const { wikiToken, tableId, viewId } = parsedUrl;
    instructions += `飞书Wiki链接解析:\n`;
    instructions += `- wikiToken: ${wikiToken}\n`;
    instructions += `- tableId: ${tableId}\n`;
    if (viewId) instructions += `- viewId: ${viewId}\n`;
    instructions += `\n注意：Wiki链接需要先解析为Base链接才能操作。使用以下命令获取真实baseToken:\n`;
    instructions += `\`lark-cli wiki spaces get_node ${wikiToken} --output json\`\n`;
    instructions += `从返回结果的 \`obj_token\` 字段获取baseToken（如果 \`obj_type\` 是 \`bitable\`）。\n\n`;
    instructions += `## 需要处理的商品列表:\n`;
  } else {
    const { appToken, tableId, viewId } = parsedUrl;
    instructions += `飞书多维表格: appToken=${appToken}, tableId=${tableId}\n`;
    if (viewId) instructions += `viewId: ${viewId}\n`;
    instructions += `\n## 需要处理的商品列表:\n`;
  }

  goodsNames.forEach((name, index) => {
    instructions += `${index + 1}. ${name}\n`;
  });

  instructions += `\n## 操作步骤:\n\n`;

  if (parsedUrl.isWikiLink) {
    // Wiki链接需要先解析
    instructions += `### 1. 解析Wiki链接获取baseToken\n`;
    instructions += `先执行: \`lark-cli wiki spaces get_node ${parsedUrl.wikiToken} --output json\`\n`;
    instructions += `假设返回的baseToken为 \`bas_xxx\`，tableId为 \`${parsedUrl.tableId}\`\n\n`;

    instructions += `### 2. 查询现有记录\n`;
    goodsNames.forEach((name, index) => {
      instructions += `商品${index + 1} "${name}":\n`;
      instructions += `\`lark-cli base list-records bas_xxx ${parsedUrl.tableId} --filter '英文品名="${name}"' --output json\`\n\n`;
    });

    instructions += `### 3. 创建不存在的记录\n`;
    instructions += `对于查询不到的商品，使用以下命令创建（HS编码为12345678）：\n`;
    instructions += `\`lark-cli base create-record bas_xxx ${parsedUrl.tableId} --fields '{"英文品名":"商品名称","HS编码":"12345678"}' --output json\`\n\n`;
  } else {
    // 直接Base链接
    const { appToken, tableId } = parsedUrl;

    instructions += `### 1. 查询现有记录\n`;
    goodsNames.forEach((name, index) => {
      instructions += `商品${index + 1} "${name}":\n`;
      instructions += `\`lark-cli base list-records ${appToken} ${tableId} --filter '英文品名="${name}"' --output json\`\n\n`;
    });

    instructions += `### 2. 创建不存在的记录\n`;
    instructions += `对于查询不到的商品，使用以下命令创建（HS编码为12345678）：\n`;
    instructions += `\`lark-cli base create-record ${appToken} ${tableId} --fields '{"英文品名":"商品名称","HS编码":"12345678"}' --output json\`\n\n`;
  }

  // 结果组装
  instructions += `### ${parsedUrl.isWikiLink ? '4' : '3'}. 组装最终结果\n`;
  instructions += `将每个商品名称和对应的HS编码（8位数字）组合成: "商品名称 12345678"\n`;
  instructions += `然后用逗号分隔所有商品: "商品1 12345678, 商品2 87654321, ..."\n`;

  return instructions;
}

// 导出函数
module.exports = {
  addHSCodes,
  parseFeishuBaseUrl,
  parseFeishuWikiUrl,
  parseFeishuUrl,
  queryFeishuRecord,
  createFeishuRecord,
  getOrCreateHSCode,
  getWorkflowInstructions
};

// 如果直接运行，提供使用说明
if (require.main === module) {
  console.log(`
HS编码转换工具

使用方法:
1. 直接调用函数:
   const { addHSCodes } = require('./hs-encoder');
   const result = await addHSCodes("Apple, Banana, Orange", {
     feishuBaseUrl: "https://yourdomain.feishu.cn/base/appToken?table=tableId"
   });

2. 使用lark-cli技能手动操作:
   const { getWorkflowInstructions } = require('./hs-encoder');
   const instructions = getWorkflowInstructions("Apple, Banana, Orange", "https://yourdomain.feishu.cn/base/appToken?table=tableId");
   console.log(instructions);

注意事项:
- 需要先安装并配置 lark-cli 或使用 Claude Code 中的 lark-base 技能
- 飞书多维表格中应有"英文品名"和"HS编码"字段（或相应字段名）
- 如果商品不存在，会自动创建新记录，HS编码为12345678
`);
}