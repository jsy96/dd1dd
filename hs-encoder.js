/**
 * 将无HS的商品列表转换为有HS的商品列表
 * 输入：逗号分隔的英文品名字符串
 * 输出：每个英文品名后加空格和8位HS编码，逗号分隔的字符串
 * HS编码来自飞书多维表格，如果不存在则添加新记录（HS编码填充为12345678）
 *
 * 注意：
 * 1. HS编码字段应为文本类型，存储8位字符（可以是数字或包含前导0）
 * 2. 查询时会自动将商品名称转换为大写，以匹配表格中的格式
 * 3. 需要配置有效的飞书应用凭证（FEISHU_APP_ID和FEISHU_APP_SECRET环境变量）
 *
 * 使用说明：
 * 1. 需要先配置飞书多维表格链接（直接链接或appToken和tableId）
 * 2. 可以通过以下方式使用：
 *    - 直接调用 addHSCodes() 函数
 *    - 使用 lark-cli 技能手动查询和创建记录
 */

const path = require('path');
const https = require('https');
const { exec } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);

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
 * HTTP请求辅助函数
 * @param {string} url - 请求URL
 * @param {object} options - 请求选项
 * @param {object} data - 请求数据（POST请求时使用）
 * @returns {Promise<object>} 响应数据
 */
function httpRequest(url, options = {}, data = null) {
  return new Promise((resolve, reject) => {
    const urlObj = new URL(url);
    const reqOptions = {
      hostname: urlObj.hostname,
      port: urlObj.port || 443,
      path: urlObj.pathname + urlObj.search,
      method: options.method || 'GET',
      headers: {
        'Content-Type': 'application/json',
        ...options.headers
      }
    };

    const req = https.request(reqOptions, (res) => {
      let responseData = '';
      res.on('data', (chunk) => {
        responseData += chunk;
      });
      res.on('end', () => {
        try {
          const parsed = JSON.parse(responseData);
          if (parsed.code && parsed.code !== 0) {
            reject(new Error(`飞书API错误: ${parsed.msg || '未知错误'} (code: ${parsed.code})`));
          } else {
            resolve(parsed);
          }
        } catch (error) {
          reject(new Error(`解析飞书API响应失败: ${error.message}`));
        }
      });
    });

    req.on('error', (error) => {
      reject(new Error(`飞书API请求失败: ${error.message}`));
    });

    if (data) {
      req.write(JSON.stringify(data));
    }
    req.end();
  });
}

/**
 * 获取飞书租户访问令牌
 * @param {string} appId - 应用ID
 * @param {string} appSecret - 应用密钥
 * @returns {Promise<string>} 访问令牌
 */
async function getFeishuAccessToken(appId, appSecret) {
  if (!appId || !appSecret) {
    throw new Error('缺少飞书应用配置，请设置FEISHU_APP_ID和FEISHU_APP_SECRET环境变量');
  }

  try {
    const response = await httpRequest('https://open.feishu.cn/open-apis/auth/v3/tenant_access_token', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json; charset=utf-8'
      }
    }, {
      app_id: appId,
      app_secret: appSecret
    });

    if (response.tenant_access_token) {
      return response.tenant_access_token;
    } else {
      throw new Error('飞书API未返回访问令牌');
    }
  } catch (error) {
    throw new Error(`获取飞书访问令牌失败: ${error.message}`);
  }
}

// 缓存访问令牌（简单实现，不考虑过期时间）
let cachedAccessToken = null;
let tokenFetching = false;

/**
 * 确保有有效的飞书访问令牌
 * @returns {Promise<string>} 访问令牌
 */
async function ensureAccessToken() {
  if (cachedAccessToken) {
    return cachedAccessToken;
  }

  if (tokenFetching) {
    // 如果正在获取令牌，等待
    await new Promise(resolve => setTimeout(resolve, 100));
    return ensureAccessToken();
  }

  tokenFetching = true;
  try {
    const appId = process.env.FEISHU_APP_ID;
    const appSecret = process.env.FEISHU_APP_SECRET;

    if (!appId || !appSecret) {
      throw new Error('飞书应用配置未设置，请设置FEISHU_APP_ID和FEISHU_APP_SECRET环境变量');
    }

    cachedAccessToken = await getFeishuAccessToken(appId, appSecret);
    return cachedAccessToken;
  } catch (error) {
    console.error('获取飞书访问令牌失败:', error.message);
    throw error;
  } finally {
    tokenFetching = false;
  }
}

/**
 * 查询飞书多维表格记录
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {string} fieldName - 查询字段名（默认为"英文品名"）
 * @param {string} value - 查询值
 * @param {object} options - 选项
 * @param {boolean} options.useBulkQuery - 是否使用批量查询（获取所有记录后在本地筛选），默认为true
 * @param {boolean} options.useFuzzyMatch - 是否使用模糊匹配，默认为false
 * @returns {Promise<object|null>} 记录数据，如果不存在返回 null
 */
async function queryFeishuRecord(appToken, tableId, fieldName, value, options = {}) {
  const {
    useBulkQuery = true, // 默认使用批量查询，像PowerShell命令那样
    useFuzzyMatch = false // 默认不使用模糊匹配
  } = options;

  // 将查询值转换为大写，以匹配表格中的数据格式
  const upperValue = value.toUpperCase().trim();

  const accessToken = await ensureAccessToken();

  try {
    // 方法1：批量查询（像PowerShell命令那样）- 默认使用此方法
    if (useBulkQuery) {
      console.log(`使用批量查询方法: ${fieldName} = "${upperValue}"`);
      return await queryFeishuRecordBulk(appToken, tableId, fieldName, upperValue, useFuzzyMatch);
    }

    // 方法2：API筛选查询（原有方法）
    console.log(`使用API筛选查询方法: ${fieldName} = "${upperValue}"`);

    // 构建飞书API查询URL
    // 注意：飞书API的筛选参数需要在查询字符串中传递
    const filter = JSON.stringify({
      and: [
        {
          field: fieldName,
          operator: useFuzzyMatch ? 'contains' : 'is',
          value: [upperValue]
        }
      ]
    });

    const response = await httpRequest(
      `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records?filter=${encodeURIComponent(filter)}`,
      {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.data && response.data.items && response.data.items.length > 0) {
      // 返回第一个匹配的记录
      const record = response.data.items[0];
      console.log(`飞书API查询成功: ${fieldName} = "${upperValue}"，找到记录`);
      console.log(`记录字段:`, JSON.stringify(record.fields));
      return {
        record_id: record.record_id,
        fields: record.fields
      };
    } else {
      console.log(`飞书API查询: ${fieldName} = "${upperValue}"，未找到记录`);
      console.log(`API响应:`, JSON.stringify(response).substring(0, 500));
      return null;
    }
  } catch (error) {
    console.error(`飞书API查询失败: ${fieldName} = "${upperValue}"`, error.message);
    console.error(`错误详情:`, error);
    throw error;
  }
}

/**
 * 批量查询飞书多维表格记录（像PowerShell命令那样）
 * 获取所有记录后在本地筛选
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {string} fieldName - 查询字段名
 * @param {string} upperValue - 大写的查询值
 * @param {boolean} useFuzzyMatch - 是否使用模糊匹配
 * @returns {Promise<object|null>} 记录数据，如果不存在返回 null
 */
async function queryFeishuRecordBulk(appToken, tableId, fieldName, upperValue, useFuzzyMatch = false) {
  const accessToken = await ensureAccessToken();
  let pageToken = '';
  let allRecords = [];
  let pageCount = 0;

  try {
    // 分页获取所有记录
    do {
      pageCount++;
      let url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records?page_size=100`;
      if (pageToken) {
        url += `&page_token=${pageToken}`;
      }

      console.log(`批量查询第${pageCount}页: ${url.substring(0, 100)}...`);

      const response = await httpRequest(
        url,
        {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.data && response.data.items) {
        allRecords = allRecords.concat(response.data.items);
        console.log(`第${pageCount}页获取到${response.data.items.length}条记录，总计${allRecords.length}条`);

        // 检查是否有下一页
        pageToken = response.data.page_token || '';
      } else {
        console.log(`第${pageCount}页无数据，响应:`, JSON.stringify(response).substring(0, 200));
        break;
      }

      // 防止无限循环，最多获取10页（1000条记录）
      if (pageCount >= 10) {
        console.log(`达到最大分页限制（10页），停止获取`);
        break;
      }

    } while (pageToken);

    console.log(`批量查询完成，共获取${allRecords.length}条记录`);

    // 在本地筛选记录
    let matchedRecord = null;

    for (const record of allRecords) {
      const fields = record.fields || {};

      // 尝试通过字段名查找
      if (fields[fieldName] !== undefined) {
        const fieldValue = String(fields[fieldName]).toUpperCase().trim();
        let isMatch = false;

        if (useFuzzyMatch) {
          isMatch = fieldValue.includes(upperValue) || upperValue.includes(fieldValue);
        } else {
          isMatch = fieldValue === upperValue;
        }

        if (isMatch) {
          console.log(`找到匹配记录（通过字段名${fieldName}）: ${fieldValue} = ${upperValue}`);
          console.log(`完整字段:`, JSON.stringify(fields));
          matchedRecord = record;
          break;
        }
      }

      // 如果通过字段名没找到，尝试通过字段值匹配（像PowerShell命令那样）
      // PowerShell命令使用索引，我们需要找到哪个字段包含商品名称
      if (!matchedRecord) {
        for (const [key, val] of Object.entries(fields)) {
          const fieldValue = String(val).toUpperCase().trim();
          let isMatch = false;

          if (useFuzzyMatch) {
            isMatch = fieldValue.includes(upperValue) || upperValue.includes(fieldValue);
          } else {
            isMatch = fieldValue === upperValue;
          }

          if (isMatch) {
            console.log(`找到匹配记录（通过字段值）: 字段"${key}"的值"${fieldValue}"匹配"${upperValue}"`);
            console.log(`完整字段:`, JSON.stringify(fields));
            console.log(`注意：商品名称字段可能是"${key}"而不是"${fieldName}"`);
            matchedRecord = record;
            break;
          }
        }
      }

      if (matchedRecord) break;
    }

    if (matchedRecord) {
      return {
        record_id: matchedRecord.record_id,
        fields: matchedRecord.fields
      };
    } else {
      console.log(`批量查询未找到匹配记录: ${fieldName} = "${upperValue}"`);
      console.log(`所有记录的字段名:`, allRecords.length > 0 ? Object.keys(allRecords[0].fields || {}) : '无记录');
      console.log(`前3条记录的字段值:`, allRecords.slice(0, 3).map(r => r.fields));
      return null;
    }

  } catch (error) {
    console.error(`批量查询失败: ${fieldName} = "${upperValue}"`, error.message);
    throw error;
  }
}

/**
 * 在飞书多维表格中添加新记录
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {object} fields - 字段键值对，例如 { "英文品名": "Apple", "HS编码": "12345678" }
 * @returns {Promise<object>} 新创建的记录
 */
async function createFeishuRecord(appToken, tableId, fields) {
  const accessToken = await ensureAccessToken();

  try {
    const response = await httpRequest(
      `https://open.feishu.cn/open-apis/bitable/v1/apps/${appToken}/tables/${tableId}/records`,
      {
        method: 'POST',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      },
      { fields }
    );

    if (response.data && response.data.record) {
      console.log(`飞书API创建记录成功: ${JSON.stringify(fields)}`);
      return {
        record_id: response.data.record.record_id,
        fields: response.data.record.fields
      };
    } else {
      throw new Error('飞书API未返回创建的记录数据');
    }
  } catch (error) {
    console.error(`飞书API创建记录失败:`, error.message);
    throw error;
  }
}

/**
 * 获取商品的HS编码，如果不存在则创建新记录
 * @param {string} goodsName - 英文品名
 * @param {string} appToken - 多维表格应用 token
 * @param {string} tableId - 表格 ID
 * @param {string} goodsNameField - 商品名字段名，默认为"英文品名"
 * @param {string} hsCodeField - HS编码字段名，默认为"HS编码"
 * @returns {Promise<string>} 8位HS编码
 */
async function getOrCreateHSCode(goodsName, appToken, tableId, goodsNameField = '英文品名', hsCodeField = 'HS编码') {
  // 确保商品名称为大写
  const upperGoodsName = goodsName.toUpperCase().trim();

  console.log(`开始查询商品"${upperGoodsName}"的HS编码...`);

  // 查询是否存在该商品 - 使用批量查询和模糊匹配
  console.log(`使用字段名查询: 商品名称字段="${goodsNameField}"，HS编码字段="${hsCodeField}"`);
  const record = await queryFeishuRecord(appToken, tableId, goodsNameField, upperGoodsName, {
    useBulkQuery: true, // 使用批量查询，像PowerShell命令那样
    useFuzzyMatch: true // 使用模糊匹配，因为商品名称可能不完全匹配
  });

  if (record) {
    console.log(`找到商品"${upperGoodsName}"的记录，字段列表:`, Object.keys(record.fields));
    console.log(`完整字段值:`, record.fields);

    // 从记录中提取HS编码字段 - 优先使用传入的字段名，然后尝试多种可能的字段名
    const hsCode = record.fields[hsCodeField] ||
                   record.fields['HS编码'] ||
                   record.fields['hs_code'] ||
                   record.fields['HS'] ||
                   record.fields['hs'] ||
                   record.fields['编码'] ||
                   record.fields['HS Code'] ||
                   record.fields['HS CODE'] ||
                   '';

    console.log(`提取的HS编码: "${hsCode}"，类型: ${typeof hsCode}，长度: ${String(hsCode).length}`);

    // 检查是否是有效的8位HS编码
    const hsCodeStr = String(hsCode || '').trim();
    if (hsCodeStr && hsCodeStr.length === 8 && /^\d{8}$/.test(hsCodeStr)) {
      console.log(`商品"${upperGoodsName}"的HS编码有效: ${hsCodeStr}`);
      return hsCodeStr;
    } else if (hsCodeStr && hsCodeStr.length > 0) {
      console.warn(`商品"${upperGoodsName}"的HS编码格式不正确: "${hsCodeStr}"，长度: ${hsCodeStr.length}，使用默认值12345678`);
      return '12345678';
    } else {
      console.warn(`商品"${upperGoodsName}"的HS编码字段为空，使用默认值12345678`);
      return '12345678';
    }
  } else {
    // 创建新记录
    console.log(`商品"${upperGoodsName}"不存在于飞书多维表格中，添加新记录...`);
    const fields = {
      [goodsNameField]: upperGoodsName,
      [hsCodeField]: '12345678'
    };
    console.log(`创建新记录的字段:`, fields);
    try {
      const newRecord = await createFeishuRecord(appToken, tableId, fields);
      console.log(`新记录添加成功，记录ID: ${newRecord.record_id}`);
      return '12345678';
    } catch (error) {
      console.error(`创建新记录失败:`, error.message);
      // 即使创建失败，也返回默认HS编码
      return '12345678';
    }
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
    console.log(`处理商品: "${goodsName}"`);
    // 获取或创建HS编码
    const hsCode = await getOrCreateHSCode(goodsName, appToken, tableId, goodsNameField, hsCodeField);
    console.log(`商品"${goodsName}"的HS编码: ${hsCode}`);
    results.push(`${goodsName} ${hsCode}`);
  }

  // 重新组合为逗号分隔的字符串
  return results.join(', ');
}

/**
 * 获取表格的字段定义列表
 * @param {string} baseToken - 多维表格应用token
 * @param {string} tableId - 表格ID
 * @returns {Promise<Array<object>>} 字段定义数组，按表格中的顺序排列
 */
async function getTableFields(baseToken, tableId) {
  const accessToken = await ensureAccessToken();

  try {
    const response = await httpRequest(
      `https://open.feishu.cn/open-apis/bitable/v1/apps/${baseToken}/tables/${tableId}/fields`,
      {
        method: 'GET',
        headers: {
          'Authorization': `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        }
      }
    );

    if (response.data && response.data.items) {
      // 返回字段列表，保持API返回的顺序（通常是表格中的顺序）
      return response.data.items;
    } else {
      throw new Error('无法获取字段定义');
    }
  } catch (error) {
    console.error(`获取表格字段定义失败:`, error.message);
    throw error;
  }
}

/**
 * 按商品描述关键词查询HS编码（模拟PowerShell命令功能）
 * 支持按字段索引或字段名进行筛选和返回
 * @param {string} baseToken - 飞书多维表格应用token（如：OoNybRydGaN6Wwspy41cnQQCnGe）
 * @param {string} tableId - 表格ID（如：tbl2uWivrvboRe2a）
 * @param {string} searchText - 搜索关键词（如："ROPE"），不区分大小写
 * @param {object} options - 选项
 * @param {number} options.fieldIndex - 要筛选的字段索引（从0开始），如果同时指定fieldName则优先使用fieldName
 * @param {string} options.fieldName - 要筛选的字段名，如果未指定则使用fieldIndex
 * @param {number} options.returnIndex - 要返回的字段索引（从0开始），如果同时指定returnFieldName则优先使用returnFieldName
 * @param {string} options.returnFieldName - 要返回的字段名，如果未指定则使用returnIndex
 * @param {number} options.limit - 最大记录数，默认为500
 * @param {boolean} options.useRegex - 是否使用正则表达式匹配，默认为false（使用包含匹配）
 * @returns {Promise<Array<string>>} 匹配记录的指定字段值数组
 */
async function queryHSByDescription(baseToken, tableId, searchText, options = {}) {
  const {
    fieldIndex = 1,           // 默认筛选第二个字段（商品描述）
    fieldName = null,         // 筛选字段名（优先使用）
    returnIndex = 0,          // 默认返回第一个字段（HS编码）
    returnFieldName = null,   // 返回字段名（优先使用）
    limit = 500,              // 默认限制500条记录
    useRegex = false          // 是否使用正则表达式
  } = options;

  // 处理正则表达式前缀，如(?i)表示不区分大小写
  let actualSearchText = searchText;
  let actualUseRegex = useRegex;

  // 自动检测(?i)前缀
  if (searchText.startsWith('(?i)')) {
    actualSearchText = searchText.substring(4);
    actualUseRegex = true;
    console.log(`检测到正则表达式前缀(?i)，使用正则表达式模式，实际搜索文本: "${actualSearchText}"`);
  }

  console.log(`开始按描述查询HS编码: baseToken=${baseToken}, tableId=${tableId}, searchText="${actualSearchText}"`);
  console.log(`选项:`, options);

  const accessToken = await ensureAccessToken();
  let pageToken = '';
  let allRecords = [];
  let pageCount = 0;
  const searchTextUpper = actualSearchText.toUpperCase();

  // 获取表格字段定义以确定字段顺序
  let tableFields = [];
  try {
    tableFields = await getTableFields(baseToken, tableId);
    console.log(`获取到表格字段定义，共${tableFields.length}个字段:`, tableFields.map(f => f.field_name));
  } catch (error) {
    console.warn(`无法获取表格字段定义，将使用字段对象键的顺序:`, error.message);
  }

  try {
    // 分页获取所有记录，直到达到限制
    do {
      pageCount++;
      let url = `https://open.feishu.cn/open-apis/bitable/v1/apps/${baseToken}/tables/${tableId}/records?page_size=100`;
      if (pageToken) {
        url += `&page_token=${pageToken}`;
      }

      console.log(`批量查询第${pageCount}页...`);

      const response = await httpRequest(
        url,
        {
          method: 'GET',
          headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
          }
        }
      );

      if (response.data && response.data.items) {
        allRecords = allRecords.concat(response.data.items);
        console.log(`第${pageCount}页获取到${response.data.items.length}条记录，总计${allRecords.length}条`);

        // 检查是否有下一页
        pageToken = response.data.page_token || '';

        // 如果已达到限制，停止获取
        if (allRecords.length >= limit) {
          console.log(`达到记录限制${limit}条，停止获取更多记录`);
          allRecords = allRecords.slice(0, limit); // 截取到限制数量
          pageToken = ''; // 停止分页
        }
      } else {
        console.log(`第${pageCount}页无数据`);
        break;
      }

      // 防止无限循环，最多获取10页（1000条记录）
      if (pageCount >= 10) {
        console.log(`达到最大分页限制（10页），停止获取`);
        break;
      }

    } while (pageToken);

    console.log(`批量查询完成，共获取${allRecords.length}条记录`);

    // 确定要筛选和返回的字段
    let filterFieldKey = null;
    let returnFieldKey = null;

    // 如果有字段定义，使用字段名查找字段键
    if (tableFields.length > 0) {
      if (fieldName) {
        const field = tableFields.find(f => f.field_name === fieldName);
        if (field) {
          filterFieldKey = field.field_id;
          console.log(`使用字段名"${fieldName}"筛选，字段键: ${filterFieldKey}`);
        } else {
          console.warn(`未找到字段名"${fieldName}"，将使用索引${fieldIndex}`);
        }
      }

      if (returnFieldName) {
        const field = tableFields.find(f => f.field_name === returnFieldName);
        if (field) {
          returnFieldKey = field.field_id;
          console.log(`使用字段名"${returnFieldName}"返回，字段键: ${returnFieldKey}`);
        } else {
          console.warn(`未找到字段名"${returnFieldName}"，将使用索引${returnIndex}`);
        }
      }
    }

    // 如果没有字段定义或未找到字段名，使用索引
    // 需要根据字段定义顺序确定字段键
    if (!filterFieldKey && tableFields.length > fieldIndex) {
      filterFieldKey = tableFields[fieldIndex].field_id;
      console.log(`使用索引${fieldIndex}筛选，字段键: ${filterFieldKey}，字段名: ${tableFields[fieldIndex].field_name}`);
    }

    if (!returnFieldKey && tableFields.length > returnIndex) {
      returnFieldKey = tableFields[returnIndex].field_id;
      console.log(`使用索引${returnIndex}返回，字段键: ${returnFieldKey}，字段名: ${tableFields[returnIndex].field_name}`);
    }

    // 筛选匹配的记录
    const matchedValues = [];

    for (const record of allRecords) {
      const fields = record.fields || {};

      let filterValue = '';
      let returnValue = '';

      // 如果确定了字段键，直接使用
      if (filterFieldKey && fields[filterFieldKey] !== undefined) {
        filterValue = String(fields[filterFieldKey] || '');
      } else if (fieldName && fields[fieldName] !== undefined) {
        // 尝试直接使用字段名作为键（某些情况下字段名可能就是键）
        filterValue = String(fields[fieldName] || '');
      } else {
        // 使用索引回退方案：将字段对象转换为数组
        const fieldEntries = Object.entries(fields);

        // 如果有字段定义，按字段定义顺序排序
        if (tableFields.length > 0) {
          // 创建字段ID到索引的映射
          const fieldIdToIndex = {};
          tableFields.forEach((field, index) => {
            fieldIdToIndex[field.field_id] = index;
          });

          // 按字段定义顺序排序
          fieldEntries.sort((a, b) => {
            const indexA = fieldIdToIndex[a[0]] !== undefined ? fieldIdToIndex[a[0]] : Infinity;
            const indexB = fieldIdToIndex[b[0]] !== undefined ? fieldIdToIndex[b[0]] : Infinity;
            return indexA - indexB;
          });
        } else {
          // 按字段名排序作为回退
          fieldEntries.sort((a, b) => a[0].localeCompare(b[0]));
        }

        const fieldValues = fieldEntries.map(([key, value]) => String(value || ''));

        if (fieldValues.length > Math.max(fieldIndex, returnIndex)) {
          filterValue = fieldValues[fieldIndex];
          returnValue = fieldValues[returnIndex];
        }
      }

      // 如果还没有获取返回值，但确定了返回字段键
      if (!returnValue && returnFieldKey && fields[returnFieldKey] !== undefined) {
        returnValue = String(fields[returnFieldKey] || '');
      } else if (!returnValue && returnFieldName && fields[returnFieldName] !== undefined) {
        returnValue = String(fields[returnFieldName] || '');
      }

      // 进行匹配
      if (filterValue) {
        const filterValueUpper = filterValue.toUpperCase();
        let isMatch = false;

        if (useRegex) {
          try {
            const regex = new RegExp(searchText, 'i'); // 不区分大小写
            isMatch = regex.test(filterValue);
          } catch (error) {
            console.warn(`正则表达式"${searchText}"无效，使用包含匹配:`, error.message);
            isMatch = filterValueUpper.includes(searchTextUpper);
          }
        } else {
          isMatch = filterValueUpper.includes(searchTextUpper);
        }

        if (isMatch) {
          if (!returnValue) {
            // 如果还没有返回值，尝试使用索引回退
            const fieldEntries = Object.entries(fields);
            const fieldValues = fieldEntries.map(([key, value]) => String(value || ''));
            if (fieldValues.length > returnIndex) {
              returnValue = fieldValues[returnIndex];
            }
          }

          if (returnValue) {
            console.log(`找到匹配记录: "${filterValue}" 包含 "${searchText}"`);
            console.log(`返回值: "${returnValue}"`);
            matchedValues.push(returnValue);
          } else {
            console.warn(`找到匹配记录但无法获取返回值: "${filterValue}"`);
          }
        }
      }
    }

    console.log(`查询完成，共找到${matchedValues.length}条匹配记录`);
    return matchedValues;

  } catch (error) {
    console.error(`按描述查询HS编码失败:`, error.message);
    throw error;
  }
}

/**
 * 完全模拟PowerShell命令的HS编码查询逻辑（按数组索引匹配）
 * @param {string} baseToken - 飞书多维表格应用token
 * @param {string} tableId - 表格ID
 * @param {string} searchText - 搜索关键词（支持(?i)前缀）
 * @param {object} options - 选项
 * @param {number} options.fieldIndex - 要筛选的字段索引（从0开始），默认1
 * @param {number} options.returnIndex - 要返回的字段索引（从0开始），默认0
 * @param {number} options.limit - 最大记录数，默认500
 * @returns {Promise<Array<string>>} 匹配记录的指定字段值数组
 */
async function queryHSByDescriptionPSLike(baseToken, tableId, searchText, options = {}) {
  const {
    fieldIndex = 1,
    returnIndex = 0,
    limit = 500
  } = options;

  // 处理PowerShell风格的(?i)前缀（不区分大小写）
  const isCaseInsensitive = searchText.startsWith('(?i)');
  const actualSearchText = isCaseInsensitive ? searchText.substring(4) : searchText;
  const searchPattern = new RegExp(actualSearchText, isCaseInsensitive ? 'i' : '');

  console.log(`[queryHSByDescriptionPSLike] 使用 lark-cli 查询: baseToken=${baseToken}, tableId=${tableId}`);
  console.log(`[queryHSByDescriptionPSLike] 搜索: "${searchText}" -> /${actualSearchText}/${isCaseInsensitive ? 'i' : ''}`);
  console.log(`[queryHSByDescriptionPSLike] 字段索引: 筛选字段[${fieldIndex}], 返回字段[${returnIndex}], 限制: ${limit}`);

  try {
    // 执行 lark-cli 命令（使用与PowerShell命令相同的认证方式）
    const command = `lark-cli base +record-list --base-token ${baseToken} --table-id ${tableId} --limit ${limit}`;
    console.log(`[queryHSByDescriptionPSLike] 执行命令: ${command}`);

    const { stdout, stderr } = await execAsync(command);

    if (stderr) {
      console.warn(`[queryHSByDescriptionPSLike] lark-cli stderr: ${stderr}`);
    }

    // 解析 JSON 输出
    const result = JSON.parse(stdout);

    if (!result.ok) {
      throw new Error(`lark-cli 返回错误: ${JSON.stringify(result)}`);
    }

    const { data, fields, field_id_list } = result.data;

    console.log(`[queryHSByDescriptionPSLike] 获取到 ${data.length} 条记录`);
    console.log(`[queryHSByDescriptionPSLike] 字段列表: ${fields.join(', ')}`);
    console.log(`[queryHSByDescriptionPSLike] 字段ID列表: ${field_id_list.join(', ')}`);

    // 筛选匹配的记录（完全模仿PowerShell的$_[1] -match逻辑）
    const matchedValues = [];

    for (const record of data) {
      // record 是数组，例如: [39269090, "PEN ORGANIZER", "笔筒 / 笔架"]
      // fields 是字段名列表: ["HS编码", "英文品名", "中文备注"]

      // 检查索引是否有效
      if (record.length <= Math.max(fieldIndex, returnIndex)) {
        console.warn(`[queryHSByDescriptionPSLike] 记录字段数不足: ${record.length}, 需要至少 ${Math.max(fieldIndex, returnIndex) + 1}`);
        continue;
      }

      // 获取筛选字段的值（对应PowerShell的$_[1]）
      const filterValue = String(record[fieldIndex] || '').trim();

      // 匹配逻辑：和PowerShell的-match完全一致
      if (searchPattern.test(filterValue)) {
        // 获取返回字段的值（对应PowerShell的$_[0]）
        const returnValue = String(record[returnIndex] || '').trim();

        if (returnValue) {
          matchedValues.push(returnValue);
          console.log(`[queryHSByDescriptionPSLike] 匹配: "${filterValue}" -> "${returnValue}"`);
        }
      }
    }

    console.log(`[queryHSByDescriptionPSLike] 找到 ${matchedValues.length} 条匹配记录`);
    return matchedValues;

  } catch (error) {
    console.error(`[queryHSByDescriptionPSLike] PowerShell风格查询失败:`, error.message);
    throw error;
  }
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

  // 创建大写的商品名称数组用于查询（表格中商品名称为大写）
  const upperGoodsNames = goodsNames.map(name => name.toUpperCase());

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
    instructions += `${index + 1}. ${name} (查询时将转换为: ${upperGoodsNames[index]})\n`;
  });

  instructions += `\n注意：表格中的商品名称为大写格式，查询时会自动转换为大写。\n`;
  instructions += `\n## 操作步骤:\n\n`;

  if (parsedUrl.isWikiLink) {
    // Wiki链接需要先解析
    instructions += `### 1. 解析Wiki链接获取baseToken\n`;
    instructions += `先执行: \`lark-cli wiki spaces get_node ${parsedUrl.wikiToken} --output json\`\n`;
    instructions += `假设返回的baseToken为 \`bas_xxx\`，tableId为 \`${parsedUrl.tableId}\`\n\n`;

    instructions += `### 2. 查询现有记录\n`;
    goodsNames.forEach((name, index) => {
      const upperName = upperGoodsNames[index];
      instructions += `商品${index + 1} "${name}" (查询名称: "${upperName}"):\n`;
      instructions += `\`lark-cli base list-records bas_xxx ${parsedUrl.tableId} --filter '英文品名="${upperName}"' --output json\`\n\n`;
    });

    instructions += `### 3. 创建不存在的记录\n`;
    instructions += `对于查询不到的商品，使用以下命令创建（HS编码为12345678）：\n`;
    instructions += `\`lark-cli base create-record bas_xxx ${parsedUrl.tableId} --fields '{"英文品名":"商品名称","HS编码":"12345678"}' --output json\`\n`;
    instructions += `注意：将"商品名称"替换为大写的商品英文名称。\n\n`;
  } else {
    // 直接Base链接
    const { appToken, tableId } = parsedUrl;

    instructions += `### 1. 查询现有记录\n`;
    goodsNames.forEach((name, index) => {
      const upperName = upperGoodsNames[index];
      instructions += `商品${index + 1} "${name}" (查询名称: "${upperName}"):\n`;
      instructions += `\`lark-cli base list-records ${appToken} ${tableId} --filter '英文品名="${upperName}"' --output json\`\n\n`;
    });

    instructions += `### 2. 创建不存在的记录\n`;
    instructions += `对于查询不到的商品，使用以下命令创建（HS编码为12345678）：\n`;
    instructions += `\`lark-cli base create-record ${appToken} ${tableId} --fields '{"英文品名":"商品名称","HS编码":"12345678"}' --output json\`\n`;
    instructions += `注意：将"商品名称"替换为大写的商品英文名称。\n\n`;
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
  getTableFields,
  queryHSByDescription,
  queryHSByDescriptionPSLike,
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

2. 按描述查询HS编码（模拟PowerShell命令）:
   const { queryHSByDescription } = require('./hs-encoder');
   // 基本用法（按字段索引）
   const results1 = await queryHSByDescription(
     "OoNybRydGaN6Wwspy41cnQQCnGe", // baseToken
     "tbl2uWivrvboRe2a", // tableId
     "ROPE", // 搜索关键词
     {
       fieldIndex: 1,     // 第二个字段是商品描述
       returnIndex: 0,    // 第一个字段是HS编码
       limit: 500         // 最多500条记录
     }
   );
   // 使用字段名（更可靠）
   const results2 = await queryHSByDescription(
     "OoNybRydGaN6Wwspy41cnQQCnGe",
     "tbl2uWivrvboRe2a",
     "ROPE",
     {
       fieldName: "商品描述",     // 筛选字段名
       returnFieldName: "HS编码", // 返回字段名
       limit: 500
     }
   );
   // 使用正则表达式（模拟PowerShell的-match操作符）
   const results3 = await queryHSByDescription(
     "OoNybRydGaN6Wwspy41cnQQCnGe",
     "tbl2uWivrvboRe2a",
     "(?i)ROPE", // (?i)表示不区分大小写，自动启用正则表达式
     {
       fieldIndex: 1,
       returnIndex: 0,
       limit: 500
     }
   );

   // 完全模拟PowerShell命令的查询（按数组索引匹配，不依赖字段定义）
   const { queryHSByDescriptionPSLike } = require('./hs-encoder');
   const results4 = await queryHSByDescriptionPSLike(
     "OoNybRydGaN6Wwspy41cnQQCnGe",
     "tbl2uWivrvboRe2a",
     "(?i)ROPE", // 搜索关键词（不区分大小写）
     {
       fieldIndex: 1,  // 匹配第2个字段（和PowerShell的$_[1]一致）
       returnIndex: 0, // 返回第1个字段（和PowerShell的$_[0]一致）
       limit: 500      // 和PowerShell的--limit 500一致
     }
   );

3. 使用lark-cli技能手动操作:
   const { getWorkflowInstructions } = require('./hs-encoder');
   const instructions = getWorkflowInstructions("Apple, Banana, Orange", "https://yourdomain.feishu.cn/base/appToken?table=tableId");
   console.log(instructions);

注意事项:
- 需要先安装并配置 lark-cli 或使用 Claude Code 中的 lark-base 技能
- 飞书多维表格中应有"英文品名"和"HS编码"字段（或相应字段名）
- HS编码字段应为文本类型，存储8位字符
- 查询时会自动将商品名称转换为大写进行匹配
- 如果商品不存在，会自动创建新记录，HS编码为12345678
- 需要设置有效的FEISHU_APP_ID和FEISHU_APP_SECRET环境变量
- queryHSByDescription函数需要表格的字段定义权限
`);
}