// 飞书多维表格API服务 - 用于获取HS编码
const axios = require('axios');

// 飞书API配置
const FEISHU_BASE_URL = 'https://open.feishu.cn/open-apis';

// 缓存token和HS编码数据
let accessToken = null;
let tokenExpireTime = 0;
const hsCodeCache = new Map();
let cacheInitTime = 0;
const CACHE_DURATION = 3600000; // 1小时缓存

// 获取tenant_access_token
async function getAccessToken() {
  if (accessToken && Date.now() < tokenExpireTime) {
    return accessToken;
  }

  try {
    const response = await axios.post(`${FEISHU_BASE_URL}/auth/v3/tenant_access_token/internal`, {
      app_id: process.env.FEISHU_APP_ID,
      app_secret: process.env.FEISHU_APP_SECRET
    });

    if (response.data.code !== 0) {
      throw new Error(`获取飞书access_token失败: ${response.data.msg}`);
    }

    accessToken = response.data.tenant_access_token;
    tokenExpireTime = Date.now() + (response.data.expire - 300) * 1000; // 提前5分钟过期
    console.log('飞书access_token已更新');
    return accessToken;
  } catch (error) {
    console.error('获取飞书access_token失败:', error.message);
    throw error;
  }
}

// 初始化HS编码缓存
async function initHSCodeCache() {
  // 检查缓存是否有效
  if (Date.now() - cacheInitTime < CACHE_DURATION && hsCodeCache.size > 0) {
    return;
  }

  try {
    const token = await getAccessToken();
    const bitableId = process.env.FEISHU_BITABLE_ID;
    const tableId = process.env.FEISHU_TABLE_ID;
    const viewId = process.env.FEISHU_VIEW_ID;

    if (!bitableId || !tableId) {
      console.warn('飞书表格ID未配置，将使用默认HS编码');
      return;
    }

    // 构建请求参数
    const params = {
      page_size: 100
    };
    if (viewId) {
      params.view_id = viewId;
    }

    const response = await axios.get(
      `${FEISHU_BASE_URL}/bitable/v1/apps/${bitableId}/tables/${tableId}/records`,
      {
        headers: { Authorization: `Bearer ${token}` },
        params: params
      }
    );

    if (response.data.code !== 0) {
      throw new Error(`获取飞书表格数据失败: ${response.data.msg}`);
    }

    hsCodeCache.clear();
    response.data.data.items.forEach(item => {
      const fields = item.fields;
      // 从飞书表格字段中获取数据
      const englishName = fields['英文品名'];
      const hsCode = fields['HS编码'];
      if (englishName && hsCode) {
        hsCodeCache.set(String(englishName).trim(), String(hsCode).trim());
      }
    });

    cacheInitTime = Date.now();
    console.log(`HS编码缓存已更新，共 ${hsCodeCache.size} 条记录`);
  } catch (error) {
    console.error('初始化HS编码缓存失败:', error.message);
    // 失败时保持原有缓存（如果有），不做清空
  }
}

// 根据英文品名获取HS编码
async function getHSCode(englishName) {
  await initHSCodeCache();

  const name = String(englishName).trim();

  // 精确匹配
  if (hsCodeCache.has(name)) {
    return hsCodeCache.get(name);
  }

  // 模糊匹配（包含关系）
  for (const [key, value] of hsCodeCache.entries()) {
    if (name.includes(key) || key.includes(name)) {
      console.log(`HS编码模糊匹配: "${name}" -> "${key}" (${value})`);
      return value;
    }
  }

  // 未找到，返回默认值
  console.warn(`未找到HS编码: ${name}，使用默认值88886666`);
  return '88886666';
}

// 批量获取HS编码
async function getHSCodes(englishNames) {
  const results = await Promise.all(
    englishNames.map(name => getHSCode(name.trim()))
  );
  return results;
}

// 清除缓存（用于测试或强制刷新）
function clearCache() {
  hsCodeCache.clear();
  cacheInitTime = 0;
  accessToken = null;
  tokenExpireTime = 0;
  console.log('飞书API缓存已清除');
}

module.exports = {
  getHSCode,
  getHSCodes,
  clearCache
};
