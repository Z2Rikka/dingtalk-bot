/**
 * 钉钉机器人配置
 * 文档收集助手 - 监听群消息/私聊，下载文档到服务器
 * 按日期（YYYYMMDD）存储，UTC+8时区
 * 只接受常见文档格式
 * 
 * 支持环境变量配置，详见 .env.example
 */

function parseEnvBool(value, defaultValue) {
  if (value === undefined || value === '') return defaultValue;
  return value.toLowerCase() === 'true';
}

function parseEnvArray(value, defaultValue) {
  if (!value) return defaultValue;
  return value.split(',').map(item => item.trim()).filter(item => item);
}

function parseEnvInt(value, defaultValue) {
  const parsed = parseInt(value, 10);
  return isNaN(parsed) ? defaultValue : parsed;
}

const env = process.env;

module.exports = {
  // ============ 机器人配置 =============
  bot: {
    name: env.BOT_NAME || '文档收集助手',
    appKey: env.BOT_APP_KEY || '',
    appSecret: env.BOT_APP_SECRET || '',
    agentId: env.BOT_AGENT_ID || '',
    port: parseEnvInt(env.BOT_PORT, 3000)
  },
  
  // ============ 存储配置 ============
  storage: {
    baseDir: env.STORAGE_BASE_DIR || './received_documents',
    allowedExtensions: [
      '.pdf',
      '.doc', '.docx', '.docm', '.dotx', '.dotm',
      '.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm',
      '.ppt', '.pptx', '.pptm', '.potx', '.potm',
      '.txt', '.csv', '.md', '.json', '.xml',
      '.rar', '.zip', '.7z'
    ],
    maxFileNameLength: parseEnvInt(env.STORAGE_MAX_FILENAME_LENGTH, 200)
  },
  
  // ============ 消息配置 ============
  message: {
    autoReply: parseEnvBool(env.MESSAGE_AUTO_REPLY, true),
    listenGroups: parseEnvBool(env.MESSAGE_LISTEN_GROUPS, true),
    listenPrivate: parseEnvBool(env.MESSAGE_LISTEN_PRIVATE, true),
    allowedGroupIds: parseEnvArray(env.MESSAGE_ALLOWED_GROUP_IDS, []),
    replyTemplates: {
      received: '✅ 收到文档，正在验证格式...',
      success: '📄 文档已保存！\n\n文件名：{filename}\n大小：{filesize}\n日期：{date}\n路径：{filepath}',
      error: '❌ 保存失败：{error}',
      unsupported: '⚠️ 不支持的格式，仅支持：{formats}\n发送的文件格式为：{actual}',
      sizeLimit: '❌ 文件过大，超过 {maxsize}MB 限制'
    }
  },
  
  // ============ 日志配置 ============
  log: {
    enabled: parseEnvBool(env.LOG_ENABLED, true),
    file: env.LOG_FILE || './download_log.json',
    maxEntries: parseEnvInt(env.LOG_MAX_ENTRIES, 1000)
  }
};
