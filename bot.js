/**
 * 钉钉文档收集机器人
 * 功能：监听群消息/私聊，下载文档到服务器
 * 按日期（YYYYMMDD）存储，UTC+8时区
 * 只接受常见文档格式
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const config = require('./config');

const app = express();
app.use(express.json());

// ============ 常量 ============

// 时区偏移（UTC+8）
const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;

// ============ 初始化目录 ============

// 基础存储目录
const baseDir = path.resolve(config.storage.baseDir);
ensureDir(baseDir);

// 日志
const logFile = path.resolve(config.log.file);

// ============ 工具函数 ============

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

/**
 * 获取当前日期（UTC+8）
 */
function getCurrentDate() {
  const now = new Date(Date.now() + TIMEZONE_OFFSET);
  return now;
}

/**
 * 格式化日期为 YYYYMMDD
 */
function formatDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}${month}${day}`;
}

/**
 * 获取Access Token
 */
async function getAccessToken() {
  const url = 'https://api.dingtalk.com/v1.0/oauth2/accessToken';
  const response = await axios.post(url, {
    appKey: config.bot.appKey,
    appSecret: config.bot.appSecret
  });
  return response.data.accessToken;
}

/**
 * 加载日志
 */
function loadLog() {
  if (!config.log.enabled) return [];
  try {
    if (fs.existsSync(logFile)) {
      return JSON.parse(fs.readFileSync(logFile, 'utf-8'));
    }
  } catch (e) {
    console.error('加载日志失败:', e);
  }
  return [];
}

/**
 * 保存日志
 */
function saveLog(logs) {
  if (!config.log.enabled) return;
  try {
    fs.writeFileSync(logFile, JSON.stringify(logs, null, 2));
  } catch (e) {
    console.error('保存日志失败:', e);
  }
}

/**
 * 添加日志
 */
function addLog(entry) {
  const logs = loadLog();
  logs.unshift({
    ...entry,
    timestamp: new Date().toISOString()
  });
  if (logs.length > config.log.maxEntries) {
    logs.length = config.log.maxEntries;
  }
  saveLog(logs);
}

/**
 * 检查文件扩展名是否允许
 */
function isAllowedExtension(fileName) {
  const ext = path.extname(fileName).toLowerCase();
  return config.storage.allowedExtensions.includes(ext);
}

/**
 * 获取扩展名（带点号）
 */
function getExtension(fileName) {
  return path.extname(fileName).toLowerCase();
}

/**
 * 生成安全文件名
 */
function sanitizeFileName(fileName) {
  // 限制文件名长度
  const maxLen = config.storage.maxFileNameLength;
  let safeName = fileName;
  
  // 移除或替换非法字符
  safeName = safeName.replace(/[<>:"/\\|?*]/g, '_');
  safeName = safeName.replace(/\s+/g, '_');
  
  // 限制长度
  const ext = getExtension(safeName);
  const baseName = path.basename(safeName, ext);
  if (baseName.length > maxLen - ext.length) {
    safeName = baseName.substring(0, maxLen - ext.length) + ext;
  }
  
  return safeName;
}

/**
 * 生成唯一文件名
 */
function generateFileName(originalName) {
  const safeName = sanitizeFileName(originalName);
  const timestamp = Date.now();
  const random = crypto.randomBytes(4).toHexString();
  return `${timestamp}_${random}_${safeName}`;
}

/**
 * 获取日期目录路径
 */
function getDateDir() {
  const date = getCurrentDate();
  const dateStr = formatDate(date);
  const dateDir = path.join(baseDir, dateStr);
  ensureDir(dateDir);
  return dateDir;
}

/**
 * 验证并返回目标存储路径
 */
function getStoragePath(fileName) {
  // 检查扩展名
  if (!isAllowedExtension(fileName)) {
    throw new Error(`不支持的格式: ${getExtension(fileName)}`);
  }
  
  // 获取日期目录
  const dateDir = getDateDir();
  
  // 生成唯一文件名
  const uniqueName = generateFileName(fileName);
  
  return path.join(dateDir, uniqueName);
}

/**
 * 下载钉钉文件
 */
async function downloadFile(fileId, fileName, accessToken) {
  const url = `https://api.dingtalk.com/v1.0/robot/file/download?robotId=${config.bot.agentId}&fileId=${fileId}`;
  
  const response = await axios.get(url, {
    headers: { 
      'x-acs-dingtalk-access-token': accessToken 
    },
    responseType: 'stream',
    timeout: 300000 // 5分钟超时
  });
  
  // 验证文件格式
  if (!isAllowedExtension(fileName)) {
    throw new Error(`不支持的格式: ${getExtension(fileName)}`);
  }
  
  // 获取存储路径
  const filePath = getStoragePath(fileName);
  
  return new Promise((resolve, reject) => {
    const writer = fs.createWriteStream(filePath);
    response.data.pipe(writer);
    writer.on('finish', () => {
      const stats = fs.statSync(filePath);
      resolve({
        path: filePath,
        name: path.basename(filePath),
        originalName: fileName,
        size: stats.size,
        dateDir: path.basename(path.dirname(filePath))
      });
    });
    writer.on('error', reject);
  });
}

/**
 * 发送消息到群/私聊
 */
async function sendMessage(targetId, targetType, content) {
  if (!config.message.autoReply) return;
  
  try {
    const accessToken = await getAccessToken();
    let url;
    
    if (targetType === 'group') {
      url = `https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend`;
    } else {
      url = `https://api.dingtalk.com/v1.0/robot/oToMessages/send`;
    }
    
    await axios.post(url, {
      robotId: config.bot.agentId,
      openConversationId: targetId,
      msgtype: 'text',
      text: { content }
    }, {
      headers: { 'x-acs-dingtalk-access-token': accessToken }
    });
  } catch (error) {
    console.error('发送消息失败:', error.message);
  }
}

/**
 * 获取允许的格式列表
 */
function getAllowedFormatsList() {
  const formats = new Set();
  config.storage.allowedExtensions.forEach(ext => {
    if (['.doc', '.xls', '.ppt'].includes(ext)) {
      formats.add(ext.toUpperCase());
    } else {
      formats.add(ext.toUpperCase().replace('.', ''));
    }
  });
  return Array.from(formats).join(', ');
}

/**
 * 处理文件消息
 */
async function handleFileMessage(message, accessToken) {
  const fileId = message.file.fileId;
  const fileName = message.file.fileName || 'unknown_file';
  
  console.log(`📥 收到文档: ${fileName}`);
  
  // 验证格式
  if (!isAllowedExtension(fileName)) {
    const allowedFormats = getAllowedFormatsList();
    const reply = config.message.replyTemplates.unsupported
      .replace('{formats}', allowedFormats)
      .replace('{actual}', getExtension(fileName));
    
    console.log(`❌ 不支持的格式: ${getExtension(fileName)}`);
    
    await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', reply);
    
    // 记录日志
    addLog({
      type: 'rejected',
      reason: 'unsupported_format',
      fileId: fileId,
      originalName: fileName,
      extension: getExtension(fileName),
      conversationId: message.conversationId,
      senderId: message.senderId,
      senderNick: message.senderNick,
      chatType: message.chatType
    });
    
    return null;
  }
  
  // 下载文件
  const result = await downloadFile(fileId, fileName, accessToken);
  const fileSizeMB = (result.size / 1024 / 1024).toFixed(2);
  
  console.log(`✅ 文档保存成功:`);
  console.log(`   原始文件名: ${result.originalName}`);
  console.log(`   保存文件名: ${result.name}`);
  console.log(`   文件大小: ${fileSizeMB} MB`);
  console.log(`   存储路径: ${result.path}`);
  console.log(`   日期目录: ${result.dateDir}`);
  
  // 记录日志
  addLog({
    type: 'file',
    fileId: fileId,
    fileName: result.name,
    originalName: result.originalName,
    extension: getExtension(result.originalName),
    fileSize: fileSizeMB,
    filePath: result.path,
    dateDir: result.dateDir,
    conversationId: message.conversationId,
    senderId: message.senderId,
    senderNick: message.senderNick,
    chatType: message.chatType
  });
  
  // 发送确认消息
  const reply = config.message.replyTemplates.success
    .replace('{filename}', result.originalName)
    .replace('{filesize}', fileSizeMB + ' MB')
    .replace('{date}', result.dateDir)
    .replace('{filepath}', result.path);
  
  await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', reply);
  
  return result;
}

/**
 * 消息处理入口
 */
async function handleMessage(message) {
  console.log(`\n📩 收到消息:`);
  console.log(`   类型: ${message.msgtype}`);
  console.log(`   会话: ${message.chatType}`);
  console.log(`   发送者: ${message.senderNick || message.senderId}`);
  
  const accessToken = await getAccessToken();
  
  // 只处理文件消息
  if (message.msgtype === 'file') {
    return await handleFileMessage(message, accessToken);
  }
  else if (message.msgtype === 'text') {
    // 处理文本命令
    const text = message.text.content.trim();
    console.log(`   文本内容: ${text}`);
    
    // 帮助命令
    if (text === '/帮助' || text === '/help') {
      const helpMsg = `📖 文档收集助手\n\n` +
        `发送文档自动保存到服务器\n\n` +
        `支持格式：PDF, Word, Excel, PPT, TXT, RAR, ZIP\n\n` +
        `命令：\n` +
        `• /状态 - 收集统计\n` +
        `• /列表 - 最近文件\n` +
        `• /目录 - 存储目录\n` +
        `• /帮助 - 显示帮助`;
      
      await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', helpMsg);
    }
    // 状态命令
    else if (text === '/状态' || text === '/stats') {
      const logs = loadLog();
      const totalSize = logs.reduce((sum, log) => sum + parseFloat(log.fileSize || 0), 0);
      const fileCount = logs.filter(l => l.type === 'file').length;
      
      const statsMsg = `📊 收集统计\n\n` +
        `📄 文档总数: ${fileCount} 个\n` +
        `💾 总大小: ${totalSize.toFixed(2)} MB\n` +
        `📁 存储目录: ${baseDir}`;
      
      await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', statsMsg);
    }
    // 列表命令
    else if (text === '/列表' || text === '/list') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let listMsg = `📋 最近文档:\n\n`;
      
      logs.forEach((log, i) => {
        listMsg += `${i+1}. ${log.originalName}\n`;
        listMsg += `   ${log.fileSize} MB | ${log.dateDir}\n\n`;
      });
      
      if (logs.length === 0) {
        listMsg = '暂无记录';
      }
      
      await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', listMsg);
    }
    // 目录命令
    else if (text === '/目录' || text === '/dir') {
      try {
        const dirs = fs.readdirSync(baseDir).filter(d => /^\d{8}$/.test(d)).sort().reverse();
        const dirMsg = `📁 存储目录结构\n\n` +
          `基础目录: ${baseDir}\n\n` +
          `日期目录:\n` +
          dirs.slice(0, 10).map(d => `• ${d}`).join('\n');
        
        await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', dirMsg);
      } catch (e) {
        await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', `目录错误: ${e.message}`);
      }
    }
  }
  else {
    // 忽略其他消息类型（图片、语音等）
    console.log(`   忽略消息类型: ${message.msgtype}（仅处理文档）`);
  }
}

// ============ Webhook 回调 ============

/**
 * 回调验证
 */
app.get('/webhook', (req, res) => {
  const { echostr } = req.query;
  console.log('🔔 收到回调验证请求');
  res.send(echostr);
});

/**
 * 消息接收
 */
app.post('/webhook', async (req, res) => {
  const message = req.body;
  
  console.log('\n========== 收到钉钉消息 ==========');
  
  // 返回成功响应
  res.send('success');
  
  // 异步处理消息
  try {
    await handleMessage(message);
  } catch (error) {
    console.error('❌ 处理消息失败:', error);
    
    if (message.conversationId) {
      const reply = config.message.replyTemplates.error.replace('{error}', error.message);
      await sendMessage(message.conversationId, message.chatType === 'group' ? 'group' : 'private', reply);
    }
  }
  
  console.log('========== 消息处理完成 ==========\n');
});

// ============ API 接口 ============

/**
 * 获取文件列表
 */
app.get('/api/files', (req, res) => {
  const logs = loadLog().filter(l => l.type === 'file');
  res.json({
    total: logs.length,
    files: logs
  });
});

/**
 * 获取统计信息
 */
app.get('/api/stats', (req, res) => {
  const logs = loadLog();
  const files = logs.filter(l => l.type === 'file');
  const totalSize = files.reduce((sum, log) => sum + parseFloat(log.fileSize || 0), 0);
  
  // 按日期统计
  const byDate = {};
  files.forEach(log => {
    const date = log.dateDir || 'unknown';
    byDate[date] = (byDate[date] || 0) + 1;
  });
  
  res.json({
    totalFiles: files.length,
    totalSize: totalSize.toFixed(2),
    byDate: byDate,
    storageDir: baseDir
  });
});

/**
 * 获取目录列表
 */
app.get('/api/dirs', (req, res) => {
  try {
    const dirs = fs.readdirSync(baseDir)
      .filter(d => /^\d{8}$/.test(d))
      .sort()
      .reverse();
    
    const result = {};
    dirs.forEach(dir => {
      const dirPath = path.join(baseDir, dir);
      const files = fs.readdirSync(dirPath);
      result[dir] = {
        path: dirPath,
        count: files.length,
        totalSize: files.reduce((sum, f) => {
          try {
            return sum + fs.statSync(path.join(dirPath, f)).size;
          } catch { return sum; }
        }, 0)
      };
      result[dir].totalSize = (result[dir].totalSize / 1024 / 1024).toFixed(2) + ' MB';
    });
    
    res.json(result);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

/**
 * 健康检查
 */
app.get('/health', (req, res) => {
  res.json({
    status: 'ok',
    name: config.bot.name,
    uptime: process.uptime(),
    storageDir: baseDir,
    allowedExtensions: config.storage.allowedExtensions
  });
});

// ============ 启动 ============

const PORT = config.bot.port;

app.listen(PORT, '0.0.0.0', () => {
  console.log('========================================');
  console.log(`🤖 ${config.bot.name} 已启动`);
  console.log(`📡 监听端口: ${PORT}`);
  console.log(`📁 存储目录: ${baseDir}`);
  console.log('========================================');
  console.log('');
  console.log(`回调地址: http://你的服务器IP:${PORT}/webhook`);
  console.log('');
  console.log('支持格式:');
  console.log('  PDF, Word, Excel, PPT, TXT, CSV, MD, JSON, XML, RAR, ZIP, 7Z');
  console.log('');
  console.log('API接口:');
  console.log('  GET /api/files   - 文件列表');
  console.log('  GET /api/stats  - 统计信息');
  console.log('  GET /api/dirs   - 目录列表');
  console.log('  GET /health    - 健康检查');
  console.log('');
  console.log('命令:');
  console.log('  /帮助   - 显示帮助');
  console.log('  /状态   - 收集统计');
  console.log('  /列表   - 最近文件');
  console.log('  /目录   - 存储目录');
  console.log('');
});

module.exports = app;
