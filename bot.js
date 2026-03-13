/**
 * 钉钉文档收集机器人
 * 支持从 .env 文件读取配置
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// ============ 加载配置 ============

const botConfig = {
  appKey: process.env.BOT_APP_KEY || '',
  appSecret: process.env.BOT_APP_SECRET || '',
  agentId: process.env.BOT_AGENT_ID || '',
  name: process.env.BOT_NAME || '文档收集助手',
  port: parseInt(process.env.BOT_PORT) || 3000
};

const storageConfig = {
  baseDir: process.env.STORAGE_BASE_DIR || './received_documents',
  allowedExtensions: [
    '.pdf', '.doc', '.docx', '.docm', '.dotx', '.dotm',
    '.xls', '.xlsx', '.xlsm', '.xlsb', '.xltx', '.xltm',
    '.ppt', '.pptx', '.pptm', '.potx', '.potm',
    '.txt', '.csv', '.md', '.json', '.xml',
    '.rar', '.zip', '.7z'
  ],
  maxFileNameLength: parseInt(process.env.STORAGE_MAX_FILENAME_LENGTH) || 200
};

const messageConfig = {
  autoReply: process.env.MESSAGE_AUTO_REPLY !== 'false',
  allowedGroupIds: (process.env.MESSAGE_ALLOWED_GROUP_IDS || '').split(',').filter(Boolean)
};

const logConfig = {
  enabled: process.env.LOG_ENABLED !== 'false',
  file: process.env.LOG_FILE || './download_log.json',
  maxEntries: parseInt(process.env.LOG_MAX_ENTRIES) || 1000
};

// ============ 验证配置 ============

if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 配置错误：BOT_APP_KEY、BOT_APP_SECRET、BOT_AGENT_ID 必须填写！');
  console.error('请编辑 .env 文件配置这些值。');
  process.exit(1);
}

console.log('✅ 配置加载成功');
console.log(`   AppKey: ${botConfig.appKey.substring(0, 8)}...`);
console.log(`   AgentId: ${botConfig.agentId}`);
console.log(`   端口: ${botConfig.port}`);

// ============ 常量 ============

const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;
const app = express();
app.use(express.json());

// ============ 初始化目录 ============

const baseDir = path.resolve(storageConfig.baseDir);
if (!fs.existsSync(baseDir)) {
  fs.mkdirSync(baseDir, { recursive: true });
}

const logFile = path.resolve(logConfig.file);

// ============ 工具函数 ============

function ensureDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath, { recursive: true });
  }
}

function getCurrentDate() {
  return new Date(Date.now() + TIMEZONE_OFFSET);
}

function formatDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

async function getAccessToken() {
  const url = 'https://api.dingtalk.com/v1.0/oauth2/accessToken';
  const response = await axios.post(url, {
    appKey: botConfig.appKey,
    appSecret: botConfig.appSecret
  });
  return response.data.accessToken;
}

function loadLog() {
  if (!logConfig.enabled) return [];
  try {
    if (fs.existsSync(logFile)) {
      return JSON.parse(fs.readFileSync(logFile, 'utf-8'));
    }
  } catch (e) {}
  return [];
}

function saveLog(logs) {
  if (!logConfig.enabled) return;
  try {
    fs.writeFileSync(logFile, JSON.stringify(logs, null, 2));
  } catch (e) {}
}

function addLog(entry) {
  const logs = loadLog();
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  if (logs.length > logConfig.maxEntries) logs.length = logConfig.maxEntries;
  saveLog(logs);
}

function isAllowedExtension(fileName) {
  const ext = path.extname(fileName).toLowerCase();
  return storageConfig.allowedExtensions.includes(ext);
}

function getExtension(fileName) {
  return path.extname(fileName).toLowerCase();
}

function sanitizeFileName(fileName) {
  let safeName = fileName.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
  const ext = getExtension(safeName);
  const baseName = path.basename(safeName, ext);
  if (baseName.length > storageConfig.maxFileNameLength - ext.length) {
    safeName = baseName.substring(0, storageConfig.maxFileNameLength - ext.length) + ext;
  }
  return safeName;
}

function generateFileName(originalName) {
  const safeName = sanitizeFileName(originalName);
  return `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName}`;
}

function getDateDir() {
  const dateStr = formatDate(getCurrentDate());
  const dateDir = path.join(baseDir, dateStr);
  ensureDir(dateDir);
  return dateDir;
}

function getStoragePath(fileName) {
  if (!isAllowedExtension(fileName)) {
    throw new Error(`不支持的格式: ${getExtension(fileName)}`);
  }
  return path.join(getDateDir(), generateFileName(fileName));
}

async function downloadFileWithCode(downloadCode, fileName, accessToken) {
  const url = 'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode';
  
  const response = await axios.post(url, {
    tmpCode: downloadCode
  }, {
    headers: { 
      'x-acs-dingtalk-access-token': accessToken,
      'Content-Type': 'application/json'
    },
    responseType: 'stream'
  });
  
  if (!isAllowedExtension(fileName)) {
    throw new Error(`不支持的格式: ${getExtension(fileName)}`);
  }
  
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

async function sendMessage(targetId, targetType, content) {
  if (!messageConfig.autoReply) return;
  try {
    const accessToken = await getAccessToken();
    const url = targetType === 'group' 
      ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
      : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
    
    await axios.post(url, {
      robotId: botConfig.agentId,
      openConversationId: targetId,
      msgtype: 'text',
      text: { content }
    }, { headers: { 'x-acs-dingtalk-access-token': accessToken } });
  } catch (e) {
    console.error('发送消息失败:', e.message);
  }
}

function getAllowedFormatsList() {
  const formats = new Set();
  storageConfig.allowedExtensions.forEach(ext => {
    formats.add(ext.replace('.', '').toUpperCase());
  });
  return Array.from(formats).join(', ');
}

// ============ 处理文件消息 ============

async function handleFileMessage(message, accessToken) {
  const content = message.content || {};
  const fileName = content.fileName || message.fileName || 'unknown_file';
  const downloadCode = content.downloadCode;
  
  console.log(`📥 收到文档: ${fileName}`);
  console.log(`   downloadCode: ${downloadCode ? '有' : '无'}`);
  
  const targetType = message.conversationType === '2' ? 'group' : 'private';
  
  // 验证格式
  if (!isAllowedExtension(fileName)) {
    const formats = getAllowedFormatsList();
    await sendMessage(message.conversationId, targetType, 
      `⚠️ 不支持此格式，仅支持: ${formats}`);
    
    addLog({ type: 'rejected', reason: 'unsupported_format', originalName: fileName, conversationId: message.conversationId });
    return null;
  }
  
  if (!downloadCode) {
    console.log('❌ 没有 downloadCode');
    await sendMessage(message.conversationId, targetType, '❌ 文件下载失败：无法获取下载凭证');
    return null;
  }
  
  // 下载
  const result = await downloadFileWithCode(downloadCode, fileName, accessToken);
  const fileSizeMB = (result.size / 1024 / 1024).toFixed(2);
  
  console.log(`✅ 保存成功: ${result.name} (${fileSizeMB} MB)`);
  
  addLog({
    type: 'file',
    fileName: result.name,
    originalName: result.originalName,
    fileSize: fileSizeMB,
    dateDir: result.dateDir,
    conversationId: message.conversationId,
    senderNick: message.senderNick
  });
  
  await sendMessage(message.conversationId, targetType, 
    `✅ 文档已保存\n\n文件名: ${result.originalName}\n大小: ${fileSizeMB} MB\n日期: ${result.dateDir}`);
  
  return result;
}

// ============ 消息处理 ============

async function handleMessage(message) {
  console.log(`\n📩 消息: ${message.msgtype} | 会话类型: ${message.conversationType} | 发送者: ${message.senderNick || message.senderId}`);
  
  const accessToken = await getAccessToken();
  const targetType = message.conversationType === '2' ? 'group' : 'private';
  
  if (message.msgtype === 'file') {
    return await handleFileMessage(message, accessToken);
  }
  else if (message.msgtype === 'text') {
    const text = message.text?.content || message.text || '';
    
    if (text === '/帮助' || text === '/help') {
      await sendMessage(message.conversationId, targetType, 
        '📖 文档收集助手\n\n发送文档自动保存\n支持: PDF,Word,Excel,PPT,TXT,RAR,ZIP\n\n命令: /状态 /列表 /目录 /帮助');
    }
    else if (text === '/状态' || text === '/stats') {
      const logs = loadLog().filter(l => l.type === 'file');
      const totalSize = logs.reduce((s, l) => s + parseFloat(l.fileSize || 0), 0);
      await sendMessage(message.conversationId, targetType, 
        `📊 统计\n\n文档: ${logs.length} 个\n大小: ${totalSize.toFixed(2)} MB`);
    }
    else if (text === '/列表' || text === '/list') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let msg = '📋 最近文档:\n\n';
      logs.forEach((l, i) => msg += `${i+1}. ${l.originalName}\n${l.fileSize}MB | ${l.dateDir}\n\n`);
      await sendMessage(message.conversationId, targetType, msg || '暂无');
    }
    else if (text === '/目录' || text === '/dir') {
      try {
        const dirs = fs.readdirSync(baseDir).filter(d => /^\d{8}$/.test(d)).sort().reverse();
        await sendMessage(message.conversationId, targetType, 
          `📁 ${baseDir}\n\n${dirs.slice(0,10).map(d => '• ' + d).join('\n')}`);
      } catch(e) {}
    }
  }
}

// ============ Webhook ============

app.get('/webhook', (req, res) => {
  console.log('🔔 回调验证');
  res.send(req.query.echostr);
});

app.post('/webhook', async (req, res) => {
  console.log('\n========== 收到消息 ==========');
  res.send('success');
  try {
    await handleMessage(req.body);
  } catch (e) {
    console.error('❌ 处理失败:', e.message);
  }
  console.log('========== 处理完成 ==========\n');
});

// ============ API ============

app.get('/api/files', (req, res) => res.json({ total: loadLog().length, files: loadLog() }));
app.get('/api/stats', (req, res) => {
  const logs = loadLog().filter(l => l.type === 'file');
  const totalSize = logs.reduce((s, l) => s + parseFloat(l.fileSize || 0), 0);
  res.json({ totalFiles: logs.length, totalSize: totalSize.toFixed(2), storageDir: baseDir });
});
app.get('/health', (req, res) => res.json({ status: 'ok', name: botConfig.name, port: botConfig.port }));

// ============ 启动 ============

app.listen(botConfig.port, '0.0.0.0', () => {
  console.log('========================================');
  console.log(`🤖 ${botConfig.name} 已启动`);
  console.log(`📡 端口: ${botConfig.port}`);
  console.log(`📁 目录: ${baseDir}`);
  console.log('========================================');
  console.log(`\n回调地址: http://你的IP:${botConfig.port}/webhook`);
  console.log('支持: PDF,Word,Excel,PPT,TXT,RAR,ZIP');
  console.log('命令: /帮助 /状态 /列表 /目录\n');
});

module.exports = app;
