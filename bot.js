/**
 * 钉钉文档收集机器人 v3
 * 修复下载和回复问题
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');

// ============ 配置 ============

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
  ]
};

const logConfig = {
  enabled: process.env.LOG_ENABLED !== 'false',
  file: process.env.LOG_FILE || './download_log.json'
};

// 验证配置
if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 请先配置 .env 文件中的 BOT_APP_KEY、BOT_APP_SECRET、BOT_AGENT_ID');
  process.exit(1);
}

console.log('✅ 配置加载成功');
console.log(`   AgentId: ${botConfig.agentId}`);

// ============ 常量 ============

const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;
const app = express();
app.use(express.json());

// ============ 初始化 ============

const baseDir = path.resolve(storageConfig.baseDir);
if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

const logFile = path.resolve(logConfig.file);

// ============ 工具函数 ============

function getCurrentDate() {
  return new Date(Date.now() + TIMEZONE_OFFSET);
}

function formatDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

async function getAccessToken() {
  const url = 'https://api.dingtalk.com/v1.0/oauth2/accessToken';
  const res = await axios.post(url, {
    appKey: botConfig.appKey,
    appSecret: botConfig.appSecret
  });
  return res.data.accessToken;
}

function loadLog() {
  if (!logConfig.enabled) return [];
  try { return fs.existsSync(logFile) ? JSON.parse(fs.readFileSync(logFile, 'utf-8')) : []; } catch { return []; }
}

function saveLog(logs) {
  if (!logConfig.enabled) return;
  try { fs.writeFileSync(logFile, JSON.stringify(logs, null, 2)); } catch {}
}

function addLog(entry) {
  const logs = loadLog();
  logs.unshift({ ...entry, timestamp: new Date().toISOString() });
  saveLog(logs);
}

function isAllowedExtension(fileName) {
  return storageConfig.allowedExtensions.includes(path.extname(fileName).toLowerCase());
}

function sanitizeFileName(fileName) {
  return fileName.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

// ============ 下载文件 ============

async function downloadFile(downloadCode, originalName, accessToken) {
  // 方法1: 使用官方文档的API
  const url = 'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode';
  
  console.log(`📥 开始下载: ${originalName}`);
  
  try {
    const response = await axios.post(url, {
      tmpCode: downloadCode
    }, {
      headers: { 
        'x-acs-dingtalk-access-token': accessToken,
        'Content-Type': 'application/json'
      },
      responseType: 'stream'
    });
    
    // 检查是否是错误响应
    if (response.headers['content-type'] && response.headers['content-type'].includes('application/json')) {
      // 读取错误响应
      const chunks = [];
      for await (const chunk of response.data) {
        chunks.push(chunk);
      }
      const errorBody = Buffer.concat(chunks).toString();
      console.log('❌ 下载失败，返回JSON:', errorBody);
      throw new Error(errorBody);
    }
    
    // 验证格式
    if (!isAllowedExtension(originalName)) {
      throw new Error(`不支持的格式: ${path.extname(originalName)}`);
    }
    
    // 保存文件
    const dateDir = path.join(baseDir, formatDate(getCurrentDate()));
    if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
    
    const ext = path.extname(originalName);
    const safeName = sanitizeFileName(path.basename(originalName, ext));
    const fileName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName}${ext}`;
    const filePath = path.join(dateDir, fileName);
    
    return new Promise((resolve, reject) => {
      const writer = fs.createWriteStream(filePath);
      response.data.pipe(writer);
      writer.on('finish', () => {
        const stats = fs.statSync(filePath);
        resolve({
          path: filePath,
          name: fileName,
          originalName: originalName,
          size: stats.size,
          dateDir: formatDate(getCurrentDate())
        });
      });
      writer.on('error', reject);
    });
  } catch (error) {
    console.log('❌ 下载错误:', error.message);
    throw error;
  }
}

// ============ 发送消息 ============

async function sendMessageBySession(sessionWebhook, content) {
  if (!sessionWebhook) {
    console.log('⚠️ 没有sessionWebhook，无法回复');
    return;
  }
  
  try {
    await axios.post(sessionWebhook, {
      msgtype: 'text',
      text: { content }
    });
    console.log('✅ 消息已发送');
  } catch (error) {
    console.log('❌ 发送消息失败:', error.response?.data || error.message);
  }
}

async function sendReply(message, content) {
  // 优先使用 sessionWebhook（私聊用这个更可靠）
  if (message.sessionWebhook) {
    return await sendMessageBySession(message.sessionWebhook, content);
  }
  
  // 否则使用 openConversationId
  if (message.conversationId) {
    try {
      const accessToken = await getAccessToken();
      const url = message.conversationType === '2' 
        ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
        : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
      
      await axios.post(url, {
        robotId: botConfig.agentId,
        openConversationId: message.conversationId,
        msgtype: 'text',
        text: { content }
      }, {
        headers: { 'x-acs-dingtalk-access-token': accessToken }
      });
      console.log('✅ 消息已发送');
    } catch (error) {
      console.log('❌ 发送消息失败:', error.response?.data || error.message);
    }
  }
}

// ============ 处理消息 ============

async function handleMessage(message) {
  console.log(`\n📩 消息: ${message.msgtype} | 类型: ${message.conversationType === '2' ? '群聊' : '私聊'}`);
  console.log(`   sessionWebhook: ${message.sessionWebhook ? '有' : '无'}`);
  console.log(`   conversationId: ${message.conversationId ? '有' : '无'}`);
  
  const accessToken = await getAccessToken();
  
  // 处理文件消息
  if (message.msgtype === 'file') {
    const content = message.content || {};
    const fileName = content.fileName || message.fileName || '未知文件';
    const downloadCode = content.downloadCode;
    
    console.log(`📄 文件: ${fileName}`);
    console.log(`🔑 downloadCode: ${downloadCode ? '有' : '无'}`);
    
    // 检查格式
    if (!isAllowedExtension(fileName)) {
      const formats = storageConfig.allowedExtensions.map(e => e.replace('.', '')).join(', ');
      await sendReply(message, `⚠️ 不支持此格式，仅支持: ${formats}`);
      addLog({ type: 'rejected', originalName: fileName, reason: 'unsupported_format' });
      return;
    }
    
    // 检查downloadCode
    if (!downloadCode) {
      await sendReply(message, '❌ 无法下载：缺少下载凭证');
      return;
    }
    
    try {
      const result = await downloadFile(downloadCode, fileName, accessToken);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`✅ 下载成功: ${result.name} (${sizeMB} MB)`);
      
      addLog({
        type: 'file',
        originalName: result.originalName,
        fileName: result.name,
        size: sizeMB,
        dateDir: result.dateDir,
        path: result.path
      });
      
      await sendReply(message, `✅ 文档已保存！\n\n文件名: ${result.originalName}\n大小: ${sizeMB} MB\n日期: ${result.dateDir}`);
      
    } catch (error) {
      console.log('❌ 下载失败:', error.message);
      await sendReply(message, `❌ 下载失败: ${error.message}`);
    }
  }
  // 处理文本命令
  else if (message.msgtype === 'text') {
    const text = typeof message.text === 'string' ? message.text : (message.text?.content || '');
    console.log(`📝 文本: ${text}`);
    
    if (text === '/帮助' || text === '/help') {
      await sendReply(message, '📖 文档收集助手\n\n发送文档自动保存\n支持: PDF,Word,Excel,PPT,TXT,RAR,ZIP\n\n命令: /状态 /列表 /目录 /帮助');
    }
    else if (text === '/状态' || text === '/stats') {
      const logs = loadLog().filter(l => l.type === 'file');
      const totalSize = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendReply(message, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${totalSize} MB`);
    }
    else if (text === '/列表' || text === '/list') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let msg = '📋 最近文档:\n\n';
      logs.forEach((l, i) => msg += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.dateDir}\n\n`);
      await sendReply(message, msg || '暂无');
    }
    else if (text === '/目录' || text === '/dir') {
      try {
        const dirs = fs.readdirSync(baseDir).filter(d => /^\d{8}$/.test(d)).sort().reverse();
        await sendReply(message, `📁 ${baseDir}\n\n${dirs.slice(0,10).map(d => '• ' + d).join('\n')}`);
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
  console.log('\n========== 收到请求 ==========');
  res.send('success');
  
  try {
    await handleMessage(req.body);
  } catch (e) {
    console.error('❌ 处理错误:', e.message);
  }
  console.log('========== 处理完成 ==========\n');
});

// ============ API ============

app.get('/api/files', (req, res) => res.json({ total: loadLog().length, files: loadLog() }));
app.get('/api/stats', (req, res) => {
  const logs = loadLog().filter(l => l.type === 'file');
  res.json({ total: logs.length, totalSize: logs.reduce((s,l) => s + parseFloat(l.size||0), 0).toFixed(2) + ' MB' });
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
  console.log('测试: 发送PDF文件给机器人\n');
});

module.exports = app;
