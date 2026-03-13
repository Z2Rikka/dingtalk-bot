/**
 * 钉钉文档收集机器人 - Stream模式
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { EventEmitter } = require('events');

// ============ 配置 ============

const config = {
  appKey: process.env.DINGTALK_APP_KEY || process.env.BOT_APP_KEY || '',
  appSecret: process.env.DINGTALK_APP_SECRET || process.env.BOT_APP_SECRET || '',
  agentId: process.env.DINGTALK_AGENT_ID || process.env.BOT_AGENT_ID || '',
  port: parseInt(process.env.BOT_PORT) || 3565,
  storageDir: process.env.STORAGE_DIR || process.env.STORAGE_BASE_DIR || './received_documents',
  allowedExt: ['.pdf', '.doc', '.docx', '.md', '.txt', '.xls', '.xlsx', '.ppt', '.pptx', '.csv', '.zip', '.rar']
};

if (!config.appKey || !config.appSecret) {
  console.error('❌ 请配置环境变量');
  process.exit(1);
}

console.log('✅ 配置加载');

// ============ 初始化 ============

const app = express();
app.use(express.json());

const baseDir = path.resolve(config.storageDir);
if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

// ============ 工具函数 ============

function getToday() {
  const d = new Date(Date.now() + 8 * 60 * 60 * 1000);
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, '0')}${String(d.getDate()).padStart(2, '0')}`;
}

let tokenCache = { token: null, expire: 0 };

async function getToken() {
  const now = Date.now();
  if (tokenCache.token && now < tokenCache.expire) return tokenCache.token;
  
  const res = await axios.post('https://api.dingtalk.com/v1.0/oauth2/accessToken', {
    appKey: config.appKey,
    appSecret: config.appSecret
  });
  
  tokenCache.token = res.data.accessToken;
  tokenCache.expire = now + (res.data.expireIn - 300) * 1000;
  return tokenCache.token;
}

function loadLog() {
  try { return fs.existsSync('./download_log.json') ? JSON.parse(fs.readFileSync('./download_log.json')) : []; } catch { return []; }
}

function saveLog(logs) {
  try { fs.writeFileSync('./download_log.json', JSON.stringify(logs.slice(0, 1000), null, 2)); } catch {}
}

function isAllowed(name) {
  return config.allowedExt.includes(path.extname(name).toLowerCase());
}

function safeName(name) {
  return name.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

// ============ 下载文件 ============

async function downloadFile(content, fileName) {
  const token = await getToken();
  const { downloadCode } = content;
  
  console.log(`📥 下载: ${fileName}`);
  
  const res = await axios.post(
    'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
    { tmpCode: downloadCode },
    { headers: { 'x-acs-dingtalk-access-token': token, 'Content-Type': 'application/json' }, responseType: 'stream' }
  );
  
  if (res.headers['content-type']?.includes('application/json')) {
    throw new Error('返回JSON错误');
  }
  
  const dateDir = path.join(baseDir, getToday());
  if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
  
  const ext = path.extname(fileName);
  const finalName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName(path.basename(fileName, ext))}${ext}`;
  const filePath = path.join(dateDir, finalName);
  
  return new Promise((resolve, reject) => {
    res.data.pipe(fs.createWriteStream(filePath)).on('finish', () => {
      const stats = fs.statSync(filePath);
      resolve({ name: finalName, original: fileName, size: stats.size, date: getToday() });
    }).on('error', reject);
  });
}

// ============ 发送消息 ============

async function sendText(conversationId, text, conversationType = '1') {
  try {
    const token = await getToken();
    const url = conversationType === '2' 
      ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
      : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
    
    await axios.post(url, {
      robotId: config.agentId,
      openConversationId: conversationId,
      msgtype: 'text',
      text: { content: text }
    }, { headers: { 'x-acs-dingtalk-access-token': token } });
  } catch (e) {
    console.log('发送失败:', e.message);
  }
}

// ============ 处理消息 ============

async function onMessage(msg) {
  console.log(`\n📩 ${msg.msgtype} | ${msg.conversationType === '2' ? '群聊' : '私聊'}`);
  
  const conversationId = msg.conversationId;
  const conversationType = msg.conversationType;
  
  if (msg.msgtype === 'file') {
    let content = {};
    let fileName = '未知文件';
    
    if (msg.content) {
      try { content = typeof msg.content === 'string' ? JSON.parse(msg.content) : msg.content; } catch {}
    }
    
    fileName = content.fileName || msg.fileName || '未知文件';
    console.log(`   文件: ${fileName}`);
    
    if (!isAllowed(fileName)) {
      await sendText(conversationId, '⚠️ 不支持此格式', conversationType);
      return;
    }
    
    if (!content.downloadCode) {
      await sendText(conversationId, '❌ 无法获取下载凭证', conversationType);
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`   ✅ 成功: ${result.name} (${sizeMB}MB)`);
      
      const logs = loadLog();
      logs.unshift({ type: 'file', originalName: result.original, fileName: result.name, size: sizeMB, date: result.date, timestamp: new Date().toISOString() });
      saveLog(logs);
      
      await sendText(conversationId, `✅ 已保存！\n\n文件名: ${result.original}\n大小: ${sizeMB} MB\n日期: ${result.date}`, conversationType);
    } catch (e) {
      console.log(`   ❌ 失败: ${e.message}`);
      await sendText(conversationId, `❌ 下载失败: ${e.message.substring(0, 50)}`, conversationType);
    }
  }
  else if (msg.msgtype === 'text') {
    const text = typeof msg.text === 'string' ? msg.text : (msg.text?.content || '').trim();
    console.log(`   文本: ${text}`);
    
    if (text === '/帮助') await sendText(conversationId, '📖 发送文档自动保存\n支持: PDF,Word,Excel,MD\n\n命令: /状态 /列表', conversationType);
    else if (text === '/状态') {
      const logs = loadLog().filter(l => l.type === 'file');
      const total = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendText(conversationId, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${total} MB`, conversationType);
    }
    else if (text === '/列表') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let m = '📋 最近:\n\n';
      logs.forEach((l, i) => m += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.date}\n\n`);
      await sendText(conversationId, m || '暂无', conversationType);
    }
  }
}

// ============ Stream 模式 ============

const WebSocket = require('ws');

async function startStream() {
  const token = await getToken();
  
  // 获取 websocket 地址
  const res = await axios.post('https://api.dingtalk.com/v1.0/robot/oToMessages/getWebsocketEndpoint', 
    { robotCode: config.agentId },
    { headers: { 'x-acs-dingtalk-access-token': token } }
  );
  
  const endpoint = res.data.endpoint;
  console.log(`🔗 连接Stream: ${endpoint}`);
  
  const ws = new WebSocket(endpoint);
  
  ws.on('open', () => {
    console.log('✅ Stream 连接成功\n');
  });
  
  ws.on('message', async (data) => {
    try {
      const msg = JSON.parse(data.toString());
      console.log('收到消息:', msg.topic);
      
      if (msg.topic === 'cloudimap.message.receive') {
        const content = msg.data;
        if (content && content.msgtype) {
          await onMessage(content);
        }
      }
    } catch (e) {
      console.log('解析消息失败:', e.message);
    }
  });
  
  ws.on('error', (e) => {
    console.log('❌ Stream错误:', e.message);
  });
  
  ws.on('close', () => {
    console.log('🔄 Stream断开，5秒后重连...');
    setTimeout(startStream, 5000);
  });
}

// ============ HTTP 服务器 ============

app.get('/health', (req, res) => res.json({ status: 'ok', mode: 'stream' }));

app.listen(config.port, '0.0.0.0', () => {
  console.log(`🤖 文档收集助手 | 端口: ${config.port}\n`);
});

// 启动 Stream
startStream().catch(err => {
  console.error('❌ Stream启动失败:', err.message);
});

module.exports = app;
