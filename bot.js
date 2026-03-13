/**
 * 钉钉文档收集机器人 - 支持HTTP回调签名验证
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const zlib = require('zlib');

// ============ 配置 ============

const botConfig = {
  appKey: process.env.BOT_APP_KEY || '',
  appSecret: process.env.BOT_APP_SECRET || '',
  agentId: process.env.BOT_AGENT_ID || '',
  // 钉钉事件订阅的密钥
  token: process.env.DINGTALK_TOKEN || '53oMEiOSL3WNmNniyGwVz3ogjXyTL',
  aesKey: process.env.DINGTALK_AES_KEY || 'odOv37WkKoIMLtpOKH8cRbxrrzpyMhQx9A0t6H8k2cX',
  name: process.env.BOT_NAME || '文档收集助手',
  port: parseInt(process.env.BOT_PORT) || 3000
};

const storageConfig = {
  baseDir: process.env.STORAGE_BASE_DIR || './received_documents',
  allowedExtensions: ['.pdf', '.doc', '.docx', '.docm', '.xls', '.xlsx', '.xlsm', '.ppt', '.pptx', '.txt', '.csv', '.md', '.json', '.xml', '.zip', '.rar']
};

if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 请配置 .env');
  process.exit(1);
}

console.log('✅ 配置加载');

// ============ 初始化 ============

const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;
const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const baseDir = path.resolve(storageConfig.baseDir);
if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

// ============ 钉钉签名验证和解密 ============

function decodeAESKey(encodedKey) {
  const keyBuffer = Buffer.from(encodedKey + '=', 'base64');
  return keyBuffer.slice(0, 32);
}

function signature(token, timestamp, nonce, encrypt) {
  const sortArr = [token, timestamp, nonce, encrypt].sort();
  const signStr = sortArr.join('');
  return crypto.createHash('sha1').update(signStr).digest('hex');
}

function decrypt(encrypt, aesKey) {
  try {
    const key = decodeAESKey(aesKey);
    const cipher = crypto.createDecipheriv('aes-256-cbc', key, key);
    let decrypted = cipher.update(encrypt, 'base64', 'utf8');
    decrypted += cipher.final('utf8');
    
    // 去除 PKCS7 填充
    const pad = decrypted.charCodeAt(decrypted.length - 1);
    return decrypted.slice(0, decrypted.length - pad);
  } catch (e) {
    console.error('解密失败:', e.message);
    return null;
  }
}

// ============ 工具函数 ============

function getCurrentDate() {
  return new Date(Date.now() + TIMEZONE_OFFSET);
}

function formatDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

let cachedToken = null;
let tokenExpireTime = 0;

async function getAccessToken() {
  const now = Date.now();
  if (cachedToken && now < tokenExpireTime) return cachedToken;
  
  const res = await axios.post('https://api.dingtalk.com/v1.0/oauth2/accessToken', {
    appKey: botConfig.appKey,
    appSecret: botConfig.appSecret
  });
  
  cachedToken = res.data.accessToken;
  tokenExpireTime = now + (res.data.expireIn - 300) * 1000;
  console.log('🔑 获取Token');
  return cachedToken;
}

function loadLog() {
  try { return fs.existsSync('./download_log.json') ? JSON.parse(fs.readFileSync('./download_log.json', 'utf-8')) : []; } catch { return []; }
}

function saveLog(logs) {
  try { fs.writeFileSync('./download_log.json', JSON.stringify(logs.slice(0, 1000), null, 2)); } catch {}
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

async function downloadWithTmpCode(tmpCode, fileName, token) {
  console.log('📥 使用 tmpCode 下载...');
  
  const response = await axios.post(
    'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
    { tmpCode: tmpCode },
    {
      headers: {
        'x-acs-dingtalk-access-token': token,
        'Content-Type': 'application/json'
      },
      responseType: 'stream'
    }
  );
  
  const contentType = response.headers['content-type'] || '';
  console.log('   Content-Type:', contentType);
  
  if (contentType.includes('application/json')) {
    let data = '';
    for await (const chunk of response.data) { data += chunk.toString(); }
    throw new Error(data.substring(0, 100));
  }
  
  return response.data;
}

async function downloadFile(content, fileName) {
  const token = await getAccessToken();
  const tmpCode = content.downloadCode;
  
  console.log(`📥 下载: ${fileName}`);
  
  const stream = await downloadWithTmpCode(tmpCode, fileName, token);
  
  const dateDir = path.join(baseDir, formatDate(getCurrentDate()));
  if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
  
  const ext = path.extname(fileName);
  const safeName = sanitizeFileName(path.basename(fileName, ext));
  const finalName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName}${ext}`;
  const filePath = path.join(dateDir, finalName);
  
  return new Promise((resolve, reject) => {
    const writer = fs.createWriteStream(filePath);
    stream.pipe(writer);
    writer.on('finish', () => {
      const stats = fs.statSync(filePath);
      resolve({ path: filePath, name: finalName, originalName: fileName, size: stats.size, dateDir: formatDate(getCurrentDate()) });
    });
    writer.on('error', reject);
  });
}

// ============ 发送回复 ============

async function sendReply(message, text) {
  try {
    if (message.sessionWebhook) {
      await axios.post(message.sessionWebhook, { msgtype: 'text', text: { content: text } });
      console.log('   ✅ 已回复');
      return;
    }
    
    if (message.conversationId) {
      const token = await getAccessToken();
      const url = message.conversationType === '2' 
        ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
        : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
      
      await axios.post(url, {
        robotId: botConfig.agentId,
        openConversationId: message.conversationId,
        msgtype: 'text',
        text: { content: text }
      }, { headers: { 'x-acs-dingtalk-access-token': token } });
      
      console.log('   ✅ 已回复');
    }
  } catch (e) {
    console.log('   ⚠️ 回复失败:', e.message);
  }
}

// ============ 处理消息 ============

async function handleMessage(message) {
  const isGroup = message.conversationType === '2';
  console.log(`\n📩 ${message.msgtype} | ${isGroup ? '群聊' : '私聊'}`);
  
  if (message.msgtype === 'file') {
    let content = {};
    let fileName = '未知文件';
    
    if (message.content) {
      if (typeof message.content === 'string') {
        try { content = JSON.parse(message.content); } catch {}
      } else {
        content = message.content;
      }
    }
    
    fileName = content.fileName || message.fileName || '未知文件';
    console.log('   文件:', fileName);
    
    if (!isAllowedExtension(fileName)) {
      await sendReply(message, '⚠️ 不支持此格式');
      return;
    }
    
    if (!content.downloadCode) {
      await sendReply(message, '❌ 无法获取下载凭证');
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`   ✅ 成功: ${result.name} (${sizeMB}MB)`);
      
      addLog({ type: 'file', originalName: result.originalName, fileName: result.name, size: sizeMB, dateDir: result.dateDir });
      
      await sendReply(message, `✅ 已保存！\n\n文件名: ${result.originalName}\n大小: ${sizeMB} MB\n日期: ${result.dateDir}`);
    } catch (e) {
      console.log('   ❌ 失败:', e.message);
      await sendReply(message, `❌ 下载失败: ${e.message.substring(0, 50)}`);
    }
  }
  else if (message.msgtype === 'text') {
    const text = (typeof message.text === 'string' ? message.text : message.text?.content || '').trim();
    console.log('   文本:', text);
    
    if (text === '/帮助') {
      await sendReply(message, '📖 发送文档自动保存\n支持: PDF,Word,Excel,PPT,TXT,MD\n\n命令: /状态 /列表');
    }
    else if (text === '/状态') {
      const logs = loadLog().filter(l => l.type === 'file');
      const totalSize = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendReply(message, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${totalSize} MB`);
    }
    else if (text === '/列表') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let msg = '📋 最近:\n\n';
      logs.forEach((l, i) => msg += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.dateDir}\n\n`);
      await sendReply(message, msg || '暂无');
    }
  }
}

// ============ Webhook 回调 ============

app.get('/webhook', (req, res) => {
  console.log('🔔 回调验证');
  
  const { signature, timestamp, nonce, echostr } = req.query;
  
  // 验证签名
  const mySignature = signature(botConfig.token, timestamp, nonce, echostr);
  
  if (mySignature !== signature) {
    console.log('❌ 签名验证失败');
    return res.status(403).send('签名验证失败');
  }
  
  // 解密
  const decryptStr = decrypt(echostr, botConfig.aesKey);
  console.log('✅ 解密成功:', decryptStr);
  
  res.send(decryptStr);
});

app.post('/webhook', async (req, res) => {
  res.send('success');
  
  try {
    const { signature, timestamp, nonce, msg_signature } = req.query;
    const encrypt = req.body.encrypt;
    
    console.log('📨 收到加密消息');
    
    // 验证签名
    const mySignature = signature(botConfig.token, timestamp, nonce, encrypt);
    
    if (mySignature !== signature) {
      console.log('❌ 签名验证失败');
      return;
    }
    
    // 解密
    const decryptStr = decrypt(encrypt, botConfig.aesKey);
    console.log('📝 解密后:', decryptStr.substring(0, 200));
    
    // 解析JSON
    let message;
    try {
      message = JSON.parse(decryptStr);
    } catch {
      console.log('❌ 解析失败');
      return;
    }
    
    await handleMessage(message);
    
  } catch (e) {
    console.error('❌ 处理错误:', e.message);
  }
});

// ============ 健康检查 ============

app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ============ 启动 ============

app.listen(botConfig.port, '0.0.0.0', () => {
  console.log(`\n🤖 ${botConfig.name} 启动 | 端口: ${botConfig.port}\n`);
  console.log(`📝 Token: ${botConfig.token}`);
  console.log(`🔑 AES Key: ${botConfig.aesKey.substring(0, 10)}...\n`);
});

module.exports = app;
