/**
 * 钉钉文档收集机器人 - 修复版
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
  token: process.env.DINGTALK_TOKEN || '',
  aesKey: process.env.DINGTALK_AES_KEY || '',
  name: process.env.BOT_NAME || '文档收集助手',
  port: parseInt(process.env.BOT_PORT) || 3000
};

const storageConfig = {
  baseDir: process.env.STORAGE_BASE_DIR || './received_documents',
  allowedExtensions: ['.pdf', '.doc', '.docx', '.md', '.txt', '.xls', '.xlsx', '.ppt', '.pptx', '.csv', '.zip', '.rar']
};

if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 请配置 .env');
  process.exit(1);
}

console.log('✅ 配置加载');

// ============ 初始化 ============

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

const baseDir = path.resolve(storageConfig.baseDir);
if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

// ============ 钉钉加解密 ============

function getSignature(token, timestamp, nonce, encrypt) {
  const str = [token, timestamp, nonce, encrypt].sort().join('');
  return crypto.createHash('sha1').update(str).digest('hex');
}

function decodeKey(encodedKey) {
  const str = encodedKey + '=';
  return Buffer.from(str, 'base64').slice(0, 32);
}

function decrypt(text, aesKey) {
  try {
    const key = decodeKey(aesKey);
    const decipher = crypto.createDecipheriv('aes-256-cbc', key, key);
    let decrypted = decipher.update(text, 'base64', 'utf8');
    decrypted += decipher.final('utf8');
    
    // 去除PKCS7填充
    const pad = decrypted.charCodeAt(decrypted.length - 1);
    if (pad > 0 && pad <= 16) {
      decrypted = decrypted.slice(0, -pad);
    }
    return decrypted;
  } catch (e) {
    console.error('解密错误:', e.message);
    return null;
  }
}

// ============ 工具函数 ============

function getCurrentDate() {
  const now = new Date(Date.now() + 8 * 60 * 60 * 1000);
  return `${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}`;
}

let cachedToken = null, tokenExpire = 0;

async function getAccessToken() {
  const now = Date.now();
  if (cachedToken && now < tokenExpire) return cachedToken;
  
  const res = await axios.post('https://api.dingtalk.com/v1.0/oauth2/accessToken', {
    appKey: botConfig.appKey,
    appSecret: botConfig.appSecret
  });
  
  cachedToken = res.data.accessToken;
  tokenExpire = now + (res.data.expireIn - 300) * 1000;
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

function isAllowed(fileName) {
  return storageConfig.allowedExtensions.includes(path.extname(fileName).toLowerCase());
}

function safeName(fileName) {
  return fileName.replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_');
}

// ============ 下载文件 ============

async function downloadFile(content, fileName) {
  const token = await getAccessToken();
  const tmpCode = content.downloadCode;
  
  console.log(`📥 下载: ${fileName}`);
  
  try {
    const res = await axios.post(
      'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
      { tmpCode },
      { headers: { 'x-acs-dingtalk-access-token': token, 'Content-Type': 'application/json' }, responseType: 'stream' }
    );
    
    if (res.headers['content-type']?.includes('application/json')) {
      throw new Error('返回JSON而非文件');
    }
    
    const dateDir = path.join(baseDir, getCurrentDate());
    if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
    
    const ext = path.extname(fileName);
    const finalName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName(path.basename(fileName, ext))}${ext}`;
    const filePath = path.join(dateDir, finalName);
    
    return new Promise((resolve, reject) => {
      res.data.pipe(fs.createWriteStream(filePath)).on('finish', () => {
        const stats = fs.statSync(filePath);
        resolve({ path: filePath, name: finalName, originalName: fileName, size: stats.size, date: getCurrentDate() });
      }).on('error', reject);
    });
  } catch (e) {
    throw new Error('下载失败: ' + e.message);
  }
}

// ============ 发送回复 ============

async function sendReply(msg, text) {
  try {
    if (msg.sessionWebhook) {
      await axios.post(msg.sessionWebhook, { msgtype: 'text', text: { content: text } });
      return;
    }
    
    if (msg.conversationId) {
      const token = await getAccessToken();
      const url = msg.conversationType === '2' 
        ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
        : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
      
      await axios.post(url, {
        robotId: botConfig.agentId,
        openConversationId: msg.conversationId,
        msgtype: 'text',
        text: { content: text }
      }, { headers: { 'x-acs-dingtalk-access-token': token } });
    }
  } catch (e) {
    console.log('回复失败:', e.message);
  }
}

// ============ 处理消息 ============

async function handleMessage(msg) {
  console.log(`\n📩 ${msg.msgtype} | ${msg.conversationType === '2' ? '群聊' : '私聊'}`);
  
  if (msg.msgtype === 'file') {
    let content = {};
    let fileName = '未知文件';
    
    if (msg.content) {
      if (typeof msg.content === 'string') {
        try { content = JSON.parse(msg.content); } catch {}
      } else {
        content = msg.content;
      }
    }
    
    fileName = content.fileName || msg.fileName || '未知文件';
    console.log(`   文件: ${fileName}`);
    
    if (!isAllowed(fileName)) {
      await sendReply(msg, '⚠️ 不支持此格式');
      return;
    }
    
    if (!content.downloadCode) {
      await sendReply(msg, '❌ 无下载凭证');
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`   ✅ 成功: ${result.name} (${sizeMB}MB)`);
      addLog({ type: 'file', originalName: result.originalName, fileName: result.name, size: sizeMB, date: result.date });
      
      await sendReply(msg, `✅ 已保存！\n\n文件名: ${result.originalName}\n大小: ${sizeMB} MB\n日期: ${result.date}`);
    } catch (e) {
      console.log(`   ❌ 失败: ${e.message}`);
      await sendReply(msg, `❌ 下载失败: ${e.message.substring(0, 50)}`);
    }
  }
  else if (msg.msgtype === 'text') {
    const text = typeof msg.text === 'string' ? msg.text : (msg.text?.content || '').trim();
    
    if (text === '/帮助') await sendReply(msg, '📖 发送文档自动保存\n支持: PDF,Word,Excel,MD\n\n命令: /状态 /列表');
    else if (text === '/状态') {
      const logs = loadLog().filter(l => l.type === 'file');
      const total = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendReply(msg, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${total} MB`);
    }
    else if (text === '/列表') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let m = '📋 最近:\n\n';
      logs.forEach((l, i) => m += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.date}\n\n`);
      await sendReply(msg, m || '暂无');
    }
  }
}

// ============ Webhook ============

app.get('/webhook', (req, res) => {
  const { signature, timestamp, nonce, echostr } = req.query;
  
  if (!botConfig.token || !botConfig.aesKey) {
    return res.send(echostr);
  }
  
  const sig = getSignature(botConfig.token, timestamp, nonce, echostr);
  if (sig !== signature) {
    console.log('❌ 签名验证失败');
    return res.status(403).send('error');
  }
  
  const result = decrypt(echostr, botConfig.aesKey);
  console.log('✅ 验证成功:', result);
  res.send(result);
});

app.post('/webhook', async (req, res) => {
  res.send('success');
  
  const { signature, timestamp, nonce } = req.query;
  const encrypt = req.body.encrypt;
  
  if (!encrypt) {
    await handleMessage(req.body);
    return;
  }
  
  const sig = getSignature(botConfig.token, timestamp, nonce, encrypt);
  if (sig !== signature) {
    console.log('❌ 签名失败');
    return;
  }
  
  const decryptStr = decrypt(encrypt, botConfig.aesKey);
  if (!decryptStr) {
    console.log('❌ 解密失败');
    return;
  }
  
  console.log('📝 解密:', decryptStr.substring(0, 100));
  
  try {
    const msg = JSON.parse(decryptStr);
    await handleMessage(msg);
  } catch (e) {
    console.log('❌ 解析失败:', e.message);
  }
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ============ 启动 ============

app.listen(botConfig.port, '0.0.0.0', () => {
  console.log(`\n🤖 ${botConfig.name} 启动 | 端口: ${botConfig.port}\n`);
});

module.exports = app;
