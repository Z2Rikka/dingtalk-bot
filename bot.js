/**
 * 钉钉文档收集机器人 - 支持 Webhook + Stream 双模式
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const crypto = require('crypto');
const { DWClient, EventAck } = require('dingtalk-stream');

// ============ 配置 ============

const config = {
  appKey: process.env.DINGTALK_APP_KEY || process.env.BOT_APP_KEY || '',
  appSecret: process.env.DINGTALK_APP_SECRET || process.env.BOT_APP_SECRET || '',
  agentId: process.env.DINGTALK_AGENT_ID || process.env.BOT_AGENT_ID || '',
  // Webhook 模式需要
  webhookToken: process.env.DINGTALK_WEBHOOK_TOKEN || '',
  webhookAesKey: process.env.DINGTALK_WEBHOOK_AES_KEY || '',
  // 服务器配置
  port: parseInt(process.env.BOT_PORT) || 3565,
  storageDir: process.env.STORAGE_DIR || process.env.STORAGE_BASE_DIR || './received_documents',
  allowedExt: ['.pdf', '.doc', '.docx', '.md', '.txt', '.xls', '.xlsx', '.ppt', '.pptx', '.csv', '.zip', '.rar']
};

if (!config.appKey || !config.appSecret) {
  console.error('❌ 请配置环境变量 DINGTALK_APP_KEY 和 DINGTALK_APP_SECRET');
  process.exit(1);
}

console.log('✅ 配置加载');

// ============ 初始化 ============

const app = express();
app.use(express.json({ limit: '10mb' }));

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
  console.log('🔑 获取Token');
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

// ============ AES 解密 (Webhook 模式) ============

function decrypt(encrypted, aesKey, token) {
  try {
    // AES key 是 base64 编码的
    const key = Buffer.from(aesKey, 'base64');
    
    // 消息体是 JSON 字符串，先 base64 解密，再 AES 解密
    const encryptedBuf = Buffer.from(encrypted, 'base64');
    
    // 使用 AES-256-CBC 解密
    // 前 16 字节是 IV，后面是实际数据
    const iv = encryptedBuf.subarray(0, 16);
    const data = encryptedBuf.subarray(16);
    
    // 去除 PKCS7 padding
    const decipher = crypto.createDecipheriv('aes-256-cbc', key, iv);
    let decrypted = decipher.update(data, null, 'utf8');
    decrypted += decipher.final('utf8');
    
    // 解析 JSON
    const msg = JSON.parse(decrypted);
    
    // 验证 token
    if (msg.token !== token) {
      throw new Error('Token 验证失败');
    }
    
    return msg;
  } catch (e) {
    console.error('解密错误:', e.message);
    throw e;
  }
}

// ============ 下载文件 ============

async function downloadFile(content, fileName, robotCode) {
  const token = await getToken();
  const { downloadCode } = content;
  
  // 使用回调消息中传入的 robotCode，而不是配置的 agentId
  const actualRobotCode = robotCode || config.agentId;
  console.log(`📥 下载: ${fileName}, downloadCode: ${downloadCode}, robotCode: ${actualRobotCode}`);
  
  try {
    // 获取下载链接
    const res = await axios.post(
      'https://api.dingtalk.com/v1.0/robot/messageFiles/download',
      { downloadCode: downloadCode, robotCode: actualRobotCode },
      { headers: { 'x-acs-dingtalk-access-token': token, 'Content-Type': 'application/json' } }
    );
    
    console.log('   API响应:', JSON.stringify(res.data));
    
    const { downloadUrl } = res.data;
    if (!downloadUrl) {
      throw new Error('无法获取下载链接: ' + JSON.stringify(res.data));
    }
    
    // 下载文件
    const fileRes = await axios.get(downloadUrl, { responseType: 'stream' });
    
    const dateDir = path.join(baseDir, getToday());
    if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
    
    const ext = path.extname(fileName);
    const finalName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName(path.basename(fileName, ext))}${ext}`;
    const filePath = path.join(dateDir, finalName);
    
    return new Promise((resolve, reject) => {
      fileRes.data.pipe(fs.createWriteStream(filePath)).on('finish', () => {
        const stats = fs.statSync(filePath);
        resolve({ name: finalName, original: fileName, size: stats.size, date: getToday() });
      }).on('error', reject);
    });
  } catch (e) {
    console.error('   下载错误:', e.message);
    if (e.response) {
      console.error('   响应状态:', e.response.status);
      console.error('   响应数据:', e.response.data);
    }
    throw e;
  }
}

// ============ 发送消息 ============

async function sendText(conversationId, text, conversationType = '1', customRobotId = null) {
  try {
    const token = await getToken();
    const url = conversationType === '2' 
      ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
      : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
    
    // 使用传入的 robotId 或默认的 agentId
    const botId = customRobotId || config.agentId;
    console.log('   发送消息 robotId:', botId);
    
    await axios.post(url, {
      robotId: botId,
      openConversationId: conversationId,
      msgtype: 'text',
      text: { content: text }
    }, { headers: { 'x-acs-dingtalk-access-token': token } });
    console.log('   ✅ 已回复');
  } catch (e) {
    console.log('   ⚠️ 回复失败:', e.message, e.response?.data);
  }
}

// ============ 处理消息 ============

// 从 chatbotUserId 提取真正的 robotCode (格式: $:LWCP_v1:$xxxxx)
function extractRobotCode(chatbotUserId) {
  if (!chatbotUserId) return null;
  // 去掉 $:LWCP_v1: 前缀
  const match = chatbotUserId.match(/^\$:LWCP_v1:\$(.+)$/);
  return match ? match[1] : chatbotUserId;
}

async function processMessage(msg, res) {
  console.log(`\n📩 收到消息处理`);
  console.log('   消息:', JSON.stringify(msg).substring(0, 500));
  
  // 从消息中提取会话信息和机器人标识
  const conversationId = msg.conversationId || msg.openConversationId;
  const conversationType = msg.conversationType || (msg.isGroup ? '2' : '1');
  
  // Webhook 回调中使用 chatbotUserId 作为 robotCode
  const robotCode = extractRobotCode(msg.chatbotUserId);
  console.log('   提取的robotCode:', robotCode);
  
  if (msg.msgtype === 'file') {
    let content = {};
    let fileName = '未知文件';
    
    if (msg.content) {
      try { content = typeof msg.content === 'string' ? JSON.parse(msg.content) : msg.content; } catch { console.log('   内容解析失败'); }
    }
    
    fileName = content.fileName || msg.fileName || '未知文件';
    console.log(`   文件: ${fileName}`);
    console.log('   完整content:', JSON.stringify(content));
    
    if (!isAllowed(fileName)) {
      await sendText(conversationId, '⚠️ 不支持此格式', conversationType, robotCode);
      return;
    }
    
    if (!content.downloadCode) {
      console.log('   ❌ 没有downloadCode');
      await sendText(conversationId, '❌ 无法获取下载凭证', conversationType, robotCode);
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName, robotCode);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`   ✅ 成功: ${result.name} (${sizeMB}MB)`);
      
      const logs = loadLog();
      logs.unshift({ type: 'file', originalName: result.original, fileName: result.name, size: sizeMB, date: result.date, timestamp: new Date().toISOString() });
      saveLog(logs);
      
      await sendText(conversationId, `✅ 已保存！\n\n文件名: ${result.original}\n大小: ${sizeMB} MB\n日期: ${result.date}`, conversationType, robotCode);
    } catch (e) {
      console.log(`   ❌ 失败: ${e.message}`);
      await sendText(conversationId, `❌ 下载失败: ${e.message.substring(0, 50)}`, conversationType, robotCode);
    }
  }
  else if (msg.msgtype === 'text') {
    const text = typeof msg.text === 'string' ? msg.text : (msg.text?.content || '').trim();
    console.log(`   文本: ${text}`);
    
    if (text === '/帮助') await sendText(conversationId, '📖 发送文档自动保存\n支持: PDF,Word,Excel,MD\n\n命令: /状态 /列表', conversationType, robotCode);
    else if (text === '/状态') {
      const logs = loadLog().filter(l => l.type === 'file');
      const total = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendText(conversationId, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${total} MB`, conversationType, robotCode);
    }
    else if (text === '/列表') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let m = '📋 最近:\n\n';
      logs.forEach((l, i) => m += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.date}\n\n`);
      await sendText(conversationId, m || '暂无', conversationType, robotCode);
    }
    else {
      await sendText(conversationId, '收到！发送文件自动保存', conversationType, robotCode);
    }
  }
  else {
    console.log('   未处理的消息类型:', msg.msgtype);
  }
}

// ============ Webhook 模式 ============

// Webhook 回调地址
app.post('/webhook', async (req, res) => {
  console.log('\n========== 收到 Webhook 回调 ==========');
  console.log('Query:', req.query);
  console.log('Body:', JSON.stringify(req.body).substring(0, 500));
  
  try {
    const { signature, msg_signature, timestamp, nonce } = req.query;
    const { encrypt } = req.body;
    
    // 如果有加密内容，需要解密
    if (encrypt && config.webhookAesKey && config.webhookToken) {
      try {
        const decrypted = decrypt(encrypt, config.webhookAesKey, config.webhookToken);
        console.log('解密后:', JSON.stringify(decrypted).substring(0, 500));
        
        // 解密后的消息格式可能不同
        if (decrypted.eventType === 'im.message' || decrypted.msgtype) {
          await processMessage(decrypted, res);
        }
      } catch (e) {
        console.error('解密失败:', e.message);
        // 尝试直接处理（测试阶段可能不加密）
        if (req.body.msgtype) {
          await processMessage(req.body, res);
        }
      }
    } else if (req.body.msgtype) {
      // 未加密的直接消息
      await processMessage(req.body, res);
    }
    
    res.json({ success: true });
  } catch (e) {
    console.error('Webhook 处理错误:', e);
    res.status(500).json({ error: e.message });
  }
});

// ============ Stream 模式 ============

let reconnectTimer = null;
let isConnected = false;

const client = new DWClient({
  clientId: config.appKey,
  clientSecret: config.appSecret,
});

const onEventReceived = (event) => {
  console.log('\n========== 收到 Stream 事件 ==========');
  console.log('headers:', JSON.stringify(event.headers));
  console.log('data类型:', typeof event.data);
  console.log('data:', typeof event.data === 'string' ? event.data : JSON.stringify(event.data).substring(0, 1000));
  
  try {
    const topic = event.headers?.topic || event.headers?.eventType;
    console.log('topic:', topic);
    
    if (topic === 'im.message.receive' || topic === 'im.messageReceive') {
      let content;
      
      if (typeof event.data === 'string') {
        try {
          content = JSON.parse(event.data);
        } catch {
          content = event.data;
        }
      } else if (typeof event.data === 'object') {
        content = event.data;
      }
      
      console.log('解析后content:', JSON.stringify(content).substring(0, 500));
      
      if (content && content.msgtype) {
        processMessage(content);
      } else {
        console.log('没有msgtype，跳过');
      }
    } else {
      console.log('不处理此topic');
    }
  } catch (e) {
    console.log('   解析错误:', e.message);
    console.log(e.stack);
  }
  
  return { status: EventAck.SUCCESS, message: 'OK' };
};

function connect() {
  client
    .registerAllEventListener(onEventReceived)
    .connect()
    .then(() => {
      console.log('✅ Stream 连接成功');
      isConnected = true;
      if (reconnectTimer) {
        clearTimeout(reconnectTimer);
        reconnectTimer = null;
      }
    })
    .catch((err) => {
      console.error('❌ Stream 连接失败:', err.message);
      isConnected = false;
      scheduleReconnect();
    });
}

function scheduleReconnect() {
  if (reconnectTimer) return;
  console.log('⏰ 5秒后尝试重连...');
  reconnectTimer = setTimeout(() => {
    reconnectTimer = null;
    connect();
  }, 5000);
}

client.on('close', () => {
  console.log('❌ Stream 连接已关闭，准备重连...');
  isConnected = false;
  scheduleReconnect();
});

client.on('error', (err) => {
  console.error('❌ Stream 错误:', err.message);
  isConnected = false;
  scheduleReconnect();
});

// 尝试连接 Stream（可选）
if (config.appKey && config.appSecret) {
  connect();
}

// ============ HTTP 服务器 ============

app.get('/health', (req, res) => res.json({ 
  status: 'ok', 
  mode: 'webhook+stream',
  streamConnected: isConnected 
}));

app.listen(config.port, '0.0.0.0', () => {
  console.log(`\n🤖 文档收集助手 | 端口: ${config.port}`);
  console.log(`📍 Webhook: http://<your-server>:${config.port}/webhook`);
  console.log(`📍 Health:  http://<your-server>:${config.port}/health\n`);
});

module.exports = app;
