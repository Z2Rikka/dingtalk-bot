/**
 * 钉钉文档收集机器人 v8
 * 修复API调用和消息解析
 */

require('dotenv').config();

const express = require('express');
const axios = require('axios');
const fs = require('fs');
const path = require('path');
const zlib = require('zlib');
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

if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 请配置 .env 文件');
  process.exit(1);
}

console.log('✅ 配置加载成功');

// ============ 常量 ============

const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;
const app = express();
app.use(express.json({ verify: (req, res, buf) => { req.rawBody = buf; } }));

const baseDir = path.resolve(storageConfig.baseDir);
if (!fs.existsSync(baseDir)) fs.mkdirSync(baseDir, { recursive: true });

// ============ 工具函数 ============

function getCurrentDate() {
  return new Date(Date.now() + TIMEZONE_OFFSET);
}

function formatDate(date) {
  return `${date.getFullYear()}${String(date.getMonth() + 1).padStart(2, '0')}${String(date.getDate()).padStart(2, '0')}`;
}

async function getAccessToken() {
  const res = await axios.post('https://api.dingtalk.com/v1.0/oauth2/accessToken', {
    appKey: botConfig.appKey,
    appSecret: botConfig.appSecret
  });
  return res.data.accessToken;
}

function loadLog() {
  try { return fs.existsSync('./download_log.json') ? JSON.parse(fs.readFileSync('./download_log.json', 'utf-8')) : []; } catch { return []; }
}

function saveLog(logs) {
  try { fs.writeFileSync('./download_log.json', JSON.stringify(logs, null, 2)); } catch {}
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

async function downloadFile(content, fileName, accessToken) {
  const { downloadCode, fileId } = content;
  
  console.log(`📥 下载: ${fileName}`);
  console.log(`   downloadCode: ${downloadCode ? '有' : '无'}, fileId: ${fileId || '无'}`);
  
  // 方法1: 通过 tmpCode
  if (downloadCode) {
    try {
      console.log('   方式1: downloadByTmpCode');
      const res = await axios.post(
        'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
        { tmpCode: downloadCode },
        { 
          headers: { 
            'x-acs-dingtalk-access-token': accessToken,
            'Content-Type': 'application/json',
            'Accept-Encoding': 'gzip'
          }, 
          responseType: 'stream',
          decompress: false
        }
      );
      
      if (res.headers['content-type']?.includes('application/json')) {
        let data = '';
        const stream = res.data;
        const gunzip = zlib.createGunzip();
        
        await new Promise((resolve, reject) => {
          stream.pipe(gunzip).on('data', (chunk) => { data += chunk.toString(); })
            .on('end', resolve).on('error', reject);
        });
        
        const error = JSON.parse(data);
        console.log('   方式1返回错误:', error.message || error);
      } else {
        console.log('   方式1成功！');
        return await saveFileStream(res.data, fileName);
      }
    } catch (e) {
      console.log('   方式1错误:', e.message);
    }
  }
  
  // 方法2: 通过 fileId (media)
  if (fileId) {
    try {
      console.log('   方式2: media/download');
      const res = await axios.get(
        `https://api.dingtalk.com/v1.0/robot/media/download?robotId=${botConfig.agentId}&mediaId=${fileId}`,
        { 
          headers: { 
            'x-acs-dingtalk-access-token': accessToken,
            'Accept-Encoding': 'gzip'
          }, 
          responseType: 'stream',
          decompress: false
        }
      );
      
      if (res.headers['content-type']?.includes('application/json')) {
        console.log('   方式2返回JSON');
      } else {
        console.log('   方式2成功！');
        return await saveFileStream(res.data, fileName);
      }
    } catch (e) {
      console.log('   方式2错误:', e.message);
    }
  }
  
  // 方法3: 通过 fileId (robot/file/download)
  if (fileId) {
    try {
      console.log('   方式3: robot/file/download');
      const res = await axios.post(
        'https://api.dingtalk.com/v1.0/robot/file/download',
        { fileId: fileId, robotId: botConfig.agentId },
        { 
          headers: { 
            'x-acs-dingtalk-access-token': accessToken,
            'Content-Type': 'application/json',
            'Accept-Encoding': 'gzip'
          }, 
          responseType: 'stream',
          decompress: false
        }
      );
      
      if (res.headers['content-type']?.includes('application/json')) {
        console.log('   方式3返回JSON');
      } else {
        console.log('   方式3成功！');
        return await saveFileStream(res.data, fileName);
      }
    } catch (e) {
      console.log('   方式3错误:', e.message);
    }
  }
  
  throw new Error('所有方式都失败');
}

async function saveFileStream(stream, originalName) {
  if (!isAllowedExtension(originalName)) {
    throw new Error(`不支持的格式: ${path.extname(originalName)}`);
  }
  
  const dateDir = path.join(baseDir, formatDate(getCurrentDate()));
  if (!fs.existsSync(dateDir)) fs.mkdirSync(dateDir, { recursive: true });
  
  const ext = path.extname(originalName);
  const safeName = sanitizeFileName(path.basename(originalName, ext));
  const fileName = `${Date.now()}_${crypto.randomBytes(4).toHexString()}_${safeName}${ext}`;
  const filePath = path.join(dateDir, fileName);
  
  return new Promise((resolve, reject) => {
    const gunzip = zlib.createGunzip();
    const writer = fs.createWriteStream(filePath);
    stream.pipe(gunzip).pipe(writer);
    writer.on('finish', () => {
      const stats = fs.statSync(filePath);
      resolve({ path: filePath, name: fileName, originalName, size: stats.size, dateDir: formatDate(getCurrentDate()) });
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
  
  const token = await getAccessToken();
  
  if (message.msgtype === 'file') {
    let content = {};
    let fileName = '未知文件';
    
    // 解析 content
    if (message.content) {
      if (typeof message.content === 'string') {
        try { content = JSON.parse(message.content); } catch { content = {}; }
      } else {
        content = message.content;
      }
    }
    
    fileName = content.fileName || message.fileName || '未知文件';
    console.log(`   文件: ${fileName}`);
    
    if (!isAllowedExtension(fileName)) {
      await sendReply(message, `⚠️ 不支持此格式`);
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName, token);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`   ✅ 成功: ${result.name} (${sizeMB}MB)`);
      
      addLog({ type: 'file', originalName: result.originalName, fileName: result.name, size: sizeMB, dateDir: result.dateDir });
      
      await sendReply(message, `✅ 已保存！\n\n文件名: ${result.originalName}\n大小: ${sizeMB} MB\n日期: ${result.dateDir}`);
    } catch (e) {
      console.log(`   ❌ 失败: ${e.message}`);
      await sendReply(message, `❌ 下载失败: ${e.message}`);
    }
  }
  else if (message.msgtype === 'text') {
    let text = '';
    if (typeof message.text === 'string') {
      text = message.text;
    } else if (message.text && message.text.content) {
      text = message.text.content;
    }
    
    console.log(`   文本: ${text.substring(0, 50)}`);
    
    if (text === '/帮助') {
      await sendReply(message, '📖 发送文档自动保存\n支持: PDF,Word,Excel,PPT\n\n命令: /状态 /列表');
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

// ============ Webhook ============

app.get('/webhook', (req, res) => res.send(req.query.echostr));

app.post('/webhook', async (req, res) => {
  res.send('success');
  try {
    await handleMessage(req.body);
  } catch (e) {
    console.error('错误:', e.message);
  }
});

app.get('/health', (req, res) => res.json({ status: 'ok' }));

// ============ 启动 ============

app.listen(botConfig.port, '0.0.0.0', () => {
  console.log(`\n🤖 ${botConfig.name} 已启动 | 端口: ${botConfig.port}\n`);
});

module.exports = app;
