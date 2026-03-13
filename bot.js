/**
 * 钉钉文档收集机器人 v4
 * 修复文件下载问题
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

if (!botConfig.appKey || !botConfig.appSecret || !botConfig.agentId) {
  console.error('❌ 请配置 .env 文件');
  process.exit(1);
}

console.log('✅ 配置加载成功');

// ============ 常量 ============

const TIMEZONE_OFFSET = 8 * 60 * 60 * 1000;
const app = express();
app.use(express.json());

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

// ============ 下载文件 - 尝试多种方式 ============

async function downloadFile(content, fileName, accessToken) {
  const downloadCode = content.downloadCode;
  const spaceId = content.spaceId;
  const fileId = content.fileId;
  
  console.log(`📥 尝试下载: ${fileName}`);
  console.log(`   downloadCode: ${downloadCode ? '有' : '无'}`);
  console.log(`   spaceId: ${spaceId ? '有' : '无'}`);
  console.log(`   fileId: ${fileId ? '有' : '无'}`);
  
  // 方法1: 使用 tmpCode 下载
  if (downloadCode) {
    try {
      console.log('   尝试方法1: tmpCode下载...');
      const res = await axios.post(
        'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
        { tmpCode: downloadCode },
        { 
          headers: { 
            'x-acs-dingtalk-access-token': accessToken,
            'Content-Type': 'application/json'
          },
          responseType: 'stream'
        }
      );
      
      // 检查是否返回错误
      const contentType = res.headers['content-type'] || '';
      if (contentType.includes('application/json')) {
        const chunks = [];
        for await (const chunk of res.data) chunks.push(chunk);
        const errorBody = Buffer.concat(chunks).toString();
        console.log('   方法1失败:', errorBody);
      } else {
        return await saveFile(res.data, fileName);
      }
    } catch (e) {
      console.log('   方法1错误:', e.response?.data || e.message);
    }
  }
  
  // 方法2: 通过 spaceId 获取下载链接
  if (spaceId && fileId) {
    try {
      console.log('   尝试方法2: 获取下载链接...');
      const res = await axios.post(
        'https://api.dingtalk.com/v1.0/robot/file/getUploadUrl',
        {
          spaceId: spaceId,
          fileId: fileId,
          fileName: fileName
        },
        { headers: { 'x-acs-dingtalk-access-token': accessToken } }
      );
      
      if (res.data.downloadUrl) {
        console.log('   获取到下载链接，下载中...');
        const fileRes = await axios.get(res.data.downloadUrl, { responseType: 'stream' });
        return await saveFile(fileRes.data, fileName);
      }
    } catch (e) {
      console.log('   方法2错误:', e.response?.data || e.message);
    }
  }
  
  // 方法3: 使用旧版API (webhook)
  if (downloadCode) {
    try {
      console.log('   尝试方法3: 通过OAuth token...');
      const res = await axios.post(
        'https://api.dingtalk.com/v1.0/robot/message/files/downloadByTmpCode',
        { tmpCode: downloadCode },
        { 
          headers: { 
            'x-acs-dingtalk-access-token': accessToken,
            'Content-Type': 'application/json'
          },
          responseType: 'stream'
        }
      );
      return await saveFile(res.data, fileName);
    } catch (e) {
      console.log('   方法3错误:', e.response?.data || e.message);
    }
  }
  
  throw new Error('无法下载：所有方法都失败');
}

async function saveFile(stream, originalName) {
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
    const writer = fs.createWriteStream(filePath);
    stream.pipe(writer);
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
}

// ============ 发送消息 ============

async function sendReply(message, content) {
  try {
    // 优先用 sessionWebhook
    if (message.sessionWebhook) {
      await axios.post(message.sessionWebhook, {
        msgtype: 'text',
        text: { content }
      });
      return;
    }
    
    // 用 openConversationId
    if (message.conversationId) {
      const token = await getAccessToken();
      const url = message.conversationType === '2' 
        ? 'https://api.dingtalk.com/v1.0/robot/oToMessages/batchSend'
        : 'https://api.dingtalk.com/v1.0/robot/oToMessages/send';
      
      await axios.post(url, {
        robotId: botConfig.agentId,
        openConversationId: message.conversationId,
        msgtype: 'text',
        text: { content }
      }, { headers: { 'x-acs-dingtalk-access-token': token } });
    }
  } catch (e) {
    console.log('⚠️ 回复失败:', e.response?.data || e.message);
  }
}

// ============ 处理消息 ============

async function handleMessage(message) {
  console.log(`\n📩 ${message.msgtype} | ${message.conversationType === '2' ? '群聊' : '私聊'}`);
  
  const accessToken = await getAccessToken();
  
  if (message.msgtype === 'file') {
    const content = message.content || {};
    const fileName = content.fileName || message.fileName || '未知文件';
    
    if (!isAllowedExtension(fileName)) {
      await sendReply(message, `⚠️ 不支持此格式`);
      return;
    }
    
    try {
      const result = await downloadFile(content, fileName, accessToken);
      const sizeMB = (result.size / 1024 / 1024).toFixed(2);
      
      console.log(`✅ 成功: ${result.name} (${sizeMB}MB)`);
      
      addLog({
        type: 'file',
        originalName: result.originalName,
        fileName: result.name,
        size: sizeMB,
        dateDir: result.dateDir
      });
      
      await sendReply(message, `✅ 已保存！\n\n文件名: ${result.originalName}\n大小: ${sizeMB} MB\n日期: ${result.dateDir}`);
      
    } catch (e) {
      console.log(`❌ 失败: ${e.message}`);
      await sendReply(message, `❌ 下载失败: ${e.message}`);
    }
  }
  else if (message.msgtype === 'text') {
    const text = typeof message.text === 'string' ? message.text : (message.text?.content || '');
    
    if (text === '/帮助' || text === '/help') {
      await sendReply(message, '📖 发送文档自动保存\n支持: PDF,Word,Excel,PPT,TXT\n\n命令: /状态 /列表 /目录 /帮助');
    }
    else if (text === '/状态' || text === '/stats') {
      const logs = loadLog().filter(l => l.type === 'file');
      const totalSize = logs.reduce((s, l) => s + parseFloat(l.size || 0), 0).toFixed(2);
      await sendReply(message, `📊 统计\n\n文档: ${logs.length} 个\n大小: ${totalSize} MB`);
    }
    else if (text === '/列表' || text === '/list') {
      const logs = loadLog().filter(l => l.type === 'file').slice(0, 10);
      let msg = '📋 最近:\n\n';
      logs.forEach((l, i) => msg += `${i+1}. ${l.originalName}\n${l.size}MB | ${l.dateDir}\n\n`);
      await sendReply(message, msg || '暂无');
    }
    else if (text === '/目录' || text === '/dir') {
      try {
        const dirs = fs.readdirSync(baseDir).filter(d => /^\d{8}$/.test(d)).sort().reverse();
        await sendReply(message, `📁 ${baseDir}\n\n${dirs.slice(0,10).map(d => '• ' + d).join('\n')}`);
      } catch {}
    }
  }
}

// ============ Webhook ============

app.get('/webhook', (req, res) => {
  res.send(req.query.echostr);
});

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
