# 钉钉文档收集机器人

一个简洁的钉钉机器人，监听群消息/私聊，自动下载文档到服务器，按日期分类存储。

## 功能

- 📥 自动下载群聊/私聊中的文档
- 📅 按日期（YYYYMMDD）自动分类存储
- 📄 只接受常见文档格式
- 📊 常用管理命令
- 🔌 RESTful API

## 支持的格式

| 类型 | 格式 |
|------|------|
| PDF | .pdf |
| Word | .doc, .docx, .docm, .dotx, .dotm |
| Excel | .xls, .xlsx, .xlsm, .xlsb, .xltx, .xltm |
| PowerPoint | .ppt, .pptx, .pptm, .potx, .potm |
| 其他 | .txt, .csv, .md, .json, .xml, .rar, .zip, .7z |

**注意**: 不接受图片、语音、视频等媒体文件

## 快速开始

### 1. 安装依赖

```bash
cd D:\dingtalk-bot
npm install
```

### 2. 配置

复制配置示例文件并编辑：

```bash
# 复制配置示例
copy .env.example .env
```

编辑 `.env` 文件，填写钉钉凭证：

```env
BOT_APP_KEY=你的AppKey
BOT_APP_SECRET=你的AppSecret
BOT_AGENT_ID=你的AgentId
```

### 3. 启动

```bash
node bot.js
```

或双击 `start.bat`

## 配置说明

### 获取钉钉凭证

1. 访问 [钉钉开放平台](https://open.dingtalk.com/)
2. 创建自建应用
3. 获取：
   - **AppKey**: 凭证与基础信息
   - **AppSecret**: 凭证与基础信息
   - **AgentId**: 应用详情页

4. 在应用权限中开启：
   - 消息→接收消息
   - 群会话→消息推送（群机器人）

### 配置回调

1. 在应用 → 开发管理 → 消息推送
2. 配置回调URL：`http://你的服务器IP:3000/webhook`
3. 启用所有事件回调

### 完整配置说明

所有配置通过 `.env` 文件管理，完整配置项如下：

| 变量名 | 必填 | 说明 | 默认值 |
|--------|------|------|--------|
| BOT_APP_KEY | ✅ | 钉钉开放平台 AppKey | - |
| BOT_APP_SECRET | ✅ | 钉钉开放平台 AppSecret | - |
| BOT_AGENT_ID | ✅ | 钉钉应用 AgentId | - |
| BOT_NAME | - | 机器人名称 | 文档收集助手 |
| BOT_PORT | - | 服务端口 | 3000 |
| STORAGE_BASE_DIR | - | 文件存储目录 | ./received_documents |
| STORAGE_MAX_FILENAME_LENGTH | - | 文件名最大长度 | 200 |
| MESSAGE_AUTO_REPLY | - | 是否自动回复 | true |
| MESSAGE_LISTEN_GROUPS | - | 是否监听群消息 | true |
| MESSAGE_LISTEN_PRIVATE | - | 是否监听私聊 | true |
| MESSAGE_ALLOWED_GROUP_IDS | - | 允许的群ID(逗号分隔) | - |
| LOG_ENABLED | - | 是否启用日志 | true |
| LOG_FILE | - | 日志文件路径 | ./download_log.json |
| LOG_MAX_ENTRIES | - | 日志保留最大条数 | 1000 |

## 存储目录结构

```
received_documents/
├── 20260314/      # 2026年3月14日的文件
│   ├── 1234567890_abcd_document.pdf
│   └── 1234567891_xyz_report.docx
├── 20260315/      # 2026年3月15日的文件
└── ...
```

## 使用

### 发送文档

直接在群聊或私聊中发送文档，机器人会自动下载并按日期存储。

### 命令

在钉钉中发送以下命令：

| 命令 | 说明 |
|------|------|
| `/帮助` | 显示帮助信息 |
| `/状态` | 查看收集统计 |
| `/列表` | 查看最近文件 |
| `/目录` | 查看存储目录 |

## API接口

| 接口 | 说明 |
|------|------|
| `GET /api/files` | 获取文件列表 |
| `GET /api/stats` | 获取统计信息 |
| `GET /api/dirs` | 列出日期目录 |
| `GET /health` | 健康检查 |

## 部署到服务器

```bash
# 上传代码
scp -r ./dingtalk-bot user@your-server:/home/ubuntu/

# SSH登录
ssh user@your-server

# 进入目录
cd ~/dingtalk-bot

# 安装依赖
npm install

# 启动
node bot.js

# 后台运行
nohup node bot.js > bot.log 2>&1 &
```

## 常见问题

### 回调地址需要公网IP怎么办？

使用内网穿透：
- ngrok: `ngrok http 3000`
- frp: 配置frp端口映射

### 文档下载失败

1. 检查应用权限是否开启
2. 检查agentId是否正确
3. 查看日志输出
