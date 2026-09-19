# Report Agent Web Plus

基于 DeepSeek 的企业汇报与 PowerPoint 智能生成工具。

项目提供通用问答、汇报材料整理和PPT文件生成能力，并配备Windows一键启动脚本，方便没有开发环境的用户在本地运行。

## 项目简介

Report Agent Web Plus 面向工作汇报、月报总结、项目复盘和演示材料制作等场景。

用户可以通过网页输入问题、工作记录或PPT制作要求，由智能体完成内容理解、结构整理和幻灯片生成，并输出可下载的 `.pptx` 文件。

## 核心功能

### 通用问答

- 接入 DeepSeek 大模型
- 支持连续对话和历史会话管理
- 可用于信息问答、内容梳理和表达优化
- 大模型暂时不可用时提供基础本地响应

### 汇报材料整理

- 将零散工作记录整理为正式汇报内容
- 自动归纳总体概述、重点工作和阶段成果
- 整理存在问题及下一阶段计划
- 尽量保留原始事实，避免随意编造数据和项目名称

### PPT智能生成

- 根据主题、材料和受众规划PPT结构
- 自动生成标题页、目录页、内容页和结束页
- 支持企业PPT模板
- 根据内容选择相应版式
- 自动生成页面标题、要点和演讲备注
- 输出可直接下载的 `.pptx` 文件

### 本地一键启动

项目提供 `start.bat`，可自动完成：

- 检查本机Node.js环境
- 检查并安装项目依赖
- 检查环境变量配置
- 启动本地Web服务
- 自动打开浏览器
- 关闭窗口时停止本地服务

## 技术栈

- 前端：HTML、CSS、JavaScript
- 后端：Node.js、Express
- 大模型：DeepSeek API
- 演示文稿：基于PPTX文件结构生成和填充
- 配置管理：dotenv
- 本地启动：Windows Batch Script

## 工作流程

```text
用户输入问题或PPT制作要求
             ↓
网页将请求发送至Node.js后端
             ↓
DeepSeek理解需求并生成结构化内容
             ↓
后端根据内容组织汇报结构
             ↓
调用企业模板或默认版式生成PPTX
             ↓
用户在网页中下载PPT文件
```

## 项目结构

```text
report-agent-web-plus/
├── public/
│   ├── index.html          # 网页结构
│   ├── style.css           # 页面样式
│   └── app.js              # 前端交互与接口调用
├── prompts/                # PPT内容与排版提示词
├── templates/              # PPT模板文件
├── server.js               # Express服务与PPT生成逻辑
├── start.bat               # Windows一键启动脚本
├── package.json            # 项目依赖和运行命令
└── .env                    # 本地环境变量，请勿上传
```

## 本地运行

### 方式一：一键启动

确保电脑已安装Node.js，然后双击：

```text
start.bat
```

脚本完成检查后会启动服务，并自动打开浏览器。

默认访问地址：

```text
http://localhost:3000
```

关闭启动窗口后，本地服务也会随之停止。

### 方式二：手动启动

安装项目依赖：

```bash
npm install
```

在项目根目录创建 `.env` 文件：

```env
PORT=3000

DEEPSEEK_API_KEY=your-deepseek-api-key
DEEPSEEK_BASE_URL=https://api.deepseek.com
DEEPSEEK_MODEL=deepseek-chat
DEEPSEEK_TIMEOUT_MS=30000
DEEPSEEK_MAX_RETRIES=2
```

启动服务：

```bash
npm start
```

开发模式：

```bash
npm run dev
```

## 使用方法

1. 启动项目并打开本地网页。
2. 选择“问答”或“生成PPT”。
3. 输入问题、汇报材料或PPT制作要求。
4. 在制作PPT时，建议说明主题、受众、页数、材料重点和期望风格。
5. 等待智能体整理内容并生成文件。
6. 点击“下载PPT”保存生成结果。
7. 正式使用前检查事实、数据、措辞和版面。

示例输入：

```text
请根据以下材料制作一份8页的季度工作汇报PPT，
面向部门管理层，重点展示工作进展、阶段成果、
现存问题和下一季度计划，整体风格要求简洁、正式。
```

## 环境变量

| 变量 | 作用 | 是否必需 |
|---|---|---|
| `PORT` | 本地服务监听端口 | 否 |
| `DEEPSEEK_API_KEY` | DeepSeek API密钥 | 是 |
| `DEEPSEEK_BASE_URL` | DeepSeek接口地址 | 否 |
| `DEEPSEEK_MODEL` | 使用的模型名称 | 否 |
| `DEEPSEEK_TIMEOUT_MS` | 请求超时时间 | 否 |
| `DEEPSEEK_MAX_RETRIES` | 请求失败后的重试次数 | 否 |
| `PPT_TEMPLATE_PATH` | 自定义PPT模板路径 | 否 |

## 安全说明

请勿将以下文件上传到公开仓库：

```text
.env
node_modules/
server.log
*.msi
```

真实API密钥只能保存在本地 `.env` 或云平台的环境变量中，不应直接写入前端代码或提交至GitHub。

如PPT模板包含企业品牌资产、内部信息或版权材料，请取得相应授权后再公开。

建议在 `.gitignore` 中加入：

```gitignore
.env
node_modules/
server.log
*.log
*.msi
.DS_Store
Thumbs.db
```

## 当前限制

- 上传附件目前主要记录文件名，尚未完整解析附件正文。
- 如需大模型理解文件内容，建议将关键文字粘贴到输入框。
- AI生成内容可能存在遗漏或表达偏差。
- 复杂PPT仍需要人工检查和适当美化。
- 生成结果不应直接替代事实核验和业务审核。

## 后续规划

- 自动解析Word、Excel和PowerPoint附件
- 合并问答与PPT制作入口并自动识别用户意图
- 增加PPT视觉质检智能体
- 支持更多企业PPT模板
- 优化图片、图表和版式生成能力
- 增加生成过程和修改意见展示

## 作者

万晓凯

GitHub：[wanxiaokai-0413](https://github.com/wanxiaokai-0413)
