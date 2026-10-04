# 🤖 AgentByMyself — 基于智谱大模型的本地办公助手

一个轻量级的本地办公助手 Agent，基于**智谱 GLM 大模型**构建，开箱即用。
无论是文档处理、资料整理，还是日常办公问答，它都能成为你的得力助手。

## ✨ 功能特性

- 🧠 **智谱大模型驱动**：基于智谱 GLM 系列模型，智能理解并执行办公任务
- 💻 **纯本地运行**：代码拉取到本地即可使用，数据不出本机，安全可控
- 🔧 **开箱即用**：只需配置一个 API Key，三步完成部署
- 🚀 **轻量易扩展**：代码结构清晰，便于二次开发和定制自己的 Agent 能力

## 📦 快速开始

### 1. 克隆仓库

```bash
git clone https://github.com/yr-Y04/AgentByMyself.git
cd AgentByMyself
```

### 2. 安装依赖

```bash
pip install -r requirements.txt
```

### 3. 配置智谱 API
在项目根目录创建 .env 文件，填入你的智谱 API Key：
```ini
ZHIPU_API_KEY=你的APIKey
```
💡 API Key 获取方式：前往 智谱 AI 开放平台 注册并创建 API Key。

### 4. 运行
```bash
python main.py
```
