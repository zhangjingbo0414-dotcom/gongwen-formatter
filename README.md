# 公文格式化工具 - Web 前端

基于 Streamlit 框架的公文格式化 Web 界面，支持一键生成符合 **GB/T 9704-2012** 规范的公文 Word 文档。

## 🚀 快速启动

### 1. 安装依赖

```bash
pip install streamlit python-docx
```

### 2. 启动应用

```bash
cd 公文格式化
streamlit run app.py
```

启动后浏览器会自动打开 `http://localhost:8501`。

如需指定端口或允许外部访问：

```bash
streamlit run app.py --server.port 8501 --server.address 0.0.0.0
```

## 📖 使用说明

### 文本输入

在左侧文本框中粘贴公文内容，系统会自动识别以下结构：

| 结构 | 识别规则 | 格式 |
|------|----------|------|
| 主标题 | 第一个非空行 | 方正小标宋简体/黑体，二号，居中 |
| 发文字号 | 匹配 `〔YYYY〕N号` | 仿宋，三号，居中 |
| 一级标题 | `一、` `二、` 等 | 黑体，三号 |
| 二级标题 | `（一）` `（二）` 等 | 楷体，三号 |
| 三级标题 | `1.` `2.` 等 | 仿宋加粗，三号 |
| 正文 | 其余非空行 | 仿宋，三号，首行缩进2字符 |
| 附件 | 以"附件"开头 | 仿宋，三号，左对齐 |
| 表格 | Markdown 格式（`\|...\|`） | 黑体表头+仿宋内容，居中 |

### 文件上传

支持上传 `.txt` 或 `.docx` 文件，内容会自动提取并填入文本框。

### 生成与下载

1. 点击 **「一键生成公文」** 按钮
2. 等待生成完成（通常几秒）
3. 点击 **「下载公文文档」** 下载生成的 `.docx` 文件

## 📁 文件结构

```
公文格式化/
├── app.py                  # Streamlit Web 前端
├── gongwen_formatter.py    # 公文格式化核心模块
├── requirements.txt        # Python 依赖
├── .streamlit/
│   └── config.toml         # Streamlit 主题配置
├── .gitignore              # Git 忽略规则
├── README.md               # 本说明文件
└── 公文_YYYYMMDD_HHMMSS.docx  # 生成的公文文档（已忽略）
```

## ⚙️ 技术说明

- **前端框架**：Streamlit 1.57+
- **文档生成**：python-docx
- **格式标准**：GB/T 9704-2012《党政机关公文格式》
- **模块调用**：通过 `import gongwen_formatter` 直接调用，无需 subprocess

## ☁️ Streamlit Community Cloud 部署

本应用支持免费部署到 [Streamlit Community Cloud](https://streamlit.io/cloud)，无需服务器即可在线使用。

### 部署步骤

1. **准备 GitHub 仓库**
   - 将 `公文格式化` 目录推送到 GitHub 仓库（可 fork 或新建仓库）
   - 确保仓库根目录包含 `app.py`、`requirements.txt` 和 `.streamlit/config.toml`

2. **登录 Streamlit Cloud**
   - 访问 [share.streamlit.io](https://share.streamlit.io)
   - 使用 GitHub 账号登录

3. **部署应用**
   - 点击 **"New app"**
   - 选择对应的 GitHub 仓库
   - 设置主文件路径为 `app.py`
   - 点击 **"Deploy"**

4. **访问应用**
   - 部署完成后，系统会分配一个 `https://<app-name>.streamlit.app` 的访问地址
   - 分享该地址即可让他人在线使用

### 免费部署说明

- Streamlit Community Cloud 提供 **免费** 托管，无需信用卡
- 每个账号可部署多个应用
- 应用在无访问时会自动休眠，首次访问需等待几秒唤醒
- 资源有限，适合个人或小型团队使用，不建议用于高并发生产环境

### 注意事项

- 部署后生成的文档中字体可能因服务器环境缺少中文字体而显示为替代字体，建议用户下载后在本地 Word/WPS 中查看最终效果
- 如需自定义域名，可参考 [Streamlit 官方文档](https://docs.streamlit.io/streamlit-community-cloud/get-started/custom-domains)

## ⚠️ 注意事项

1. 生成的文档中字体依赖本机安装的字体，若缺少「方正小标宋简体」「仿宋_GB2312」等字体，Word 会以替代字体显示
2. 本工具仅提供格式化排版，不涉及公文内容的合法性审查
3. 生成的文档建议在 Microsoft Word 或 WPS 中打开查看完整效果
