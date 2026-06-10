# 📬 邮箱智能下载助手

> 一款 Windows 桌面工具，批量下载 Outlook 邮件附件，并借助 AI 视觉大模型智能识别附件内容，自动为文件生成有意义的名称。

---

## 功能特性

- **Outlook 邮件集成** — 通过 COM/MAPI 接口直接连接本地 Outlook，无需配置 IMAP/SMTP
- **灵活的邮件筛选** — 支持按发件人、主题关键词、时间范围组合筛选
- **批量下载附件** — 自动保存所有匹配邮件的附件，跳过小于 10KB 的内嵌图片/签名
- **AI 智能命名（核心功能）** — 利用多模态视觉大模型分析图片和 PDF 内容，自动生成描述性文件名
- **8 家 AI 服务商支持** — 内置通义千问、Kimi、智谱、豆包、MiniMax、小米 MiMo、百度千帆，以及自定义 OpenAI 兼容接口
- **现代化 GUI 界面** — tkinter 打造，左右分栏布局，实时日志，DPI 自适应
- **开箱即用** — 提供打包好的 `.exe` 文件，双击即运行，无需安装 Python

## 界面预览

应用采用左右分栏布局：左侧面板配置下载参数和 AI 服务，右侧面板实时显示下载日志。

## 快速开始

### 方式一：直接运行 EXE（推荐）

1. 从 [Releases](../../releases) 下载最新版本
2. 双击 `邮箱智能下载助手.exe` 运行（如需访问 Outlook，可能需要以管理员身份运行）

### 方式二：从源码运行

**环境要求：** Windows + Python 3.8 ~ 3.11 + Microsoft Outlook (2016+)

```bash
# 克隆仓库
git clone https://github.com/your-username/邮箱智能下载助手.git
cd 邮箱智能下载助手

# 安装依赖
pip install pywin32 requests pdf2image pillow

# 运行
python email_attachment_downloader.py
```

> **注意：** PDF 智能命名功能依赖 [Poppler](https://github.com/oschwartz10612/poppler-windows/releases)，请下载并将其 `bin` 目录路径设置到环境变量 `POPPLER_PATH` 中，或放置在程序目录的 `poppler/Library/bin` 下。

### 方式三：自行打包

```bash
# 确保已安装所有依赖，且 Poppler 位于 dist/poppler 目录
pip install pyinstaller
pyinstaller email_attachment_downloader.spec
```

打包产物位于 `dist/邮箱智能下载助手.exe`。

## 使用说明

1. **配置筛选条件**：填写发件人名称（必填，模糊匹配）、主题关键词（可选）、搜索天数
2. **选择保存路径**：点击"浏览"选择附件下载目录
3. **开启 AI 命名**（可选）：
   - 勾选"启用 AI 智能命名"
   - 选择 AI 服务商和模型
   - 填入 API Key（点击"测试连接"验证是否可用）
4. **开始下载**：点击"开始下载"按钮，右侧日志面板实时显示进度

### AI 命名效果

| 文件类型 | 处理方式 | 示例 |
|---------|---------|------|
| 图片 (JPG/PNG/BMP/TIFF) | 直接发送至视觉模型分析 | `IMG_001.jpg` → `公司季度销售报表.jpg` |
| PDF 文档 | 转换为图片（前3页）后分析 | `scan_002.pdf` → `员工劳动合同.pdf` |
| 其他文件 | 保留原始文件名 | 不处理 |

## 支持的 AI 服务商

| 服务商 | 推荐模型 |
|-------|---------|
| 阿里百炼（通义千问） | qwen3-vl-plus |
| 月之暗面（Kimi） | kimi-k2.6 |
| 智谱 AI（GLM） | glm-5v-turbo |
| 字节跳动（豆包） | doubao-seed-2-0-lite-260215 |
| MiniMax | MiniMax-M3 |
| 小米 MiMo | mimo-v2.5 |
| 百度千帆（文心） | ernie-4.5-8k |
| 自定义 | 任意 OpenAI 兼容接口 |

## 配置文件

AI 配置自动保存在：

```
%APPDATA%\OutlookAttachmentDownloader\ai_config.json
```

```json
{
  "provider": "阿里百炼 (通义千问)",
  "model": "qwen3-vl-plus",
  "api_key": "sk-...",
  "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1/chat/completions"
}
```

## 项目结构

```
邮箱智能下载助手/
├── email_attachment_downloader.py    # 主程序（单文件，约 1400 行）
├── email_attachment_downloader.spec  # PyInstaller 打包配置
├── dist/
│   └── 邮箱智能下载助手.exe          # 打包产物
└── README.md
```

## 技术栈

- **语言：** Python
- **GUI：** tkinter + ttk
- **邮件接口：** pywin32 (Outlook COM)
- **AI 调用：** requests（OpenAI 兼容格式）
- **PDF 处理：** pdf2image + Poppler
- **打包：** PyInstaller

## 常见问题

**Q: 启动时报错"无法连接 Outlook"？**
A: 请确保 Microsoft Outlook 已安装并至少启动过一次，且本程序以足够权限运行。

**Q: PDF 文件 AI 命名失败？**
A: 请确认已正确安装 Poppler 并配置了 `POPPLER_PATH` 环境变量。

**Q: 下载的附件去哪了？**
A: 保存在您在程序中指定的目录下，文件名格式为 `YYYY年MM月DD日_HH时MM分SS秒.扩展名`，开启 AI 命名后会自动重命名为内容描述。

## 许可证

本项目采用 [MIT License](LICENSE) 开源许可。

---

如果这个项目对您有帮助，欢迎 Star ⭐ 支持！
