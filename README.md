# 📧 BulkMailer - Python 批量邮件发送工具

<p align="center">
  <img src="https://img.shields.io/github/stars/yourusername/BulkMailer?style=for-the-badge&color=blue" alt="GitHub stars">
  <img src="https://img.shields.io/github/forks/yourusername/BulkMailer?style=for-the-badge&color=green" alt="GitHub forks">
  <img src="https://img.shields.io/github/license/yourusername/BulkMailer?style=for-the-badge&color=orange" alt="License">
  <img src="https://img.shields.io/badge/Python-3.8%2B-blue?style=for-the-badge&logo=python" alt="Python">
  <img src="https://img.shields.io/badge/PyQt5-GUI-success?style=for-the-badge&logo=qt" alt="PyQt5">
  <img src="https://img.shields.io/badge/Platform-Windows%20%7C%20Linux%20%7C%20MacOS-lightgrey?style=for-the-badge" alt="Platform">
</p>

<p align="center">
  <a href="#-项目简介">项目简介</a> •
  <a href="#-功能特性">功能特性</a> •
  <a href="#-截图预览">截图预览</a> •
  <a href="#-快速开始">快速开始</a> •
  <a href="#-使用方法">使用方法</a> •
  <a href="#-配置说明">配置说明</a> •
  <a href="#-打包指南">打包指南</a> •
  <a href="#-常见问题">常见问题</a> •
  <a href="#-贡献指南">贡献指南</a> •
  <a href="#-许可证">许可证</a> •
  <a href="#-联系作者">联系作者</a>
</p>

<p align="center">
  <img src="https://github-readme-stats.vercel.app/api/pin/?username=yourusername&repo=BulkMailer&theme=radical&show_icons=true" alt="GitHub Stats">
</p>

---

## 🚀 项目简介

**BulkMailer** 是一款基于 **Python + PyQt5** 开发的现代化批量邮件发送工具，专为需要高效、稳定、美观的邮件发送场景设计。无论是营销推广、通知提醒、活动邀请，还是日常办公，BulkMailer 都能提供强大的邮件发送能力和友好的图形化界面。

> ✨ **核心优势**：支持多账号轮询发送、HTML 富文本邮件、Excel 收件人列表导入、实时发送进度监控，并已打包为独立可执行文件，无需安装 Python 环境即可使用！

---

## ✨ 功能特性

### 📋 核心功能
- ✅ **批量发送**：支持 Excel/CSV 列表导入，一键群发上千封邮件
- ✅ **多账号轮询**：自动切换发件人账号，避免单个账号发送限制
- ✅ **HTML 富文本**：支持 HTML 格式邮件，可插入图片、链接、样式
- ✅ **附件支持**：可添加任意类型附件（图片、文档、压缩包等）
- ✅ **实时进度**：图形化进度条显示发送状态，成功/失败一目了然
- ✅ **错误重试**：自动记录发送失败邮箱，支持手动重试

### 🛠 技术特色
- 🎨 **现代化 GUI**：基于 PyQt5 开发，界面美观、操作流畅
- 🔒 **安全稳定**：采用 SMTP SSL 加密传输，保障账号安全
- 📊 **数据统计**：发送完成后生成详细统计报告
- 🚀 **高性能**：多线程发送，充分利用系统资源
- 📦 **独立可执行**：已打包为 .exe 文件，开箱即用

### 🌐 邮箱支持
- QQ 邮箱 (@qq.com)
- 163 邮箱 (@163.com)
- 126 邮箱 (@126.com)
- Gmail (需配置)
- 其他 SMTP 服务商（可自定义配置）

---

## 🖼 截图预览

> ⚠️ **截图待补充**：请将程序运行截图放入 screenshots/ 目录，并更新以下链接。

| 主界面 | 收件人管理 | 发送进度 |
|--------|------------|----------|
| ![主界面](screenshots/main.png) | ![收件人管理](screenshots/recipients.png) | ![发送进度](screenshots/progress.png) |

*提示：您可以使用快捷键 Alt+PrtSc 截取窗口，将图片保存至 screenshots 文件夹并更新上述链接。*

---

## ⚡ 快速开始

### 环境要求
- **Python 3.8+**（仅开发需要）
- **Windows 7/10/11**（已打包版本）
- **Linux/MacOS**（需源码运行）

### 方法一：使用已打包版本（推荐）
1. 前往 [Release 页面](https://github.com/yourusername/BulkMailer/releases) 下载最新版 邮件发送工具.exe
2. 双击运行即可，无需安装任何依赖

### 方法二：从源码运行
\\\ash
# 1. 克隆仓库
git clone https://github.com/yourusername/BulkMailer.git
cd BulkMailer

# 2. 安装依赖
pip install -r requirements.txt

# 3. 运行程序
python Email_sender_pyqt5.py
\\\

---

## 📖 使用方法

### 1. 准备发件人账号
- 支持多个发件人账号，格式为 邮箱,密码,昵称(可选)
- 例如：
  \\\
  sender1@qq.com,password123,客服小张
  sender2@163.com,password456
  \\\

### 2. 导入收件人列表
- 支持 Excel (.xlsx) 和 CSV 格式
- 文件需包含 邮箱 列，可包含 姓名、公司 等字段用于邮件个性化

### 3. 编辑邮件内容
- **主题**：填写邮件主题
- **正文**：支持纯文本和 HTML 格式
- **附件**：点击“添加附件”按钮选择文件

### 4. 开始发送
- 点击“开始发送”按钮
- 实时查看发送进度和成功/失败统计
- 发送完成后可导出失败列表

### 5. 高级功能
- **延迟发送**：设置每封邮件发送间隔，避免被识别为垃圾邮件
- **个性化变量**：在邮件正文中使用 {姓名}、{公司} 等变量
- **测试发送**：先发送一封测试邮件验证配置

---

## ⚙️ 配置说明

### 配置文件
程序支持通过 config.ini 文件进行高级配置（首次运行后自动生成）：

\\\ini
[email]
smtp_server = smtp.qq.com
smtp_port = 465
default_sender = your_email@qq.com
default_nickname = 默认发件人

[sending]
delay_between_emails = 1.5
max_retries = 3
timeout = 30

[ui]
language = zh_CN
theme = light
font_size = 12
\\\

### 环境变量
\\\ash
# 设置默认发件人（优先级高于配置文件）
export BULKMAILER_DEFAULT_SENDER="your_email@qq.com"
export BULKMAILER_DEFAULT_PASSWORD="your_password"
\\\

---

## 📦 打包指南

如果您想自行打包为可执行文件：

### 使用 PyInstaller
\\\ash
# 安装 PyInstaller
pip install pyinstaller

# 打包（单文件）
pyinstaller --onefile --windowed --icon=favicon.ico Email_sender_pyqt5.py

# 打包（带控制台，用于调试）
pyinstaller --console --icon=favicon.ico Email_sender_pyqt5.py
\\\

### 打包注意事项
1. 确保所有资源文件（图标、图片）在 spec 文件中正确配置
2. 测试打包后的程序是否能在其他电脑正常运行
3. 可使用 --add-data 参数添加数据文件

---

## ❓ 常见问题

### Q1: 发送失败，提示“认证失败”？
A: 请检查：
- 邮箱密码是否正确（QQ 邮箱需使用授权码，非登录密码）
- 是否开启了 SMTP 服务（QQ 邮箱需在设置中开启）
- 防火墙是否阻止了程序连接

### Q2: 如何提高发送成功率？
A:
- 使用多个发件人账号轮流发送
- 设置合理的发送间隔（建议 1-3 秒）
- 避免邮件内容包含敏感词汇
- 使用真实的发件人昵称和签名

### Q3: 支持海外邮箱吗？
A: 支持，但需要手动配置 SMTP 服务器和端口。可在代码中扩展 get_smtp_server 函数。

### Q4: 如何实现邮件个性化？
A: 在 Excel 收件人列表中增加 姓名、公司 等列，在邮件正文中使用 {姓名} 占位符。

### Q5: 程序卡顿或无响应？
A: 发送大量邮件时，程序会占用较多资源，这是正常现象。建议：
- 分批发送（每次不超过 500 封）
- 关闭其他占用资源的程序
- 使用性能更好的电脑

---

## 🤝 贡献指南

我们欢迎任何形式的贡献！以下是参与项目的方式：

### 报告问题
- 使用 [GitHub Issues](https://github.com/yourusername/BulkMailer/issues) 报告 bug 或提出建议
- 请提供详细的重现步骤、预期行为和实际行为

### 提交代码
1. Fork 本仓库
2. 创建功能分支 (\git checkout -b feature/AmazingFeature\)
3. 提交更改 (\git commit -m 'Add some AmazingFeature'\)
4. 推送到分支 (\git push origin feature/AmazingFeature\)
5. 开启 Pull Request

### 开发规范
- 遵循 PEP 8 代码风格
- 为新增功能编写文档
- 确保代码通过基本测试

### 待开发功能
- [ ] 邮件模板系统
- [ ] 发送计划（定时发送）
- [ ] 邮件打开率统计
- [ ] 多语言支持（英文、日文等）
- [ ] 云端配置同步

---

## 📄 许可证

本项目采用 **MIT License** - 查看 [LICENSE](LICENSE) 文件了解详情。

\\\
MIT License

Copyright (c) 2025 寂寞沙洲冷

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
\\\

---

## 📞 联系作者

**寂寞沙洲冷** (QQ/VX: 1638276310)

- 📧 邮箱：\your_email@example.com\
- 🐱 GitHub：[@yourusername](https://github.com/yourusername)
- 📖 博客：[个人技术博客](https://blog.example.com)

### 支持与赞助
如果这个项目对您有帮助，欢迎：

- ⭐ **Star 本项目** - 让更多人看到
- 🐛 **报告问题** - 帮助改进项目
- 💖 **赞助作者** - 支持后续开发

<p align="center">
  <img src="https://img.shields.io/badge/微信-1638276310-brightgreen?style=for-the-badge&logo=wechat" alt="微信">
  <img src="https://img.shields.io/badge/QQ-1638276310-blue?style=for-the-badge&logo=tencentqq" alt="QQ">
</p>

---

<p align="center">
  <sub>✨ 感谢使用 BulkMailer！如果觉得不错，请给个 Star 支持一下哦~ ✨</sub>
</p>

<p align="center">
  <img src="https://komarev.com/ghpvc/?username=yourusername&repo=BulkMailer&color=blueviolet&style=flat-square" alt="访问量">
</p>
