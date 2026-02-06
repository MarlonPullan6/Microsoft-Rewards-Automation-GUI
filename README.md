# Microsoft-Rewards-Automation-GUI

* 微软Rewards每日搜索任务自动完成工具

* 支持电脑/手机任务完成

* 支持多账户同时进行任务

![App Screenshot](Assets/app.png)

### 添加账号
* 登录完成后点击头像自动完成添加

## 安装与使用

### 1. 克隆仓库

```
git clone https://github.com/<yourname>/Microsoft-Rewards-Automation-GUI.git
cd Microsoft-Rewards-Automation-GUI
```

### 2. 安装依赖

```
python -m venv .venv
.venv/Scripts/activate                 # Windows
pip install -r requirements.txt
python -m playwright install chromium  # 安装浏览器
```

### 3. 运行程序

```
python gui.py
```

------

## 技术栈

- **Python**
- **PyQt6**：图形界面
- **Playwright**：浏览器自动化
- **PyInstaller**：打包 Windows 可执行文件

------

## 声明（重要）⚠️ 

本项目旨在提供：

- Web 自动化的学习案例
- PyQt6 GUI 程序结构示例
- Playwright 与 Python 结合的实践样例

**本工具仅供学习与研究目的，不提供任何保证，亦不鼓励绕过任何服务条款。
 使用本工具可能违反 Microsoft Rewards 服务协议，请自行承担风险。**

作者不对使用者的行为造成的任何后果负责。

------

## 开源许可协议

本项目基于 **AGPL-3.0 License** 开源。
 请遵守许可证条款进行二次开发或分发。
